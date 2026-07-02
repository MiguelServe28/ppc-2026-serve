"""
Lógica partilhada por toda a plataforma SERVE (registo central de clientes,
ligação ao Supabase, autenticação, e utilitários usados por mais do que um
imposto). Cada página em paginas/ importa deste módulo.

Não corras este ficheiro diretamente — é só uma biblioteca. O ponto de entrada
da app é o app.py.
"""

import io
import math
import re
import smtplib
import ssl
from datetime import date
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import pandas as pd
import pdfplumber
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from supabase import Client, create_client

# ---------------------------------------------------------------------------
# Registo central de clientes (partilhado por toda a plataforma)
# ---------------------------------------------------------------------------
CLIENT_COLS = [
    "NIF", "Nome", "Email", "Gestor_Nome", "Gestor_Email",
    "Tipo_Empresa", "Tipo_AL", "Tipo_Trab_Independente", "Tipo_Rep_Fiscal",
    "Aplica_PPC", "Aplica_IVA", "Aplica_IMI", "Aplica_IRS", "Aplica_SS",
    "Notas",
]
TIPO_COLS = ["Tipo_Empresa", "Tipo_AL", "Tipo_Trab_Independente", "Tipo_Rep_Fiscal"]
APLICA_COLS = ["Aplica_PPC", "Aplica_IVA", "Aplica_IMI", "Aplica_IRS", "Aplica_SS"]
BOOL_COLS = TIPO_COLS + APLICA_COLS
TEXT_COLS = ["NIF", "Nome", "Email", "Gestor_Nome", "Gestor_Email", "Notas"]

COLUMN_MAP_TO_DB = {
    "NIF": "nif", "Nome": "nome", "Email": "email",
    "Gestor_Nome": "gestor_nome", "Gestor_Email": "gestor_email",
    "Tipo_Empresa": "tipo_empresa", "Tipo_AL": "tipo_al",
    "Tipo_Trab_Independente": "tipo_trabalhador_independente", "Tipo_Rep_Fiscal": "tipo_representacao_fiscal",
    "Aplica_PPC": "aplica_ppc", "Aplica_IVA": "aplica_iva", "Aplica_IMI": "aplica_imi",
    "Aplica_IRS": "aplica_irs", "Aplica_SS": "aplica_ss",
    "Notas": "notas",
}
COLUMN_MAP_FROM_DB = {v: k for k, v in COLUMN_MAP_TO_DB.items()}

# ---------------------------------------------------------------------------
# Dados específicos do PPC (tabela própria, ligada por NIF)
# ---------------------------------------------------------------------------
PPC_COLS = [
    "NIF", "Volume_2025", "Coleta_2025", "Retencoes_2025",
    "Guia1_Emitida", "Guia2_Emitida", "Guia3_Emitida",
    "Email1_Enviado", "Email2_Enviado", "Email3_Enviado",
]
PPC_BOOL_COLS = [c for c in PPC_COLS if c.startswith("Guia") or c.startswith("Email")]
PPC_NUM_COLS = ["Volume_2025", "Coleta_2025", "Retencoes_2025"]

PPC_COLUMN_MAP_TO_DB = {
    "NIF": "nif", "Volume_2025": "volume_2025", "Coleta_2025": "coleta_2025", "Retencoes_2025": "retencoes_2025",
    "Guia1_Emitida": "guia1_emitida", "Guia2_Emitida": "guia2_emitida", "Guia3_Emitida": "guia3_emitida",
    "Email1_Enviado": "email1_enviado", "Email2_Enviado": "email2_enviado", "Email3_Enviado": "email3_enviado",
}
PPC_COLUMN_MAP_FROM_DB = {v: k for k, v in PPC_COLUMN_MAP_TO_DB.items()}

DEFAULT_TEMPLATES = {
    1: {
        "assunto": "Pagamentos por Conta 2026 — {nome}",
        "corpo": (
            "Exmo(a). Sr(a).,\n\n"
            "No seguimento do apuramento do IRC referente a 2025, informamos que a {nome} "
            "(NIF {nif}) tem pagamentos por conta a efetuar em 2026, nos seguintes montantes e prazos:\n\n"
            "• 1.º Pagamento por Conta: {pag1} € — até {data1}\n"
            "• 2.º Pagamento por Conta: {pag2} € — até {data2}\n"
            "• 3.º Pagamento por Conta: {pag3} € — até {data3}\n\n"
            "Total anual: {total} €\n\n"
            "Segue em anexo a guia referente ao 1.º pagamento. Solicitamos que proceda ao pagamento até à data "
            "indicada, de forma a evitar juros de mora.\n\n"
            "As guias do 2.º e 3.º pagamento serão enviadas atempadamente.\n\n"
            "Ficamos ao dispor para qualquer esclarecimento.\n\n"
            "Com os melhores cumprimentos,"
        ),
    },
    2: {
        "assunto": "2.º Pagamento por Conta 2026 — {nome}",
        "corpo": (
            "Exmo(a). Sr(a).,\n\n"
            "No seguimento do 1.º pagamento por conta já efetuado, relembramos que o 2.º pagamento por conta "
            "da {nome} (NIF {nif}) vence a {data2}, no valor de {pag2} €.\n\n"
            "Segue em anexo a respetiva guia.\n\n"
            "Ficamos ao dispor para qualquer esclarecimento.\n\n"
            "Com os melhores cumprimentos,"
        ),
    },
    3: {
        "assunto": "3.º Pagamento por Conta 2026 — {nome}",
        "corpo": (
            "Exmo(a). Sr(a).,\n\n"
            "No seguimento dos pagamentos por conta já efetuados, relembramos que o 3.º e último pagamento por "
            "conta da {nome} (NIF {nif}) vence a {data3}, no valor de {pag3} €.\n\n"
            "Segue em anexo a respetiva guia.\n\n"
            "Ficamos ao dispor para qualquer esclarecimento.\n\n"
            "Com os melhores cumprimentos,"
        ),
    },
}

# ---------------------------------------------------------------------------
# Ligação ao Supabase
# ---------------------------------------------------------------------------
SUPABASE_URL = st.secrets.get("SUPABASE_URL")
SUPABASE_ANON_KEY = st.secrets.get("SUPABASE_ANON_KEY")
SUPABASE_SERVICE_KEY = st.secrets.get("SUPABASE_SERVICE_KEY")  # só necessário para o admin criar gestores


def get_client() -> Client:
    """Devolve o cliente Supabase desta sessão de browser (um por utilizador —
    nunca partilhado entre utilizadores, para que a sessão de autenticação de
    cada um não se misture com a de outro)."""
    if "sb_client" not in st.session_state:
        st.session_state.sb_client = create_client(SUPABASE_URL, SUPABASE_ANON_KEY)
    return st.session_state.sb_client


def get_admin_client() -> "Client | None":
    """Cliente com a chave 'service_role' — só usado nas ações de administração
    de contas (criar gestores). Nunca deve ser usado para ler/escrever clientes."""
    if not SUPABASE_SERVICE_KEY:
        return None
    return create_client(SUPABASE_URL, SUPABASE_SERVICE_KEY)


# ---------------------------------------------------------------------------
# Autenticação e perfil (admin vs. gestor)
# ---------------------------------------------------------------------------
def carregar_perfil():
    client = get_client()
    user_id = st.session_state.user.id
    resp = client.table("perfis").select("*").eq("id", user_id).execute()
    if resp.data:
        st.session_state.perfil = resp.data[0]
    else:
        st.session_state.perfil = {"id": user_id, "email": st.session_state.user.email, "nome": st.session_state.user.email, "role": "gestor"}


def sou_admin() -> bool:
    return "perfil" in st.session_state and st.session_state.perfil["role"] == "admin"


def meu_email() -> str:
    return st.session_state.perfil["email"]


def ecra_login():
    st.title("🔒 Gestão Fiscal SERVE")
    st.caption("SERVE — Contabilidade e Viabilização Empresarial. Inicia sessão com a tua conta de gestor.")

    if not SUPABASE_URL or not SUPABASE_ANON_KEY:
        st.error(
            "A app ainda não está ligada ao Supabase. Falta configurar SUPABASE_URL e "
            "SUPABASE_ANON_KEY em Settings → Secrets. Ver GUIA_SUPABASE.md."
        )
        st.stop()

    email = st.text_input("Email")
    pwd = st.text_input("Password", type="password")
    if st.button("Entrar", type="primary"):
        try:
            client = get_client()
            res = client.auth.sign_in_with_password({"email": email, "password": pwd})
            st.session_state.user = res.user
            carregar_perfil()
            st.rerun()
        except Exception:
            st.error("Email ou password incorretos.")


def sidebar_utilizador():
    perfil = st.session_state.perfil
    with st.sidebar:
        papel = "Administrador" if perfil["role"] == "admin" else "Gestor"
        st.success(f"👤 {perfil['nome'] or perfil['email']}  \n**{papel}**")
        if st.button("Sair"):
            try:
                get_client().auth.sign_out()
            except Exception:
                pass
            for k in ("user", "perfil", "sb_client", "clientes", "ppc_dados", "params", "templates", "log_envio"):
                st.session_state.pop(k, None)
            st.rerun()
        st.divider()


def requer_login():
    """Chama-se no topo de cada página. Se não houver sessão, mostra o login e para."""
    if "user" not in st.session_state:
        ecra_login()
        st.stop()
    sidebar_utilizador()


# ---------------------------------------------------------------------------
# Persistência — registo central de clientes (Supabase)
# ---------------------------------------------------------------------------
def clean_clientes_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for c in CLIENT_COLS:
        if c not in df.columns:
            df[c] = False if c in BOOL_COLS else ""
    for c in BOOL_COLS:
        df[c] = df[c].fillna(False).astype(bool)
    for c in TEXT_COLS:
        df[c] = df[c].fillna("").astype(str).str.strip()
    return df[CLIENT_COLS]


def clean_ppc_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for c in PPC_COLS:
        if c not in df.columns:
            df[c] = False if c in PPC_BOOL_COLS else ("" if c == "NIF" else 0.0)
    df["NIF"] = df["NIF"].fillna("").astype(str).str.strip()
    for c in PPC_NUM_COLS:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    for c in PPC_BOOL_COLS:
        df[c] = df[c].fillna(False).astype(bool)
    df = df[df["NIF"] != ""]
    return df[PPC_COLS]


def carregar_clientes_db() -> pd.DataFrame:
    client = get_client()
    resp = client.table("clientes").select("*").execute()
    rows = resp.data or []
    if not rows:
        return pd.DataFrame(columns=CLIENT_COLS)
    df = pd.DataFrame(rows).rename(columns=COLUMN_MAP_FROM_DB)
    return clean_clientes_df(df)


def perfil_nome_vazio() -> bool:
    return not (st.session_state.perfil.get("nome") or "").strip()


def guardar_clientes_db(df: pd.DataFrame):
    """Substitui os clientes visíveis por este utilizador pelo conteúdo de df.
    Um gestor só vê/apaga/insere os seus próprios clientes (RLS trata do âmbito);
    o admin substitui a lista completa."""
    client = get_client()
    df2 = clean_clientes_df(df).copy()

    if not sou_admin():
        df2["Gestor_Email"] = meu_email()
        if not perfil_nome_vazio():
            df2["Gestor_Nome"] = st.session_state.perfil["nome"]

    client.table("clientes").delete().neq("nif", "").execute()
    if not df2.empty:
        registos = df2.rename(columns=COLUMN_MAP_TO_DB).to_dict("records")
        client.table("clientes").upsert(registos, on_conflict="nif").execute()


def persistir_clientes(df: pd.DataFrame):
    """Atualiza a sessão E grava imediatamente no Supabase."""
    df = clean_clientes_df(df)
    if not sou_admin():
        df["Gestor_Email"] = meu_email()
        if not perfil_nome_vazio():
            df["Gestor_Nome"] = st.session_state.perfil["nome"]
    st.session_state.clientes = df
    guardar_clientes_db(df)


# ---------------------------------------------------------------------------
# Persistência — dados específicos do PPC
# ---------------------------------------------------------------------------
def carregar_ppc_db() -> pd.DataFrame:
    client = get_client()
    resp = client.table("ppc_dados").select("*").execute()
    rows = resp.data or []
    if not rows:
        return pd.DataFrame(columns=PPC_COLS)
    df = pd.DataFrame(rows).rename(columns=PPC_COLUMN_MAP_FROM_DB)
    return clean_ppc_df(df)


def guardar_ppc_db(df: pd.DataFrame):
    client = get_client()
    df2 = clean_ppc_df(df).copy()
    client.table("ppc_dados").delete().neq("nif", "").execute()
    if not df2.empty:
        registos = df2.rename(columns=PPC_COLUMN_MAP_TO_DB).to_dict("records")
        client.table("ppc_dados").upsert(registos, on_conflict="nif").execute()


def persistir_ppc(df: pd.DataFrame):
    df = clean_ppc_df(df)
    st.session_state.ppc_dados = df
    guardar_ppc_db(df)


def montar_base_ppc() -> pd.DataFrame:
    """Junta os clientes com 'Aplica_PPC' ligado com os respetivos dados de PPC
    (mesmo que ainda não tenham nenhum valor preenchido)."""
    clientes = clean_clientes_df(st.session_state.clientes)
    elegiveis = clientes[clientes["Aplica_PPC"]].copy()
    ppc = clean_ppc_df(st.session_state.ppc_dados)
    base = elegiveis.merge(ppc, on="NIF", how="left")
    for c in PPC_NUM_COLS:
        base[c] = pd.to_numeric(base[c], errors="coerce").fillna(0.0)
    for c in PPC_BOOL_COLS:
        base[c] = base[c].fillna(False).astype(bool)
    return base


# ---------------------------------------------------------------------------
# Persistência — log de envios e configuração
# ---------------------------------------------------------------------------
def carregar_log_db() -> list:
    client = get_client()
    resp = client.table("log_envios").select("data, nif, nome, pagamento, estado").order("id").execute()
    return resp.data or []


def guardar_log_entry_db(entry: dict):
    client = get_client()
    client.table("log_envios").insert(entry).execute()


def registar_log(entry: dict):
    st.session_state.log_envio.append(entry)
    guardar_log_entry_db(entry)


def carregar_config_db():
    client = get_client()
    resp = client.table("config").select("params_json, templates_json, templates_irs_json").eq("id", 1).execute()
    if not resp.data:
        return None, None, None
    row = resp.data[0]
    params_json, templates_json = row.get("params_json"), row.get("templates_json")
    templates_irs_json = row.get("templates_irs_json")
    params, templates, template_irs = None, None, None
    if params_json:
        params = {
            "limiar_volume": params_json["limiar_volume"], "taxa_baixa": params_json["taxa_baixa"],
            "taxa_alta": params_json["taxa_alta"], "limite_dispensa": params_json["limite_dispensa"],
            "data1": date.fromisoformat(params_json["data1"]), "data2": date.fromisoformat(params_json["data2"]),
            "data3": date.fromisoformat(params_json["data3"]),
        }
    if templates_json:
        templates = {int(k): v for k, v in templates_json.items()}
    if templates_irs_json:
        template_irs = templates_irs_json
    return params, templates, template_irs


def guardar_config_db(params: dict, templates: dict, template_irs: dict = None):
    """Só o admin tem permissão (RLS) para escrever a configuração global."""
    if not sou_admin():
        return
    params_serializ = dict(params)
    for k in ("data1", "data2", "data3"):
        params_serializ[k] = params[k].isoformat()
    templates_serializ = {str(k): v for k, v in templates.items()}
    registo = {"id": 1, "params_json": params_serializ, "templates_json": templates_serializ}
    if template_irs is not None:
        registo["templates_irs_json"] = template_irs
    client = get_client()
    client.table("config").upsert(registo).execute()


# ---------------------------------------------------------------------------
# Estado (carregado uma vez por sessão)
# ---------------------------------------------------------------------------
def init_state():
    if "perfil" not in st.session_state:
        carregar_perfil()
    if "clientes" not in st.session_state:
        st.session_state.clientes = carregar_clientes_db()
    if "ppc_dados" not in st.session_state:
        st.session_state.ppc_dados = carregar_ppc_db()
    if "irs_dados" not in st.session_state:
        st.session_state.irs_dados = carregar_irs_db()
    if "guias" not in st.session_state:
        st.session_state.guias = {}  # {(nif, n_pagamento): (filename, bytes)} — só dura a sessão atual
    if "guias_irs" not in st.session_state:
        st.session_state.guias_irs = {}  # {nif: (filename, bytes)} — guia de pagamento de IRS, só dura a sessão
    if "faturas_irs" not in st.session_state:
        st.session_state.faturas_irs = {}  # {nif: (filename, bytes)} — fatura do serviço de IRS, só dura a sessão
    if "params" not in st.session_state or "templates" not in st.session_state or "template_irs" not in st.session_state:
        params_db, templates_db, template_irs_db = carregar_config_db()
        if "params" not in st.session_state:
            st.session_state.params = params_db or {
                "limiar_volume": 500000.0,
                "taxa_baixa": 0.80,
                "taxa_alta": 0.95,
                "limite_dispensa": 200.0,
                "data1": date(2026, 7, 31),
                "data2": date(2026, 9, 30),
                "data3": date(2026, 12, 15),
            }
        if "templates" not in st.session_state:
            st.session_state.templates = templates_db or {k: v.copy() for k, v in DEFAULT_TEMPLATES.items()}
        if "template_irs" not in st.session_state:
            st.session_state.template_irs = template_irs_db or DEFAULT_TEMPLATE_IRS.copy()
    if "log_envio" not in st.session_state:
        st.session_state.log_envio = carregar_log_db()


# ---------------------------------------------------------------------------
# Importação de ficheiros (CSV/Excel)
# ---------------------------------------------------------------------------
def parse_numero_pt(serie: pd.Series) -> pd.Series:
    """Converte números em formato português (vírgula decimal, ponto de milhares) ou internacional para float."""
    def conv(v):
        if pd.isna(v):
            return None
        s = str(v).strip()
        if s == "":
            return None
        s = s.replace(" ", "").replace("€", "")
        if "," in s and "." in s:
            s = s.replace(".", "").replace(",", ".")
        elif "," in s:
            s = s.replace(",", ".")
        try:
            return float(s)
        except ValueError:
            return None
    return serie.apply(conv)


def ler_ficheiro_importacao(uploaded_file) -> pd.DataFrame:
    """Lê CSV ou Excel de forma tolerante: deteta o encoding do CSV automaticamente
    (ficheiros exportados do Excel em português costumam vir em Windows-1252/ISO-8859-1,
    não UTF-8) e trata números com vírgula decimal."""
    nome = uploaded_file.name.lower()
    if nome.endswith(".csv"):
        raw = uploaded_file.read()
        texto = None
        for enc in ("utf-8-sig", "utf-8", "cp1252", "iso-8859-1"):
            try:
                texto = raw.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        if texto is None:
            texto = raw.decode("latin1", errors="replace")
        df = pd.read_csv(io.StringIO(texto), sep=None, engine="python")
    else:
        df = pd.read_excel(uploaded_file)

    df.columns = [str(c).strip() for c in df.columns]
    for c in ("Volume_2025", "Coleta_2025", "Retencoes_2025"):
        if c in df.columns:
            df[c] = parse_numero_pt(df[c])
    if "Email" in df.columns:
        df["Email"] = df["Email"].astype(str).str.strip().str.rstrip(",").str.strip()
        df.loc[df["Email"] == "nan", "Email"] = ""
    if "NIF" in df.columns:
        df["NIF"] = df["NIF"].astype(str).str.strip()
    return df


# ---------------------------------------------------------------------------
# Cálculo PPC
# ---------------------------------------------------------------------------
def calcular_ppc(df: pd.DataFrame, params: dict) -> pd.DataFrame:
    df = df.copy()
    df["Base_Calculo"] = (df["Coleta_2025"] - df["Retencoes_2025"]).clip(lower=0)
    df["Taxa"] = df["Volume_2025"].apply(
        lambda v: params["taxa_baixa"] if v <= params["limiar_volume"] else params["taxa_alta"]
    )
    df["Dispensado"] = df["Base_Calculo"] < params["limite_dispensa"]
    df["Total_PPC_Base"] = df["Base_Calculo"] * df["Taxa"]

    def parcela(row):
        if row["Dispensado"] or row["Total_PPC_Base"] <= 0:
            return 0
        return math.ceil(round(row["Total_PPC_Base"] / 3, 6) - 1e-9)

    df["Pag1"] = df.apply(parcela, axis=1)
    df["Pag2"] = df["Pag1"]
    df["Pag3"] = df["Pag1"]
    df["Total_PPC"] = df["Pag1"] + df["Pag2"] + df["Pag3"]
    return df


# ---------------------------------------------------------------------------
# Export Excel (folha de controlo do PPC)
# ---------------------------------------------------------------------------
def gerar_excel_ppc(df_calc: pd.DataFrame, params: dict) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Controlo PPC 2026"
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A5"

    FONT = "Arial"
    HEADER_FILL = PatternFill("solid", start_color="1F4E78", end_color="1F4E78")
    DISPENSA_FILL = PatternFill("solid", start_color="E2EFDA", end_color="E2EFDA")
    HEADER_FONT = Font(name=FONT, color="FFFFFF", bold=True, size=10)
    TITLE_FONT = Font(name=FONT, bold=True, size=14, color="1F4E78")
    BLACK = Font(name=FONT, color="000000")
    thin = Side(style="thin", color="BFBFBF")
    BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws["A1"] = "Controlo de Pagamentos por Conta — 2026"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:D1")

    headers = [
        "NIF", "Nome", "Email", "Gestor (nome)", "Gestor (email)", "Volume 2025", "Coleta 2025", "Retenções 2025",
        "Base de Cálculo", "Taxa", "Total PPC", "Dispensado",
        "1º Pagamento", "2º Pagamento", "3º Pagamento",
        "Data Limite 1º", "Data Limite 2º", "Data Limite 3º",
        "Guia1 Emitida", "Guia2 Emitida", "Guia3 Emitida",
        "Email1 Enviado", "Email2 Enviado", "Email3 Enviado", "Notas",
    ]
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=4, column=i, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
    ws.row_dimensions[4].height = 30

    widths = [12, 26, 24, 18, 24, 13, 13, 13, 13, 8, 12, 11, 11, 11, 11, 12, 12, 12, 11, 11, 11, 11, 11, 11, 20]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    row = 5
    for _, r in df_calc.iterrows():
        vals = [
            r["NIF"], r["Nome"], r["Email"], r["Gestor_Nome"], r["Gestor_Email"],
            r["Volume_2025"], r["Coleta_2025"], r["Retencoes_2025"],
            r["Base_Calculo"], r["Taxa"], r["Total_PPC"], "Sim" if r["Dispensado"] else "Não",
            r["Pag1"], r["Pag2"], r["Pag3"],
            params["data1"] if not r["Dispensado"] else None,
            params["data2"] if not r["Dispensado"] else None,
            params["data3"] if not r["Dispensado"] else None,
            "Sim" if r["Guia1_Emitida"] else "Não",
            "Sim" if r["Guia2_Emitida"] else "Não",
            "Sim" if r["Guia3_Emitida"] else "Não",
            "Sim" if r["Email1_Enviado"] else "Não",
            "Sim" if r["Email2_Enviado"] else "Não",
            "Sim" if r["Email3_Enviado"] else "Não",
            r["Notas"],
        ]
        for i, v in enumerate(vals, start=1):
            c = ws.cell(row=row, column=i, value=v)
            c.font = BLACK
            c.border = BORDER
            if i in (6, 7, 8, 9, 11, 13, 14, 15):
                c.number_format = "#,##0.00"
            if i == 10:
                c.number_format = "0%"
            if i in (16, 17, 18):
                c.number_format = "dd/mm/yyyy"
        if r["Dispensado"]:
            for i in range(1, len(vals) + 1):
                ws.cell(row=row, column=i).fill = DISPENSA_FILL
        row += 1

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Emails
# ---------------------------------------------------------------------------
def formatar_valor(v):
    return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def render_template(template: dict, row: pd.Series, params: dict) -> tuple[str, str]:
    ctx = {
        "nome": row["Nome"],
        "nif": row["NIF"],
        "email": row["Email"],
        "pag1": formatar_valor(row["Pag1"]),
        "pag2": formatar_valor(row["Pag2"]),
        "pag3": formatar_valor(row["Pag3"]),
        "total": formatar_valor(row["Total_PPC"]),
        "data1": params["data1"].strftime("%d/%m/%Y"),
        "data2": params["data2"].strftime("%d/%m/%Y"),
        "data3": params["data3"].strftime("%d/%m/%Y"),
    }
    assunto = template["assunto"].format(**ctx)
    corpo = template["corpo"].format(**ctx)
    return assunto, corpo


def enviar_email(smtp_cfg, destinatario, assunto, corpo, anexos, cc=None):
    cc_list = [e.strip() for e in (cc or []) if e and e.strip()]

    msg = MIMEMultipart()
    msg["From"] = smtp_cfg["remetente"]
    msg["To"] = destinatario
    if cc_list:
        msg["Cc"] = ", ".join(cc_list)
    msg["Subject"] = assunto
    msg.attach(MIMEText(corpo, "plain", "utf-8"))
    for filename, filebytes in anexos:
        part = MIMEApplication(filebytes, Name=filename)
        part["Content-Disposition"] = f'attachment; filename="{filename}"'
        msg.attach(part)

    todos_destinatarios = [destinatario] + cc_list

    context = ssl.create_default_context()
    if smtp_cfg["tls"]:
        with smtplib.SMTP(smtp_cfg["host"], smtp_cfg["porta"], timeout=30) as server:
            server.starttls(context=context)
            server.login(smtp_cfg["utilizador"], smtp_cfg["password"])
            server.sendmail(smtp_cfg["remetente"], todos_destinatarios, msg.as_string())
    else:
        with smtplib.SMTP_SSL(smtp_cfg["host"], smtp_cfg["porta"], context=context, timeout=30) as server:
            server.login(smtp_cfg["utilizador"], smtp_cfg["password"])
            server.sendmail(smtp_cfg["remetente"], todos_destinatarios, msg.as_string())


def extrair_nif_de_filename(filename: str):
    m = re.search(r"\b(\d{9})\b", filename)
    return m.group(1) if m else None


# ---------------------------------------------------------------------------
# Configuração SMTP partilhada (evita ter de reintroduzir as credenciais em
# cada página que envia email, dentro da mesma sessão)
# ---------------------------------------------------------------------------
def smtp_config_form() -> dict:
    if "smtp_cfg" not in st.session_state:
        st.session_state.smtp_cfg = {
            "host": "smtp.office365.com", "porta": 587, "tls": True,
            "utilizador": "", "password": "", "remetente": "",
        }
    cfg = st.session_state.smtp_cfg
    st.markdown("### Configuração SMTP")
    c1, c2 = st.columns(2)
    with c1:
        cfg["host"] = st.text_input("Servidor SMTP", value=cfg["host"], key="smtp_host")
        cfg["utilizador"] = st.text_input("Utilizador (email de login)", value=cfg["utilizador"], key="smtp_user")
        cfg["remetente"] = st.text_input("Remetente (From)", value=cfg["utilizador"] or cfg["remetente"], key="smtp_from")
    with c2:
        cfg["porta"] = st.number_input("Porta", value=int(cfg["porta"]), step=1, key="smtp_port")
        cfg["tls"] = st.checkbox("Usar STARTTLS (recomendado, porta 587)", value=cfg["tls"], key="smtp_tls")
        cfg["password"] = st.text_input("Password / App Password", value=cfg["password"], type="password", key="smtp_pass")
    st.caption("Gmail: smtp.gmail.com, porta 587, TLS — requer 'App Password'. Office365/Outlook: smtp.office365.com, porta 587, TLS. A password não é guardada na base de dados — fica só nesta sessão do browser.")
    return {"host": cfg["host"], "porta": int(cfg["porta"]), "tls": cfg["tls"],
            "utilizador": cfg["utilizador"], "password": cfg["password"], "remetente": cfg["remetente"]}


# ---------------------------------------------------------------------------
# IRS — dados específicos (tabela própria, ligada por NIF) e leitura de PDFs
# ---------------------------------------------------------------------------
IRS_COLS = ["NIF", "Numero_Liquidacao", "Valor_Apurado", "Valor_Pendente", "Incluido_Avenca", "Email_Enviado"]
IRS_NUM_COLS = ["Valor_Apurado", "Valor_Pendente"]
IRS_BOOL_COLS = ["Incluido_Avenca", "Email_Enviado"]

IRS_COLUMN_MAP_TO_DB = {
    "NIF": "nif", "Numero_Liquidacao": "numero_liquidacao",
    "Valor_Apurado": "valor_apurado", "Valor_Pendente": "valor_pendente",
    "Incluido_Avenca": "incluido_avenca", "Email_Enviado": "email_enviado",
}
IRS_COLUMN_MAP_FROM_DB = {v: k for k, v in IRS_COLUMN_MAP_TO_DB.items()}

DEFAULT_TEMPLATE_IRS = {
    "assunto": "Liquidação de IRS — {nome}",
    "corpo": (
        "Exmo(a). Sr(a).,\n\n"
        "Junto enviamos a Demonstração de Liquidação de IRS referente ao ano de 2025 "
        "(NIF {nif}{ref_liquidacao}).\n\n"
        "{frase_valor}\n\n"
        "{frase_pendente}"
        "Segue também em anexo a guia de pagamento, quando aplicável.\n\n"
        "Ficamos ao dispor para qualquer esclarecimento.\n\n"
        "Com os melhores cumprimentos,"
    ),
}


def clean_irs_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for c in IRS_COLS:
        if c not in df.columns:
            if c in IRS_BOOL_COLS:
                df[c] = False
            elif c in IRS_NUM_COLS:
                df[c] = 0.0
            else:
                df[c] = ""
    df["NIF"] = df["NIF"].fillna("").astype(str).str.strip()
    df["Numero_Liquidacao"] = df["Numero_Liquidacao"].fillna("").astype(str).str.strip()
    for c in IRS_NUM_COLS:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    for c in IRS_BOOL_COLS:
        df[c] = df[c].fillna(False).astype(bool)
    df = df[df["NIF"] != ""]
    return df[IRS_COLS]


def carregar_irs_db() -> pd.DataFrame:
    client = get_client()
    resp = client.table("irs_dados").select("*").execute()
    rows = resp.data or []
    if not rows:
        return pd.DataFrame(columns=IRS_COLS)
    df = pd.DataFrame(rows).rename(columns=IRS_COLUMN_MAP_FROM_DB)
    return clean_irs_df(df)


def guardar_irs_db(df: pd.DataFrame):
    client = get_client()
    df2 = clean_irs_df(df).copy()
    client.table("irs_dados").delete().neq("nif", "").execute()
    if not df2.empty:
        registos = df2.rename(columns=IRS_COLUMN_MAP_TO_DB).to_dict("records")
        client.table("irs_dados").upsert(registos, on_conflict="nif").execute()


def persistir_irs(df: pd.DataFrame):
    df = clean_irs_df(df)
    st.session_state.irs_dados = df
    guardar_irs_db(df)


def montar_base_irs() -> pd.DataFrame:
    """Junta os clientes com 'Aplica_IRS' ligado com os respetivos dados de IRS
    (mesmo que ainda não tenham nenhum valor preenchido)."""
    clientes = clean_clientes_df(st.session_state.clientes)
    elegiveis = clientes[clientes["Aplica_IRS"]].copy()
    irs = clean_irs_df(st.session_state.get("irs_dados", pd.DataFrame(columns=IRS_COLS)))
    base = elegiveis.merge(irs, on="NIF", how="left")
    for c in IRS_NUM_COLS:
        base[c] = pd.to_numeric(base[c], errors="coerce").fillna(0.0)
    for c in IRS_BOOL_COLS:
        base[c] = base[c].fillna(False).astype(bool)
    base["Numero_Liquidacao"] = base["Numero_Liquidacao"].fillna("")
    return base


def _parse_valor_pt(texto: str):
    """Converte um número em formato português (ex: '-1.234,56') encontrado em
    texto extraído de PDF para float. Devolve None se não conseguir converter."""
    if texto is None:
        return None
    s = texto.strip().replace(" ", "").replace("€", "")
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def _extrair_texto_pdf(ficheiro_bytes: bytes) -> str:
    texto_paginas = []
    with pdfplumber.open(io.BytesIO(ficheiro_bytes)) as pdf:
        for pagina in pdf.pages:
            texto_paginas.append(pagina.extract_text() or "")
    return "\n".join(texto_paginas)


def extrair_dados_liquidacao_irs(ficheiro_bytes: bytes, nif_esperado: str = None) -> dict:
    """Lê uma 'Demonstração de Liquidação de IRS' da Autoridade Tributária e tenta
    extrair o número de liquidação e o valor final (a pagar, a receber, ou
    apurado quando é zero). Se 'nif_esperado' for indicado, verifica apenas se
    esse NIF aparece em algum lado do documento — não tenta identificar/listar
    todos os NIFs do documento (podem ser um ou dois, em declarações conjuntas,
    e não interessa a posição, só se o do cliente lá está). Devolve sempre um
    dicionário com as chaves — valores a None se não encontrar, para a página
    poder avisar e pedir preenchimento manual em vez de adivinhar."""
    resultado = {
        "nif_confirmado": None, "numero_liquidacao": None, "periodo": None,
        "valor_apurado": None, "tipo_valor": None, "texto_bruto": "",
    }
    try:
        texto = _extrair_texto_pdf(ficheiro_bytes)
    except Exception:
        return resultado
    resultado["texto_bruto"] = texto

    if nif_esperado:
        resultado["nif_confirmado"] = bool(re.search(rf"\b{re.escape(nif_esperado)}\b", texto))

    # Número de liquidação (formato "AAAA.dígitos") seguido do período de
    # rendimentos — não precisamos de capturar o(s) NIF(s) que vêm antes.
    m_linha = re.search(
        r"(?P<numero>\d{4}\.\d+)\s+(?P<periodo_ini>\d{4}-\d{2}-\d{2})\s+a\s+(?P<periodo_fim>\d{4}-\d{2}-\d{2})",
        texto,
    )
    if m_linha:
        resultado["numero_liquidacao"] = m_linha.group("numero")
        resultado["periodo"] = f"{m_linha.group('periodo_ini')} a {m_linha.group('periodo_fim')}"

    # O valor final vem sempre rotulado — "Valor a pagar", "Valor a receber" ou
    # "Valor apurado" (quando não há nada a pagar nem a receber). O próprio
    # rótulo diz-nos o sinal, por isso não precisamos de adivinhar a partir do
    # número (o número no PDF vem sempre positivo, mesmo quando é um reembolso).
    m_valor = re.search(r"Valor (a pagar|a receber|apurado)\s+(-?[\d\.]+,\d{2})", texto)
    if m_valor:
        rotulo = m_valor.group(1)
        numero = _parse_valor_pt(m_valor.group(2))
        if numero is not None:
            if rotulo == "a receber":
                numero = -abs(numero)
            elif rotulo == "a pagar":
                numero = abs(numero)
        resultado["valor_apurado"] = numero
        resultado["tipo_valor"] = rotulo

    return resultado


def extrair_dados_pendentes_irs(ficheiro_bytes: bytes) -> dict:
    """Lê um 'Controlo de Pendentes' (extrato de faturas em dívida à SERVE, gerado
    pelo TOConline) e tenta extrair o NIF do cliente e o total pendente."""
    resultado = {"nif": None, "valor_pendente": None, "texto_bruto": ""}
    try:
        texto = _extrair_texto_pdf(ficheiro_bytes)
    except Exception:
        return resultado
    resultado["texto_bruto"] = texto

    m_nif = re.search(r"EUR\s+(\d{9})", texto)
    if m_nif:
        resultado["nif"] = m_nif.group(1)

    m_total = re.search(r"TOTAL PENDENTE\s+([\d\.,]+)", texto)
    if m_total:
        resultado["valor_pendente"] = _parse_valor_pt(m_total.group(1))

    return resultado


def render_template_irs(template: dict, row: pd.Series) -> tuple[str, str]:
    valor = row.get("Valor_Apurado", 0.0) or 0.0
    if valor > 0:
        frase_valor = f"Do apuramento efetuado, resulta um valor a pagar de {formatar_valor(valor)} €."
    elif valor < 0:
        frase_valor = f"Do apuramento efetuado, resulta um valor a receber (reembolso) de {formatar_valor(abs(valor))} €."
    else:
        frase_valor = "Do apuramento efetuado, não resulta qualquer valor a pagar ou a receber."

    pendente = row.get("Valor_Pendente", 0.0) or 0.0
    if pendente > 0:
        frase_pendente = (
            f"Informamos ainda que, de acordo com os nossos registos, tem pendente o valor de "
            f"{formatar_valor(pendente)} € referente a honorários em dívida à SERVE.\n\n"
        )
    else:
        frase_pendente = ""

    ref_liquidacao = f", n.º de liquidação {row['Numero_Liquidacao']}" if row.get("Numero_Liquidacao") else ""

    ctx = {
        "nome": row["Nome"],
        "nif": row["NIF"],
        "email": row["Email"],
        "ref_liquidacao": ref_liquidacao,
        "frase_valor": frase_valor,
        "frase_pendente": frase_pendente,
    }
    assunto = template["assunto"].format(**ctx)
    corpo = template["corpo"].format(**ctx)
    return assunto, corpo
