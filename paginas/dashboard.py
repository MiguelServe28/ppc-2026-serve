"""
Dashboard — visão global da plataforma (todos os impostos), com alertas de
prazos que se aproximam e gráficos da carteira de clientes.
"""

from datetime import date

import pandas as pd
import streamlit as st

from common import (
    APLICA_COLS,
    TIPO_COLS,
    calcular_ppc,
    carregar_ss_mes_db,
    clean_clientes_df,
    data_limite_ss,
    montar_base_ppc,
    montar_base_ss,
    nome_mes,
    sou_admin,
)

st.title("📊 Dashboard")
st.caption("SERVE — Contabilidade e Viabilização Empresarial. Visão global da carteira de clientes.")

clientes = clean_clientes_df(st.session_state.clientes)

if clientes.empty:
    st.info("Ainda não há clientes no registo central. Vai à página 'Clientes' para importar ou adicionar.")
    st.stop()

# --- Alertas de prazos (PPC) -------------------------------------------------
params = st.session_state.params
df_ppc = calcular_ppc(montar_base_ppc(), params)
if not df_ppc.empty:
    elegiveis = df_ppc[~df_ppc["Dispensado"]]
    hoje = date.today()
    for n in (1, 2, 3):
        data_limite = params[f"data{n}"]
        dias = (data_limite - hoje).days
        pendentes = int((~elegiveis[f"Email{n}_Enviado"]).sum()) if len(elegiveis) else 0
        if pendentes == 0:
            continue
        if dias < 0:
            st.error(f"🚨 O prazo do {n}.º Pagamento por Conta ({data_limite.strftime('%d/%m/%Y')}) já passou há {-dias} dia(s) e ainda há {pendentes} cliente(s) sem email enviado.")
        elif dias <= 45:
            st.warning(f"⏰ O {n}.º Pagamento por Conta vence a {data_limite.strftime('%d/%m/%Y')} (daqui a {dias} dia(s)) — faltam {pendentes} email(s) por enviar. Vai à página PPC → Emails.")

# --- Alerta Segurança Social (mês de referência = mês anterior) --------------
base_ss = montar_base_ss()
if not base_ss.empty:
    hoje = date.today()
    ano_ref, mes_ref_n = (hoje.year, hoje.month - 1) if hoje.month > 1 else (hoje.year - 1, 12)
    mes_ref = f"{ano_ref:04d}-{mes_ref_n:02d}"
    enviados_ss = carregar_ss_mes_db(mes_ref)
    pendentes_ss = int(sum(not enviados_ss.get(n, False) for n in base_ss["NIF"]))
    limite_ss = data_limite_ss(mes_ref)
    dias_ss = (limite_ss - hoje).days
    if pendentes_ss and 0 <= dias_ss <= 10:
        st.warning(f"⏰ Segurança Social de {nome_mes(mes_ref)}: pagamento até {limite_ss.strftime('%d/%m/%Y')} (daqui a {dias_ss} dia(s)) — faltam {pendentes_ss} email(s) por enviar. Vai à página Segurança Social.")
    elif pendentes_ss and -5 <= dias_ss < 0:
        st.error(f"🚨 Segurança Social de {nome_mes(mes_ref)}: o prazo ({limite_ss.strftime('%d/%m/%Y')}) já passou e ainda há {pendentes_ss} email(s) por enviar.")

c1, c2 = st.columns(2)
c1.metric("Total de Clientes (registo central)", len(clientes))
c2.metric("Gestores distintos com carteira", clientes.loc[clientes["Gestor_Email"] != "", "Gestor_Email"].nunique())

st.divider()
tipo_labels = {
    "Tipo_Empresa": "Empresa", "Tipo_AL": "Alojamento Local",
    "Tipo_Trab_Independente": "Trabalhador Independente", "Tipo_Rep_Fiscal": "Representação Fiscal",
}
aplica_labels = {
    "Aplica_PPC": "PPC", "Aplica_IVA": "IVA", "Aplica_IMI": "IMI",
    "Aplica_IRS": "IRS", "Aplica_SS": "Segurança Social",
}

col_esq, col_dir = st.columns(2)
with col_esq:
    st.markdown("### Clientes por Tipo")
    dados_tipo = pd.DataFrame({
        "Tipo": [tipo_labels[c] for c in TIPO_COLS],
        "Clientes": [int(clientes[c].sum()) for c in TIPO_COLS],
    }).set_index("Tipo")
    st.bar_chart(dados_tipo, color="#1F4E78")
with col_dir:
    st.markdown("### Clientes por Imposto/Obrigação")
    dados_aplica = pd.DataFrame({
        "Imposto": [aplica_labels[c] for c in APLICA_COLS],
        "Clientes": [int(clientes[c].sum()) for c in APLICA_COLS],
    }).set_index("Imposto")
    st.bar_chart(dados_aplica, color="#2E7D32")

sem_imposto = int((~clientes[APLICA_COLS].any(axis=1)).sum())
if sem_imposto:
    st.caption(f"ℹ️ {sem_imposto} cliente(s) sem nenhum imposto atribuído — vê-os na página 'Clientes' com o filtro 'Sem nenhum imposto atribuído'.")

# --- Progresso dos módulos ---------------------------------------------------
st.divider()
st.markdown("### Progresso de Envios")
col_a, col_b = st.columns(2)
with col_a:
    st.markdown("**PPC**")
    if df_ppc.empty:
        st.caption("Sem clientes PPC ainda.")
    else:
        eleg = df_ppc[~df_ppc["Dispensado"]]
        for n in (1, 2, 3):
            total = len(eleg)
            enviados = int(eleg[f"Email{n}_Enviado"].sum()) if total else 0
            st.progress(enviados / total if total else 0.0, text=f"{n}.º Pagamento — {enviados}/{total} emails enviados")
with col_b:
    st.markdown("**IRS**")
    from common import montar_base_irs
    base_irs = montar_base_irs()
    if base_irs.empty:
        st.caption("Sem clientes IRS ainda.")
    else:
        enviados = int(base_irs["Email_Enviado"].sum())
        st.progress(enviados / len(base_irs), text=f"Liquidações — {enviados}/{len(base_irs)} emails enviados")

if sou_admin():
    st.divider()
    st.markdown("### Clientes por Gestor")
    resumo_gestor = (
        clientes.assign(Gestor=clientes["Gestor_Email"].replace("", "— sem gestor —"))
        .groupby("Gestor")
        .size()
        .reset_index(name="Nº Clientes")
        .sort_values("Nº Clientes", ascending=False)
    )
    st.dataframe(resumo_gestor, use_container_width=True, hide_index=True)
