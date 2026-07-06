"""
Dashboard — visão global da plataforma: alertas de prazos, resumo de cada
módulo de imposto (PPC, IRS, Segurança Social) e retrato da carteira.
"""

from datetime import date

import altair as alt
import pandas as pd
import streamlit as st

from common import (
    APLICA_COLS,
    TIPO_COLS,
    calcular_ppc,
    carregar_ss_mes_db,
    clean_clientes_df,
    data_limite_ss,
    formatar_valor,
    montar_base_irs,
    montar_base_ppc,
    montar_base_ss,
    nome_mes,
    sou_admin,
)

st.title("📊 Dashboard")
st.caption("SERVE — Contabilidade e Viabilização Empresarial · Visão global de todos os impostos.")

clientes = clean_clientes_df(st.session_state.clientes)

if clientes.empty:
    st.info("Ainda não há clientes no registo central. Vai à página 'Clientes' para importar ou adicionar.")
    st.stop()

params = st.session_state.params
ano_pag = params.get("ano_pagamentos", 2026)
ano_dados = params.get("ano_dados", 2025)
hoje = date.today()

# --- Dados dos três módulos ---------------------------------------------------
df_ppc = calcular_ppc(montar_base_ppc(), params)
eleg_ppc = df_ppc[~df_ppc["Dispensado"]] if not df_ppc.empty else df_ppc

base_irs = montar_base_irs()

base_ss = montar_base_ss()
ano_ref, mes_ref_n = (hoje.year, hoje.month - 1) if hoje.month > 1 else (hoje.year - 1, 12)
mes_ref = f"{ano_ref:04d}-{mes_ref_n:02d}"
enviados_ss = carregar_ss_mes_db(mes_ref) if not base_ss.empty else {}
ss_enviados_n = int(sum(enviados_ss.get(n, False) for n in base_ss["NIF"])) if not base_ss.empty else 0

# --- Alertas de prazos ---------------------------------------------------------
if not df_ppc.empty:
    for n in (1, 2, 3):
        data_limite = params[f"data{n}"]
        dias = (data_limite - hoje).days
        pendentes = int((~eleg_ppc[f"Email{n}_Enviado"]).sum()) if len(eleg_ppc) else 0
        if pendentes == 0:
            continue
        if dias < 0:
            st.error(f"🚨 O prazo do {n}.º Pagamento por Conta ({data_limite.strftime('%d/%m/%Y')}) já passou há {-dias} dia(s) e ainda há {pendentes} cliente(s) sem email enviado.")
        elif dias <= 45:
            st.warning(f"⏰ O {n}.º Pagamento por Conta vence a {data_limite.strftime('%d/%m/%Y')} (daqui a {dias} dia(s)) — faltam {pendentes} email(s). Página PPC → Emails.")

if not base_ss.empty:
    pendentes_ss = len(base_ss) - ss_enviados_n
    limite_ss = data_limite_ss(mes_ref)
    dias_ss = (limite_ss - hoje).days
    if pendentes_ss and 0 <= dias_ss <= 10:
        st.warning(f"⏰ Segurança Social de {nome_mes(mes_ref)}: pagamento até {limite_ss.strftime('%d/%m/%Y')} (daqui a {dias_ss} dia(s)) — faltam {pendentes_ss} email(s). Página Segurança Social.")
    elif pendentes_ss and -5 <= dias_ss < 0:
        st.error(f"🚨 Segurança Social de {nome_mes(mes_ref)}: o prazo ({limite_ss.strftime('%d/%m/%Y')}) já passou e ainda há {pendentes_ss} email(s) por enviar.")

# --- Métricas de topo ----------------------------------------------------------
c1, c2, c3, c4 = st.columns(4)
c1.metric("Clientes no registo central", len(clientes))
c2.metric(f"PPC {ano_pag} a cobrar", (formatar_valor(eleg_ppc["Total_PPC"].sum()) + " €") if len(eleg_ppc) else "—")
c3.metric(f"IRS {ano_dados} enviados", f"{int(base_irs['Email_Enviado'].sum())} / {len(base_irs)}" if len(base_irs) else "—")
c4.metric(f"SS {nome_mes(mes_ref)}", f"{ss_enviados_n} / {len(base_ss)}" if len(base_ss) else "—")

st.divider()

# --- Resumo por módulo (cartões) ------------------------------------------------
st.markdown("### Estado dos Módulos")
col_ppc, col_irs, col_ss = st.columns(3)

with col_ppc:
    with st.container(border=True):
        st.markdown(f"#### 💶 PPC {ano_pag}")
        if df_ppc.empty:
            st.caption("Sem clientes com o pisco PPC.")
        else:
            st.caption(f"{len(df_ppc)} clientes · {int(df_ppc['Dispensado'].sum())} dispensados · {len(eleg_ppc)} elegíveis")
            for n in (1, 2, 3):
                total = len(eleg_ppc)
                env = int(eleg_ppc[f"Email{n}_Enviado"].sum()) if total else 0
                st.progress(env / total if total else 0.0,
                            text=f"{n}.º pagamento (até {params[f'data{n}'].strftime('%d/%m')}) — {env}/{total}")

with col_irs:
    with st.container(border=True):
        st.markdown(f"#### 🧾 IRS {ano_dados}")
        if base_irs.empty:
            st.caption("Sem clientes com o pisco IRS.")
        else:
            a_pagar = int((base_irs["Valor_Apurado"] > 0).sum())
            a_receber = int((base_irs["Valor_Apurado"] < 0).sum())
            avulsos = int(base_irs["IRS_Avulso"].sum())
            st.caption(f"{len(base_irs)} clientes · {a_pagar} a pagar · {a_receber} a receber · {avulsos} só IRS")
            env = int(base_irs["Email_Enviado"].sum())
            st.progress(env / len(base_irs), text=f"Emails enviados — {env}/{len(base_irs)}")
            pend_serve = base_irs["Valor_Pendente"].sum()
            if pend_serve > 0:
                st.caption(f"💰 Pendentes à SERVE: {formatar_valor(pend_serve)} €")

with col_ss:
    with st.container(border=True):
        st.markdown("#### 🏛️ Segurança Social")
        if base_ss.empty:
            st.caption("Sem clientes com o pisco Seg. Social.")
        else:
            st.caption(f"{len(base_ss)} clientes · mês de referência: {nome_mes(mes_ref)}")
            st.progress(ss_enviados_n / len(base_ss), text=f"Emails enviados — {ss_enviados_n}/{len(base_ss)}")
            st.caption(f"📅 Pagamento até {data_limite_ss(mes_ref).strftime('%d/%m/%Y')}")

st.divider()

# --- Carteira de clientes --------------------------------------------------------
st.markdown("### Carteira de Clientes")

tipo_labels = {
    "Tipo_Empresa": "Empresa", "Tipo_AL": "Alojamento Local",
    "Tipo_Trab_Independente": "Trab. Independente", "Tipo_Rep_Fiscal": "Repr. Fiscal",
}
aplica_labels = {
    "Aplica_PPC": "PPC", "Aplica_IVA": "IVA", "Aplica_IMI": "IMI",
    "Aplica_IRS": "IRS", "Aplica_SS": "Seg. Social",
}


def grafico_barras(df: pd.DataFrame, campo: str, cor: str):
    """Barras horizontais, limpas, com o valor à frente de cada barra."""
    base = alt.Chart(df).encode(
        y=alt.Y(f"{campo}:N", sort="-x", title=None),
        x=alt.X("Clientes:Q", title=None, axis=None),
    )
    barras = base.mark_bar(cornerRadiusEnd=6, height=22, color=cor)
    rotulos = base.mark_text(align="left", dx=6, color="#1a1a2e").encode(text="Clientes:Q")
    return (barras + rotulos).properties(height=alt.Step(34)).configure_view(stroke=None).configure_axis(grid=False, labelFontSize=13)


col_esq, col_dir = st.columns(2)
with col_esq:
    st.markdown("**Por tipo de cliente**")
    dados_tipo = pd.DataFrame({
        "Tipo": [tipo_labels[c] for c in TIPO_COLS],
        "Clientes": [int(clientes[c].sum()) for c in TIPO_COLS],
    })
    st.altair_chart(grafico_barras(dados_tipo, "Tipo", "#1F4E78"), use_container_width=True)
with col_dir:
    st.markdown("**Por imposto/obrigação**")
    dados_aplica = pd.DataFrame({
        "Imposto": [aplica_labels[c] for c in APLICA_COLS],
        "Clientes": [int(clientes[c].sum()) for c in APLICA_COLS],
    })
    st.altair_chart(grafico_barras(dados_aplica, "Imposto", "#2E7D32"), use_container_width=True)

notas_carteira = []
sem_imposto = int((~clientes[APLICA_COLS].any(axis=1)).sum())
if sem_imposto:
    notas_carteira.append(f"{sem_imposto} sem nenhum imposto atribuído")
en_n = int((clientes["Lingua"] == "EN").sum())
if en_n:
    notas_carteira.append(f"{en_n} com emails em inglês (EN)")
sem_email_n = int((clientes["Email"].str.strip() == "").sum())
if sem_email_n:
    notas_carteira.append(f"⚠️ {sem_email_n} sem email preenchido")
if notas_carteira:
    st.caption(" · ".join(notas_carteira))

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
