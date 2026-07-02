"""
Dashboard — visão global da plataforma (todos os impostos), não específica de
nenhum imposto em particular.
"""

import streamlit as st

from common import APLICA_COLS, TIPO_COLS, clean_clientes_df, sou_admin

st.title("📊 Dashboard")
st.caption("SERVE — Contabilidade e Viabilização Empresarial. Visão global da carteira de clientes.")

clientes = clean_clientes_df(st.session_state.clientes)

if clientes.empty:
    st.info("Ainda não há clientes no registo central. Vai à página 'Clientes' para importar ou adicionar.")
    st.stop()

c1, c2 = st.columns(2)
c1.metric("Total de Clientes (registo central)", len(clientes))
c2.metric("Gestores distintos com carteira", clientes.loc[clientes["Gestor_Email"] != "", "Gestor_Email"].nunique())

st.divider()
st.markdown("### Clientes por Tipo")
tipo_labels = {
    "Tipo_Empresa": "Empresa", "Tipo_AL": "Alojamento Local",
    "Tipo_Trab_Independente": "Trabalhador Independente", "Tipo_Rep_Fiscal": "Representação Fiscal",
}
cols_tipo = st.columns(len(TIPO_COLS))
for col, campo in zip(cols_tipo, TIPO_COLS):
    col.metric(tipo_labels[campo], int(clientes[campo].sum()))

st.divider()
st.markdown("### Clientes por Imposto/Obrigação")
st.caption("Só o PPC tem página própria construída até agora — os restantes já podem ser marcados no registo central, à espera das respetivas páginas.")
aplica_labels = {
    "Aplica_PPC": "PPC", "Aplica_IVA": "IVA", "Aplica_IMI": "IMI",
    "Aplica_IRS": "IRS", "Aplica_SS": "Segurança Social",
}
cols_aplica = st.columns(len(APLICA_COLS))
for col, campo in zip(cols_aplica, APLICA_COLS):
    col.metric(aplica_labels[campo], int(clientes[campo].sum()))

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
