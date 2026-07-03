"""
Configurações — só admin. Centraliza tudo o que é configuração global da
plataforma: anos de referência, parâmetros de cálculo do PPC, datas limite,
assinatura de email e contas de email. (Os templates de email editam-se dentro
de cada módulo, junto ao envio.)
"""

import streamlit as st

from common import escolher_conta_email, guardar_config_db, sou_admin

if not sou_admin():
    st.error("Esta página é exclusiva do administrador.")
    st.stop()

st.title("⚙️ Configurações")
st.caption("Alterações aqui aplicam-se a toda a equipa. Carrega em 'Guardar configurações' no fim.")

p = st.session_state.params

st.markdown("### Anos de Referência")
c1, c2 = st.columns(2)
with c1:
    p["ano_dados"] = int(st.number_input(
        "Ano dos dados (Modelo 22 / liquidações)", value=int(p.get("ano_dados", 2025)),
        step=1, min_value=2020, max_value=2100,
        help="O ano a que se referem o volume de negócios, a coleta e as liquidações de IRS.",
    ))
with c2:
    p["ano_pagamentos"] = int(st.number_input(
        "Ano dos pagamentos", value=int(p.get("ano_pagamentos", 2026)),
        step=1, min_value=2020, max_value=2100,
        help="O ano em que os pagamentos por conta são efetuados (normalmente o ano seguinte).",
    ))
st.caption("💡 No arranque de um novo ano fiscal, basta atualizar estes dois campos e as datas abaixo — os títulos, emails e Excel adaptam-se automaticamente.")

st.divider()
st.markdown("### Parâmetros de Cálculo (PPC)")
c1, c2 = st.columns(2)
with c1:
    p["limiar_volume"] = st.number_input("Limiar Volume de Negócios (€)", value=float(p["limiar_volume"]), step=10000.0)
    p["taxa_baixa"] = st.number_input("Taxa se Volume ≤ limiar", value=float(p["taxa_baixa"]), step=0.01, format="%.2f")
with c2:
    p["taxa_alta"] = st.number_input("Taxa se Volume > limiar", value=float(p["taxa_alta"]), step=0.01, format="%.2f")
    p["limite_dispensa"] = st.number_input("Limite de dispensa (€)", value=float(p["limite_dispensa"]), step=10.0)

st.markdown("**Datas limite dos pagamentos**")
c1, c2, c3 = st.columns(3)
with c1:
    p["data1"] = st.date_input("1.º Pagamento", value=p["data1"])
with c2:
    p["data2"] = st.date_input("2.º Pagamento", value=p["data2"])
with c3:
    p["data3"] = st.date_input("3.º Pagamento", value=p["data3"])

st.divider()
st.markdown("### Assinatura de Email")
st.caption(
    "Acrescentada automaticamente no fim de todos os emails enviados pela plataforma (na versão HTML). "
    "Podes usar HTML simples: <b>negrito</b>, <br> para mudar de linha, <a href=\"...\">links</a>, etc."
)
p["assinatura_html"] = st.text_area(
    "Assinatura (HTML simples)",
    value=p.get("assinatura_html", ""),
    height=160,
    placeholder='SERVE — Contabilidade e Viabilização Empresarial<br>\nCaldas da Rainha<br>\nTel: 262 000 000 · <a href="mailto:geral@serve.pt">geral@serve.pt</a>',
)
if p["assinatura_html"]:
    with st.expander("👁️ Pré-visualizar assinatura"):
        st.markdown(p["assinatura_html"], unsafe_allow_html=True)

if st.button("💾 Guardar configurações", type="primary"):
    guardar_config_db(p, st.session_state.templates, st.session_state.get("template_irs"), st.session_state.get("template_ss"))
    st.success("Configurações guardadas para toda a equipa.")

st.divider()
st.markdown("### Contas de Email")
st.caption("Também podes gerir as contas diretamente nas páginas PPC/IRS — isto é o mesmo widget.")
escolher_conta_email("config")
