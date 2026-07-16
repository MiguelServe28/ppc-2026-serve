"""
Configurações — só admin. Centraliza tudo o que é configuração global da
plataforma: anos de referência, parâmetros de cálculo do PPC, datas limite,
assinatura de email e contas de email. (Os templates de email editam-se dentro
de cada módulo, junto ao envio.)
"""

from datetime import date

import streamlit as st

from common import (
    calcular_ppc,
    carregar_irs_db,
    carregar_ss_mes_db,
    data_limite_ss,
    enviar_email,
    escolher_conta_email,
    formatar_valor,
    guardar_config_db,
    meu_email,
    montar_base_irs,
    montar_base_ppc,
    montar_base_ss,
    nome_mes,
    sou_admin,
)

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
    guardar_config_db(
        p, st.session_state.templates, st.session_state.get("template_irs"),
        st.session_state.get("template_ss"),
        {"iva": st.session_state.get("template_iva"), "imi": st.session_state.get("template_imi"),
         "info": st.session_state.get("template_info")},
    )
    # O IRS trabalha por ano — se o "ano dos dados" mudou, recarrega os registos desse ano.
    st.session_state.irs_dados = carregar_irs_db(int(p["ano_dados"]))
    st.success("Configurações guardadas para toda a equipa.")

st.divider()
st.markdown("### Contas de Email")
st.caption("Também podes gerir as contas diretamente nas páginas dos impostos — isto é o mesmo widget.")
smtp_cfg_relatorio = escolher_conta_email("config")

st.divider()
st.markdown("### 📋 Relatório de Estado")
st.caption("Um resumo do estado de todos os módulos — vê-o aqui ou envia-o para o teu email (usa a conta escolhida acima).")


def gerar_resumo_estado() -> str:
    linhas = [f"RELATÓRIO DE ESTADO — Gestão Fiscal SERVE — {date.today().strftime('%d/%m/%Y')}", ""]
    hoje = date.today()

    df_ppc = calcular_ppc(montar_base_ppc(), p)
    if df_ppc.empty:
        linhas.append("PPC: sem clientes.")
    else:
        eleg = df_ppc[~df_ppc["Dispensado"]]
        linhas.append(f"PPC {p.get('ano_pagamentos')}: {len(df_ppc)} clientes ({len(eleg)} elegíveis, total a cobrar {formatar_valor(eleg['Total_PPC'].sum())} €)")
        for n in (1, 2, 3):
            env = int(eleg[f"Email{n}_Enviado"].sum()) if len(eleg) else 0
            linhas.append(f"  • {n}.º pagamento (até {p[f'data{n}'].strftime('%d/%m/%Y')}): {env}/{len(eleg)} emails enviados")

    base_irs = montar_base_irs()
    if base_irs.empty:
        linhas.append("IRS: sem clientes.")
    else:
        env = int(base_irs["Email_Enviado"].sum())
        linhas.append(f"IRS {p.get('ano_dados')}: {env}/{len(base_irs)} emails enviados · {int((base_irs['Valor_Apurado']>0).sum())} a pagar · {int((base_irs['Valor_Apurado']<0).sum())} a receber")
        pend = base_irs["Valor_Pendente"].sum()
        if pend > 0:
            linhas.append(f"  • Honorários pendentes à SERVE: {formatar_valor(pend)} €")

    base_ss = montar_base_ss()
    if base_ss.empty:
        linhas.append("Segurança Social: sem clientes.")
    else:
        ano_ref, mes_n = (hoje.year, hoje.month - 1) if hoje.month > 1 else (hoje.year - 1, 12)
        mes_ref = f"{ano_ref:04d}-{mes_n:02d}"
        env_ss = carregar_ss_mes_db(mes_ref)
        env = int(sum(env_ss.get(n, False) for n in base_ss["NIF"]))
        linhas.append(f"Segurança Social ({nome_mes(mes_ref)}): {env}/{len(base_ss)} emails enviados · pagamento até {data_limite_ss(mes_ref).strftime('%d/%m/%Y')}")

    linhas.append("")
    linhas.append("(Relatório gerado pela plataforma Gestão Fiscal SERVE.)")
    return "\n".join(linhas)


resumo = gerar_resumo_estado()
with st.expander("👁️ Ver relatório agora"):
    st.text(resumo)

if st.button("📨 Enviar relatório para o meu email"):
    if not smtp_cfg_relatorio.get("utilizador") or not smtp_cfg_relatorio.get("password"):
        st.error("Escolhe primeiro uma conta de email na secção acima.")
    else:
        try:
            enviar_email(
                smtp_cfg_relatorio, meu_email(),
                f"Relatório de Estado — Gestão Fiscal SERVE — {date.today().strftime('%d/%m/%Y')}",
                resumo, [], assinatura_html=p.get("assinatura_html", ""),
            )
            st.success(f"Relatório enviado para {meu_email()}.")
        except Exception as e:
            st.error(f"Erro ao enviar: {e}")
