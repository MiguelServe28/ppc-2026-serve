"""
Página de IRS — tudo o que é específico deste imposto vive aqui: dados de
liquidação por cliente (lidos automaticamente dos PDFs sempre que possível,
mas sempre confirmáveis/editáveis antes de gravar ou enviar), upload de guia,
nota de liquidação e controlo de pendentes, e envio do respetivo email.
"""

from datetime import datetime

import pandas as pd
import streamlit as st

from common import (
    IRS_COLS,
    enviar_email,
    extrair_dados_liquidacao_irs,
    extrair_dados_pendentes_irs,
    guardar_config_db,
    montar_base_irs,
    persistir_irs,
    registar_log,
    render_template_irs,
    smtp_config_form,
    sou_admin,
)

st.title("🧾 IRS — Liquidações")
st.caption("SERVE — Contabilidade e Viabilização Empresarial")
st.caption(
    "Cada cliente é tratado individualmente: seleciona o cliente, carrega os 3 documentos "
    "(Guia, Nota de Liquidação e, se aplicável, o Controlo de Pendentes), confirma os valores "
    "lidos automaticamente e envia o email."
)

base_irs = montar_base_irs()

if base_irs.empty:
    st.info("Ainda não há clientes com 'Aplica IRS' ligado. Ativa esse interruptor na página 'Clientes'.")
    st.stop()

tab_visao, tab_processar, tab_template = st.tabs(["📊 Visão Geral", "📎 Processar Cliente", "✏️ Template de Email"])

# --- Visão Geral -----------------------------------------------------------
with tab_visao:
    st.subheader("Estado por Cliente")
    show_cols = ["NIF", "Nome", "Numero_Liquidacao", "Valor_Apurado", "Valor_Pendente", "Incluido_Avenca", "Email_Enviado"]
    st.dataframe(
        base_irs[show_cols].rename(columns={
            "Numero_Liquidacao": "Nº Liquidação", "Valor_Apurado": "Valor Apurado (€)",
            "Valor_Pendente": "Pendente (€)", "Incluido_Avenca": "Incluído na Avença", "Email_Enviado": "Email Enviado",
        }),
        use_container_width=True,
        hide_index=True,
        height=400,
    )
    c1, c2, c3 = st.columns(3)
    c1.metric("Total de Clientes IRS", len(base_irs))
    c2.metric("A Pagar", int((base_irs["Valor_Apurado"] > 0).sum()))
    c3.metric("A Receber (reembolso)", int((base_irs["Valor_Apurado"] < 0).sum()))

# --- Processar Cliente -------------------------------------------------
with tab_processar:
    st.subheader("Selecionar Cliente")
    nif_escolhido = st.selectbox(
        "Cliente",
        base_irs["NIF"].tolist(),
        format_func=lambda n: f"{n} — {base_irs.loc[base_irs['NIF']==n,'Nome'].values[0]}",
        key="irs_cliente_escolhido",
    )
    row_atual = base_irs[base_irs["NIF"] == nif_escolhido].iloc[0]

    st.divider()
    incluido_avenca = st.checkbox(
        "Serviço de IRS incluído na avença deste cliente (não é faturado à parte — não mostra upload de fatura)",
        value=bool(row_atual.get("Incluido_Avenca", False)),
        key=f"incluido_avenca_{nif_escolhido}",
    )

    st.markdown("### 1. Carregar Documentos")
    colunas_upload = st.columns(3 if incluido_avenca else 4)
    col_g, col_l, col_p = colunas_upload[0], colunas_upload[1], colunas_upload[2]
    col_f = colunas_upload[3] if not incluido_avenca else None

    with col_g:
        up_guia = st.file_uploader("Guia de Pagamento (PDF)", type=["pdf"], key=f"up_guia_irs_{nif_escolhido}")
        if up_guia is not None:
            st.session_state.guias_irs[nif_escolhido] = (up_guia.name, up_guia.read())
        tem_guia = nif_escolhido in st.session_state.guias_irs
        st.caption("✅ Guia carregada" if tem_guia else "❌ Sem guia carregada ainda")

    if col_f is not None:
        with col_f:
            up_fatura = st.file_uploader("Fatura do Serviço de IRS (PDF)", type=["pdf"], key=f"up_fatura_irs_{nif_escolhido}")
            if up_fatura is not None:
                st.session_state.faturas_irs[nif_escolhido] = (up_fatura.name, up_fatura.read())
            tem_fatura = nif_escolhido in st.session_state.faturas_irs
            st.caption("✅ Fatura carregada" if tem_fatura else "❌ Sem fatura carregada ainda")
    else:
        st.session_state.faturas_irs.pop(nif_escolhido, None)

    with col_l:
        up_liq = st.file_uploader("Nota de Liquidação (PDF)", type=["pdf"], key=f"up_liq_irs_{nif_escolhido}")
        dados_liq = None
        if up_liq is not None:
            dados_liq = extrair_dados_liquidacao_irs(up_liq.getvalue(), nif_esperado=nif_escolhido)
            if dados_liq["nif_confirmado"] is False:
                st.warning(
                    f"⚠️ Não encontrei o NIF do cliente selecionado ({nif_escolhido}) neste PDF. "
                    "Confirma que carregaste o ficheiro certo."
                )
            if dados_liq["valor_apurado"] is None:
                st.warning("Não consegui ler o valor automaticamente neste PDF — preenche manualmente abaixo.")
            else:
                rotulo_legivel = {"a pagar": "a pagar", "a receber": "a receber (reembolso)", "apurado": "apurado (sem valor a pagar/receber)"}.get(dados_liq["tipo_valor"], "")
                st.success(
                    f"Valor {rotulo_legivel}: {abs(dados_liq['valor_apurado']):,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")
                )
            # Um widget com "key" só usa o "value=" passado na primeira vez que é criado — depois disso,
            # o Streamlit mantém o que já está gravado em session_state para essa key. Por isso, sempre que
            # detetamos um ficheiro novo (nome+tamanho diferente do último processado para este cliente),
            # atualizamos nós próprios o session_state antes dos campos serem criados mais abaixo.
            ficheiro_id = f"{up_liq.name}_{up_liq.size}"
            chave_rastreio = f"_liq_processado_{nif_escolhido}"
            if st.session_state.get(chave_rastreio) != ficheiro_id:
                st.session_state[chave_rastreio] = ficheiro_id
                if dados_liq["valor_apurado"] is not None:
                    st.session_state[f"valor_apurado_{nif_escolhido}"] = dados_liq["valor_apurado"]
                if dados_liq["numero_liquidacao"]:
                    st.session_state[f"num_liq_{nif_escolhido}"] = dados_liq["numero_liquidacao"]

    with col_p:
        up_pend = st.file_uploader("Controlo de Pendentes (PDF, opcional)", type=["pdf"], key=f"up_pend_irs_{nif_escolhido}")
        dados_pend = None
        if up_pend is not None:
            dados_pend = extrair_dados_pendentes_irs(up_pend.getvalue())
            if dados_pend["nif"] and dados_pend["nif"] != nif_escolhido:
                st.warning(f"⚠️ O NIF encontrado no PDF ({dados_pend['nif']}) não corresponde ao cliente selecionado ({nif_escolhido}).")
            if dados_pend["valor_pendente"] is None:
                st.warning("Não consegui ler o total pendente automaticamente — preenche manualmente abaixo, se aplicável.")
            else:
                st.success(f"Total pendente lido: {dados_pend['valor_pendente']:,.2f} €".replace(",", "X").replace(".", ",").replace("X", "."))
            ficheiro_id_pend = f"{up_pend.name}_{up_pend.size}"
            chave_rastreio_pend = f"_pend_processado_{nif_escolhido}"
            if st.session_state.get(chave_rastreio_pend) != ficheiro_id_pend:
                st.session_state[chave_rastreio_pend] = ficheiro_id_pend
                if dados_pend["valor_pendente"] is not None:
                    st.session_state[f"valor_pendente_{nif_escolhido}"] = dados_pend["valor_pendente"]

    st.divider()
    st.markdown("### 2. Confirmar Valores (edita se necessário antes de gravar)")
    valor_default = dados_liq["valor_apurado"] if dados_liq and dados_liq["valor_apurado"] is not None else float(row_atual["Valor_Apurado"])
    numero_default = dados_liq["numero_liquidacao"] if dados_liq and dados_liq["numero_liquidacao"] else row_atual["Numero_Liquidacao"]
    pendente_default = dados_pend["valor_pendente"] if dados_pend and dados_pend["valor_pendente"] is not None else float(row_atual["Valor_Pendente"])

    c1, c2, c3 = st.columns(3)
    with c1:
        numero_liq_edit = st.text_input("Nº de Liquidação (opcional)", value=numero_default, key=f"num_liq_{nif_escolhido}")
    with c2:
        valor_edit = st.number_input(
            "Valor Apurado (€) — positivo = a pagar, negativo = a receber",
            value=float(valor_default), step=0.01, format="%.2f", key=f"valor_apurado_{nif_escolhido}",
        )
    with c3:
        pendente_edit = st.number_input("Valor Pendente (€, à SERVE)", value=float(pendente_default), step=0.01, format="%.2f", key=f"valor_pendente_{nif_escolhido}")

    if st.button("💾 Guardar dados deste cliente"):
        novo_irs = pd.DataFrame(st.session_state.irs_dados)
        if nif_escolhido in novo_irs["NIF"].values:
            novo_irs = novo_irs[novo_irs["NIF"] != nif_escolhido]
        novo_linha = pd.DataFrame([{
            "NIF": nif_escolhido, "Numero_Liquidacao": numero_liq_edit,
            "Valor_Apurado": valor_edit, "Valor_Pendente": pendente_edit,
            "Incluido_Avenca": incluido_avenca, "Email_Enviado": bool(row_atual["Email_Enviado"]),
        }])
        persistir_irs(pd.concat([novo_irs, novo_linha], ignore_index=True)[IRS_COLS])
        st.success("Dados guardados.")
        st.rerun()

    st.divider()
    st.markdown("### 3. Pré-visualizar e Enviar Email")
    tpl = st.session_state.template_irs
    row_preview = row_atual.copy()
    row_preview["Numero_Liquidacao"] = numero_liq_edit
    row_preview["Valor_Apurado"] = valor_edit
    row_preview["Valor_Pendente"] = pendente_edit
    assunto, corpo = render_template_irs(tpl, row_preview)

    st.text_input("Assunto (preview)", value=assunto, disabled=True)
    if row_atual["Gestor_Email"]:
        st.caption(f"📋 CC: {row_atual['Gestor_Nome'] or ''} <{row_atual['Gestor_Email']}>")
    else:
        st.caption("📋 CC: — (sem gestor definido para este cliente)")
    st.text_area("Corpo (preview)", value=corpo, height=280, disabled=True)

    anexos_disponiveis = []
    if nif_escolhido in st.session_state.guias_irs:
        anexos_disponiveis.append(("Guia", st.session_state.guias_irs[nif_escolhido]))
    if up_liq is not None:
        anexos_disponiveis.append(("Nota de Liquidação", (up_liq.name, up_liq.getvalue())))
    if up_pend is not None:
        anexos_disponiveis.append(("Controlo de Pendentes", (up_pend.name, up_pend.getvalue())))
    if not incluido_avenca and nif_escolhido in st.session_state.faturas_irs:
        anexos_disponiveis.append(("Fatura", st.session_state.faturas_irs[nif_escolhido]))
    st.caption("📎 Anexos que vão ser enviados: " + (", ".join(a[0] for a in anexos_disponiveis) if anexos_disponiveis else "nenhum carregado ainda"))
    if not incluido_avenca and nif_escolhido not in st.session_state.faturas_irs:
        st.caption("⚠️ Este cliente não tem o serviço incluído na avença e ainda não carregaste a fatura — normalmente deve ir junto.")

    if not row_atual["Email"]:
        st.warning("Este cliente não tem email preenchido no registo central — não é possível enviar.")

    st.divider()
    smtp_cfg = smtp_config_form()

    if st.button("🚀 Enviar Email", type="primary", disabled=not row_atual["Email"]):
        if not smtp_cfg["utilizador"] or not smtp_cfg["password"]:
            st.error("Preencher utilizador e password SMTP.")
        else:
            try:
                cc_gestor = [row_atual["Gestor_Email"]] if row_atual["Gestor_Email"] else []
                anexos = [f for _, f in anexos_disponiveis]
                enviar_email(smtp_cfg, row_atual["Email"], assunto, corpo, anexos, cc=cc_gestor)

                novo_irs = pd.DataFrame(st.session_state.irs_dados)
                novo_irs = novo_irs[novo_irs["NIF"] != nif_escolhido]
                novo_linha = pd.DataFrame([{
                    "NIF": nif_escolhido, "Numero_Liquidacao": numero_liq_edit,
                    "Valor_Apurado": valor_edit, "Valor_Pendente": pendente_edit,
                    "Incluido_Avenca": incluido_avenca, "Email_Enviado": True,
                }])
                persistir_irs(pd.concat([novo_irs, novo_linha], ignore_index=True)[IRS_COLS])

                registar_log({
                    "data": datetime.now().strftime("%Y-%m-%d %H:%M"), "nif": nif_escolhido,
                    "nome": row_atual["Nome"], "pagamento": 0, "estado": "IRS - Enviado",
                })
                st.success(f"Email enviado a {row_atual['Nome']} e estado guardado.")
                st.rerun()
            except Exception as e:
                registar_log({
                    "data": datetime.now().strftime("%Y-%m-%d %H:%M"), "nif": nif_escolhido,
                    "nome": row_atual["Nome"], "pagamento": 0, "estado": f"IRS - Erro: {e}",
                })
                st.error(f"Erro ao enviar: {e}")

# --- Template de Email -------------------------------------------------
with tab_template:
    st.subheader("Template do Email de Liquidação de IRS")
    tpl = st.session_state.template_irs
    if sou_admin():
        tpl["assunto"] = st.text_input("Assunto", value=tpl["assunto"], key="irs_tpl_assunto")
        tpl["corpo"] = st.text_area("Corpo", value=tpl["corpo"], height=320, key="irs_tpl_corpo")
        st.caption(
            "Placeholders disponíveis: {nome} {nif} {email} {ref_liquidacao} {frase_valor} {frase_pendente}. "
            "{ref_liquidacao} já vem formatado como ', n.º de liquidação XXXX' (ou vazio, se não houver). "
            "{frase_valor} e {frase_pendente} são frases já prontas, geradas automaticamente a partir dos valores — não precisas de os escrever à mão."
        )
    else:
        st.caption("O template de email é definido pelo administrador.")
        st.text_input("Assunto", value=tpl["assunto"], disabled=True)
        st.text_area("Corpo", value=tpl["corpo"], height=320, disabled=True)

# Persistir template caso o admin o tenha editado (RLS bloqueia gestores).
guardar_config_db(st.session_state.params, st.session_state.templates, st.session_state.template_irs)
