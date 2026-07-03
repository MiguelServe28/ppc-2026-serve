"""
Página de IRS — tudo o que é específico deste imposto vive aqui: dados de
liquidação por cliente (lidos automaticamente dos PDFs sempre que possível,
mas sempre confirmáveis/editáveis antes de gravar ou enviar), upload de guia,
nota de liquidação e controlo de pendentes, e envio do respetivo email.
A guia e a fatura ficam guardadas no arquivo persistente (Supabase Storage).
"""

from datetime import date, datetime

import pandas as pd
import streamlit as st

from common import (
    IRS_COLS,
    clean_clientes_df,
    clean_irs_df,
    editor_template_bilingue,
    enviar_email,
    escolher_conta_email,
    extrair_dados_liquidacao_irs,
    extrair_dados_pendentes_irs,
    formatar_valor,
    gerar_excel_irs,
    guardar_config_db,
    ler_ficheiro_importacao,
    meu_email,
    montar_base_irs,
    persistir_clientes,
    persistir_irs,
    registar_log,
    render_template_irs,
    sou_admin,
    storage_download_pdf,
    storage_listar,
    storage_upload_pdf,
)

ano_dados = st.session_state.params.get("ano_dados", 2025)

st.title(f"🧾 IRS — Liquidações {ano_dados}")
st.caption("SERVE — Contabilidade e Viabilização Empresarial")
st.caption(
    "Cada cliente é tratado individualmente: seleciona o cliente, carrega os documentos "
    "(Guia, Nota de Liquidação e, se aplicável, o Controlo de Pendentes), confirma os valores "
    "lidos automaticamente e envia o email. A guia e a fatura ficam guardadas no arquivo — não se perdem ao fechar o browser."
)

base_irs = montar_base_irs()

tab_importar, tab_visao, tab_processar, tab_template = st.tabs(
    ["📥 Importar Clientes", "📊 Visão Geral", "📎 Processar Cliente", "✏️ Template de Email"]
)

# --- Importar Clientes -------------------------------------------------
with tab_importar:
    st.subheader("Importar Clientes só de IRS")
    st.caption(
        "Usa isto para trazeres de uma vez uma lista de clientes que só têm IRS (não têm PPC nem outros "
        "impostos). Entram no registo central da plataforma, mas já ficam com 'Aplica IRS' ligado "
        "automaticamente, por isso nunca aparecem misturados nas contas de PPC ou de outro imposto — "
        "podes vê-los à parte na página 'Clientes', usando o filtro 'Só IRS'."
    )
    col_up, col_tpl = st.columns([2, 1])
    with col_tpl:
        template_irs_csv = pd.DataFrame(
            [{"N.º": "123", "NIF": "123456789", "Nome": "Cliente Exemplo", "Email": "cliente@exemplo.pt", "Lingua": "PT"}]
        ).to_csv(index=False, sep=";")
        st.download_button("📥 Template CSV (IRS)", template_irs_csv, file_name="template_clientes_irs.csv", mime="text/csv")
        st.caption("Colunas: N.º, NIF, Nome, Email, Lingua (PT ou EN). Estes clientes ficam marcados como 'Só IRS (avulso)'.")
    with col_up:
        up_irs_csv = st.file_uploader(
            "Importar CSV ou Excel (colunas: N.º, NIF, Nome, Email)",
            type=["csv", "xlsx"],
            key="up_irs_clientes_csv",
        )
    if up_irs_csv is not None:
        try:
            bruto_irs = ler_ficheiro_importacao(up_irs_csv)
            novos_irs = clean_clientes_df(bruto_irs)
            novos_irs["Aplica_IRS"] = True
            novos_irs["IRS_Avulso"] = True  # marca-os como "só IRS" — não são clientes de avença
            st.markdown(f"**{len(novos_irs)} cliente(s) lidos do ficheiro:**")
            st.dataframe(novos_irs[["Numero_Cliente", "NIF", "Nome", "Email"]].rename(columns={"Numero_Cliente": "N.º"}), use_container_width=True, hide_index=True)
            from common import nifs_invalidos
            invalidos = nifs_invalidos(novos_irs)
            if invalidos:
                st.warning(f"⚠️ NIFs com dígito de controlo inválido (confirma se estão bem escritos): {', '.join(invalidos)}")
            if st.button("✅ Confirmar importação destes clientes de IRS"):
                persistir_clientes(
                    clean_clientes_df(pd.concat([st.session_state.clientes, novos_irs], ignore_index=True))
                    .drop_duplicates(subset="NIF", keep="last")
                )
                st.success(f"{len(novos_irs)} cliente(s) importados/atualizados com 'Aplica IRS' ligado.")
                st.rerun()
        except Exception as e:
            st.error(f"Erro ao importar: {e}")

if base_irs.empty:
    st.info("Ainda não há clientes com 'Aplica IRS' ligado — importa-os aqui em cima ou ativa o interruptor na página 'Clientes'.")
    st.stop()

# --- Visão Geral -----------------------------------------------------------
with tab_visao:
    st.subheader("Estado por Cliente")

    FILTRO_TIPO_IRS = {
        "Todos": None,
        "Clientes de avença (base central, com pisco IRS)": False,
        "Só IRS (importados à parte no menu IRS)": True,
    }
    filtro_tipo = st.selectbox("Mostrar", list(FILTRO_TIPO_IRS.keys()), key="filtro_tipo_irs")
    alvo = FILTRO_TIPO_IRS[filtro_tipo]
    mostrados = base_irs if alvo is None else base_irs[base_irs["IRS_Avulso"] == alvo]
    st.caption(f"A mostrar {len(mostrados)} de {len(base_irs)} cliente(s) de IRS.")

    show_cols = ["Numero_Cliente", "NIF", "Nome", "Numero_Liquidacao", "Valor_Apurado", "Valor_Pendente", "Incluido_Avenca", "Email_Enviado"]
    st.caption("✏️ Podes marcar/desmarcar diretamente os piscos 'Incluído na Avença' e 'Email Enviado' — carrega em Guardar no fim.")
    editado = st.data_editor(
        mostrados[show_cols],
        use_container_width=True,
        hide_index=True,
        height=400,
        disabled=["Numero_Cliente", "NIF", "Nome", "Numero_Liquidacao", "Valor_Apurado", "Valor_Pendente"],
        column_config={
            "Numero_Cliente": st.column_config.TextColumn("N.º"),
            "Numero_Liquidacao": st.column_config.TextColumn("Nº Liquidação"),
            "Valor_Apurado": st.column_config.NumberColumn("Valor Apurado (€)", format="%.2f"),
            "Valor_Pendente": st.column_config.NumberColumn("Pendente (€)", format="%.2f"),
            "Incluido_Avenca": st.column_config.CheckboxColumn("Incluído na Avença"),
            "Email_Enviado": st.column_config.CheckboxColumn("Email Enviado"),
        },
        key=f"editor_visao_irs_{filtro_tipo}",
    )
    if st.button("💾 Guardar piscos"):
        novo = clean_irs_df(editado[["NIF", "Numero_Liquidacao", "Valor_Apurado", "Valor_Pendente", "Incluido_Avenca", "Email_Enviado"]])
        atual = clean_irs_df(pd.DataFrame(st.session_state.irs_dados))
        resto = atual[~atual["NIF"].isin(set(novo["NIF"]))]
        persistir_irs(pd.concat([resto, novo], ignore_index=True)[IRS_COLS])
        st.success("Piscos guardados.")
        st.rerun()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Clientes IRS (no filtro)", len(mostrados))
    c2.metric("A Pagar", int((mostrados["Valor_Apurado"] > 0).sum()))
    c3.metric("A Receber (reembolso)", int((mostrados["Valor_Apurado"] < 0).sum()))
    c4.metric("Emails Enviados", f"{int(mostrados['Email_Enviado'].sum())} / {len(mostrados)}")

    st.divider()
    excel_irs = gerar_excel_irs(base_irs, st.session_state.params)
    st.download_button(
        "⬇️ Descarregar Excel de Controlo (IRS)",
        data=excel_irs,
        file_name=f"Controlo_IRS_{ano_dados}_{date.today().isoformat()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.caption("Clientes com email já enviado ficam destacados a verde no Excel.")

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
    ficheiros_arquivo = storage_listar(f"irs/{nif_escolhido}")

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
            fid = f"{up_guia.name}_{up_guia.size}"
            if st.session_state.get(f"_guia_irs_proc_{nif_escolhido}") != fid:
                storage_upload_pdf(f"irs/{nif_escolhido}/guia.pdf", up_guia.getvalue())
                st.session_state[f"_guia_irs_proc_{nif_escolhido}"] = fid
                ficheiros_arquivo.add("guia.pdf")
        tem_guia = "guia.pdf" in ficheiros_arquivo
        st.caption("✅ Guia no arquivo" if tem_guia else "❌ Sem guia no arquivo ainda")

    if col_f is not None:
        with col_f:
            up_fatura = st.file_uploader("Fatura do Serviço de IRS (PDF)", type=["pdf"], key=f"up_fatura_irs_{nif_escolhido}")
            if up_fatura is not None:
                fid = f"{up_fatura.name}_{up_fatura.size}"
                if st.session_state.get(f"_fatura_irs_proc_{nif_escolhido}") != fid:
                    storage_upload_pdf(f"irs/{nif_escolhido}/fatura.pdf", up_fatura.getvalue())
                    st.session_state[f"_fatura_irs_proc_{nif_escolhido}"] = fid
                    ficheiros_arquivo.add("fatura.pdf")
            tem_fatura = "fatura.pdf" in ficheiros_arquivo
            st.caption("✅ Fatura no arquivo" if tem_fatura else "❌ Sem fatura no arquivo ainda")

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
                st.success(f"Valor {rotulo_legivel}: {formatar_valor(abs(dados_liq['valor_apurado']))} €")
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
                st.success(f"Total pendente lido: {formatar_valor(dados_pend['valor_pendente'])} €")
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

    anexos_previstos = []
    if "guia.pdf" in ficheiros_arquivo:
        anexos_previstos.append("Guia")
    if up_liq is not None:
        anexos_previstos.append("Nota de Liquidação")
    if up_pend is not None:
        anexos_previstos.append("Controlo de Pendentes")
    if not incluido_avenca and "fatura.pdf" in ficheiros_arquivo:
        anexos_previstos.append("Fatura")
    st.caption("📎 Anexos que vão ser enviados: " + (", ".join(anexos_previstos) if anexos_previstos else "nenhum carregado ainda"))
    if not incluido_avenca and "fatura.pdf" not in ficheiros_arquivo:
        st.caption("⚠️ Este cliente não tem o serviço incluído na avença e ainda não carregaste a fatura — normalmente deve ir junto.")

    if not row_atual["Email"]:
        st.warning("Este cliente não tem email preenchido no registo central — não é possível enviar.")

    st.divider()
    smtp_cfg = escolher_conta_email("irs")

    if st.button("🚀 Enviar Email", type="primary", disabled=not row_atual["Email"]):
        if not smtp_cfg["utilizador"] or not smtp_cfg["password"]:
            st.error("Escolhe ou cria uma conta de email com utilizador e password preenchidos.")
        else:
            try:
                anexos = []
                if "guia.pdf" in ficheiros_arquivo:
                    conteudo = storage_download_pdf(f"irs/{nif_escolhido}/guia.pdf")
                    if conteudo:
                        anexos.append((f"Guia_IRS_{nif_escolhido}.pdf", conteudo))
                if up_liq is not None:
                    anexos.append((up_liq.name, up_liq.getvalue()))
                if up_pend is not None:
                    anexos.append((up_pend.name, up_pend.getvalue()))
                if not incluido_avenca and "fatura.pdf" in ficheiros_arquivo:
                    conteudo = storage_download_pdf(f"irs/{nif_escolhido}/fatura.pdf")
                    if conteudo:
                        anexos.append((f"Fatura_{nif_escolhido}.pdf", conteudo))

                cc_gestor = [row_atual["Gestor_Email"]] if row_atual["Gestor_Email"] else []
                enviar_email(smtp_cfg, row_atual["Email"], assunto, corpo, anexos, cc=cc_gestor,
                             assinatura_html=st.session_state.params.get("assinatura_html", ""))

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
                    "nome": row_atual["Nome"], "pagamento": 0, "estado": "Enviado",
                    "modulo": "IRS", "enviado_por": meu_email(),
                })
                st.success(f"Email enviado a {row_atual['Nome']} e estado guardado.")
                st.rerun()
            except Exception as e:
                registar_log({
                    "data": datetime.now().strftime("%Y-%m-%d %H:%M"), "nif": nif_escolhido,
                    "nome": row_atual["Nome"], "pagamento": 0, "estado": f"Erro: {e}",
                    "modulo": "IRS", "enviado_por": meu_email(),
                })
                st.error(f"Erro ao enviar: {e}")

# --- Template de Email -------------------------------------------------
with tab_template:
    st.subheader("Template do Email de Liquidação de IRS")
    tpl = st.session_state.template_irs
    if sou_admin():
        editor_template_bilingue(tpl, "irs_tpl", altura=320)
        st.caption(
            "Placeholders disponíveis: {nome} {nif} {email} {ref_liquidacao} {frase_valor} {frase_pendente} {ano_dados} {ano_pagamentos}. "
            "{ref_liquidacao} já vem formatado como ', n.º de liquidação XXXX' (ou vazio, se não houver). "
            "{frase_valor} e {frase_pendente} são frases já prontas, geradas automaticamente a partir dos valores e na língua do cliente — não precisas de os escrever à mão."
        )
    else:
        st.caption("O template de email é definido pelo administrador.")
        st.text_input("Assunto", value=tpl["assunto"], disabled=True)
        st.text_area("Corpo", value=tpl["corpo"], height=320, disabled=True)

# Persistir template caso o admin o tenha editado (RLS bloqueia gestores).
guardar_config_db(st.session_state.params, st.session_state.templates, st.session_state.template_irs)
