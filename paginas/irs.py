"""
Página de IRS — tudo o que é específico deste imposto vive aqui: dados de
liquidação por cliente (lidos automaticamente dos PDFs sempre que possível,
mas sempre confirmáveis/editáveis antes de gravar ou enviar), upload de guia,
nota de liquidação e controlo de pendentes, e envio do respetivo email.
A guia e a fatura ficam guardadas no arquivo persistente (Supabase Storage).

Tem também um módulo à parte, "IRS Avulso (por número)", para lotes de
clientes só de IRS que chegam já numerados (sem NIF fiável à mão) — vive
numa tabela e numeração completamente separadas do resto da plataforma.
"""

from datetime import date, datetime

import pandas as pd
import streamlit as st

from common import (
    IRS_COLS,
    PASTAS_TIPO_DOC_IRS_AVULSO,
    carregar_clientes_irs_avulso_db,
    clean_clientes_df,
    clean_irs_avulso_df,
    clean_irs_df,
    editor_template_bilingue,
    enviar_email,
    escolher_conta_email,
    extrair_dados_liquidacao_irs,
    extrair_dados_pendentes_irs,
    extrair_numero_de_filename,
    formatar_valor,
    gerar_excel_irs,
    guardar_config_db,
    ler_ficheiro_importacao,
    listar_extras_generico,
    meu_email,
    montar_base_irs,
    obter_documentos_irs_avulso,
    persistir_clientes,
    persistir_clientes_irs_avulso,
    persistir_irs,
    registar_log,
    render_template_irs,
    storage_apagar,
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

tab_importar, tab_visao, tab_processar, tab_avulso, tab_template = st.tabs(
    ["📥 Importar Clientes", "📊 Visão Geral", "📎 Processar Cliente", "🔢 IRS Avulso (por número)", "✏️ Template de Email"]
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
    st.info(
        "💡 Se preferes só associar os documentos por NÚMERO (ex: '1 - Guia - Miguel Silva.pdf'), sem "
        "estes clientes entrarem no registo central de clientes, usa antes a aba 'IRS Avulso (por número)'."
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

# --- Visão Geral -----------------------------------------------------------
with tab_visao:
    if base_irs.empty:
        st.info("Ainda não há clientes com 'Aplica IRS' ligado — importa-os na aba 'Importar Clientes' ou ativa o interruptor na página 'Clientes'.")
    else:
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
    if base_irs.empty:
        st.info("Ainda não há clientes com 'Aplica IRS' ligado — importa-os na aba 'Importar Clientes' ou ativa o interruptor na página 'Clientes'.")
    else:
        st.subheader("Selecionar Cliente")
        nif_escolhido = st.selectbox(
            "Cliente",
            base_irs["NIF"].tolist(),
            format_func=lambda n: f"{n} — {base_irs.loc[base_irs['NIF']==n,'Nome'].values[0]}",
            key="irs_cliente_escolhido",
        )
        row_atual = base_irs[base_irs["NIF"] == nif_escolhido].iloc[0]
        # Os documentos ficam arquivados por ano — mudar o "ano dos dados" nas
        # Configurações muda também o arquivo consultado.
        ficheiros_arquivo = storage_listar(f"irs/{ano_dados}/{nif_escolhido}")

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
                    storage_upload_pdf(f"irs/{ano_dados}/{nif_escolhido}/guia.pdf", up_guia.getvalue())
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
                        storage_upload_pdf(f"irs/{ano_dados}/{nif_escolhido}/fatura.pdf", up_fatura.getvalue())
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
                        conteudo = storage_download_pdf(f"irs/{ano_dados}/{nif_escolhido}/guia.pdf")
                        if conteudo:
                            anexos.append((f"Guia_IRS_{nif_escolhido}.pdf", conteudo))
                    if up_liq is not None:
                        anexos.append((up_liq.name, up_liq.getvalue()))
                    if up_pend is not None:
                        anexos.append((up_pend.name, up_pend.getvalue()))
                    if not incluido_avenca and "fatura.pdf" in ficheiros_arquivo:
                        conteudo = storage_download_pdf(f"irs/{ano_dados}/{nif_escolhido}/fatura.pdf")
                        if conteudo:
                            anexos.append((f"Fatura_{nif_escolhido}.pdf", conteudo))

                    cc_gestor = [row_atual["Gestor_Email"]] if row_atual["Gestor_Email"] else []
                    enviar_email(smtp_cfg, row_atual["Email"], assunto, corpo, anexos, cc=cc_gestor,
                                 bcc=[smtp_cfg["remetente"]],
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

# --- IRS Avulso (por número) --------------------------------------------
with tab_avulso:
    st.subheader(f"IRS Avulso — {ano_dados}")
    st.caption(
        "Registo completamente à parte do resto da plataforma — nem entra no registo central de "
        "clientes, nem partilha numeração com ele. Cada cliente é identificado só pelo NÚMERO que "
        "lhe deres aqui (o mesmo número que usas no nome dos ficheiros, ex: '1 - Guia - Miguel Silva.pdf'). "
        "Como este número é só desta tabela, nunca entra em conflito com o N.º dos clientes normais."
    )

    base_avulso = carregar_clientes_irs_avulso_db(ano_dados)
    st.session_state.irs_avulso_clientes = base_avulso

    sub_importar, sub_docs, sub_estado, sub_emails = st.tabs(
        ["📥 Importar", "📎 Documentos", "📊 Estado", "✉️ Emails"]
    )

    # --- Importar clientes avulsos (numeração própria) --------------------
    with sub_importar:
        st.markdown(f"**Clientes de IRS avulso já importados para {ano_dados}:** {len(base_avulso)}")
        col_up_av, col_tpl_av = st.columns([2, 1])
        with col_tpl_av:
            template_avulso_csv = pd.DataFrame(
                [{"Numero": "1", "NIF": "123456789", "Nome": "Cliente Exemplo", "Email": "cliente@exemplo.pt", "Lingua": "PT"}]
            ).to_csv(index=False, sep=";")
            st.download_button("📥 Template CSV (IRS Avulso)", template_avulso_csv,
                                file_name="template_irs_avulso.csv", mime="text/csv")
            st.caption("Colunas: Numero, NIF, Nome, Email, Lingua (PT ou EN). O NIF é opcional — o Número é que identifica o cliente.")
        with col_up_av:
            up_avulso_csv = st.file_uploader(
                "Importar CSV ou Excel (colunas: Numero, NIF, Nome, Email, Lingua)",
                type=["csv", "xlsx"], key="up_irs_avulso_csv",
            )
        if up_avulso_csv is not None:
            try:
                if up_avulso_csv.name.endswith(".csv"):
                    bruto_av = pd.read_csv(up_avulso_csv, sep=None, engine="python", dtype=str)
                else:
                    bruto_av = pd.read_excel(up_avulso_csv, dtype=str)
                bruto_av.columns = [str(c).strip() for c in bruto_av.columns]
                aliases_av = {"N.º": "Numero", "N°": "Numero", "Número": "Numero", "Nº": "Numero"}
                bruto_av = bruto_av.rename(columns={c: aliases_av.get(c, c) for c in bruto_av.columns})
                novos_av = clean_irs_avulso_df(bruto_av)
                st.markdown(f"**{len(novos_av)} cliente(s) lidos do ficheiro:**")
                st.dataframe(novos_av[["Numero", "NIF", "Nome", "Email", "Lingua"]], use_container_width=True, hide_index=True)

                numeros_novos = set(novos_av["Numero"])
                numeros_existentes = set(base_avulso["Numero"]) - numeros_novos
                # dentro do próprio ficheiro importado
                repetidos_no_ficheiro = novos_av["Numero"][novos_av["Numero"].duplicated()].unique().tolist()
                if repetidos_no_ficheiro:
                    st.error(f"⚠️ Números repetidos dentro do próprio ficheiro (corrige antes de importar): {', '.join(repetidos_no_ficheiro)}")
                sobrepostos = numeros_novos & set(base_avulso["Numero"])
                if sobrepostos:
                    st.warning(f"Estes números já existem em {ano_dados} e vão ser ATUALIZADOS (substituídos pelos novos dados): {', '.join(sorted(sobrepostos))}")

                if st.button("✅ Confirmar importação (IRS Avulso)", disabled=bool(repetidos_no_ficheiro)):
                    completo = pd.concat([base_avulso[base_avulso["Numero"].isin(numeros_existentes)], novos_av], ignore_index=True)
                    persistir_clientes_irs_avulso(completo, ano_dados)
                    st.success(f"{len(novos_av)} cliente(s) de IRS avulso importados/atualizados para {ano_dados}.")
                    st.rerun()
            except Exception as e:
                st.error(f"Erro ao importar: {e}")

        if not base_avulso.empty:
            st.divider()
            st.markdown("**Clientes já importados** (edita Nome/Email/Língua/Gestor diretamente se precisares)")
            edit_av = st.data_editor(
                base_avulso[["Numero", "NIF", "Nome", "Email", "Lingua", "Gestor_Nome", "Gestor_Email"]],
                use_container_width=True, hide_index=True, height=300,
                disabled=["Numero"],
                key=f"editor_irs_avulso_clientes_{ano_dados}",
            )
            if st.button("💾 Guardar alterações aos clientes de IRS avulso"):
                restante = base_avulso.drop(columns=["NIF", "Nome", "Email", "Lingua", "Gestor_Nome", "Gestor_Email"]).merge(
                    edit_av, on="Numero", how="left"
                )
                persistir_clientes_irs_avulso(clean_irs_avulso_df(restante), ano_dados)
                st.success("Alterações guardadas.")
                st.rerun()

    if base_avulso.empty:
        with sub_docs:
            st.info("Importa primeiro os clientes na aba 'Importar' acima.")
        with sub_estado:
            st.info("Importa primeiro os clientes na aba 'Importar' acima.")
        with sub_emails:
            st.info("Importa primeiro os clientes na aba 'Importar' acima.")
    else:
        # Ficheiros já no arquivo, por categoria: {pasta: {numero: [nome1.pdf, ...]}} —
        # mesmo formato usado na Segurança Social, já suporta MAIS DO QUE UM
        # ficheiro por categoria/cliente.
        arquivos_avulso = {
            pasta: listar_extras_generico(f"irs_avulso/{ano_dados}/{pasta}")
            for pasta in PASTAS_TIPO_DOC_IRS_AVULSO.values()
        }

        # --- Documentos (carregamento em massa por número) -----------------
        with sub_docs:
            st.markdown("**Carregamento em massa** (o nome do ficheiro deve começar pelo número, ex: '1 - Guia - Miguel Silva.pdf')")
            st.caption(
                "Carregar um novo ficheiro para um cliente SUBSTITUI o(s) ficheiro(s) que já lá estavam "
                "dessa categoria — não precisas de apagar antes de corrigir. Se quiseres mesmo mais do que "
                "um ficheiro do mesmo tipo para o mesmo cliente, carrega-os todos DE UMA VEZ (ficam numerados)."
            )
            col_tipo_av, col_up_av2 = st.columns([1, 3])
            with col_tipo_av:
                tipo_doc_av = st.radio("Tipo de documento", list(PASTAS_TIPO_DOC_IRS_AVULSO.keys()), key="irs_avulso_tipo_doc")
            with col_up_av2:
                up_massa_av = st.file_uploader(
                    "Carregar PDFs (nome a começar pelo número)", type=["pdf"], accept_multiple_files=True,
                    key=f"irs_avulso_up_massa_{tipo_doc_av}",
                    # A chave inclui o tipo de documento de propósito — mesma correção já aplicada na
                    # Segurança Social, para o widget não reenviar os mesmos ficheiros ao trocar de categoria.
                )
            if up_massa_av:
                ids_up_av = tuple(sorted(f"{f.name}_{f.size}" for f in up_massa_av))
                if st.session_state.get("_irs_avulso_massa_proc") != (ano_dados, tipo_doc_av, ids_up_av):
                    st.session_state["_irs_avulso_massa_proc"] = (ano_dados, tipo_doc_av, ids_up_av)
                    pasta_av = PASTAS_TIPO_DOC_IRS_AVULSO[tipo_doc_av]
                    ok_av, sem_numero, detalhes_av = 0, [], []
                    atualizacoes_liq = {}   # numero -> {"Numero_Liquidacao":..., "Valor_Apurado":...}
                    atualizacoes_pend = {}  # numero -> valor_pendente
                    ficheiros_por_numero = {}
                    for f in up_massa_av:
                        numero_f = extrair_numero_de_filename(f.name)
                        if not numero_f:
                            sem_numero.append(f.name)
                            continue
                        ficheiros_por_numero.setdefault(numero_f, []).append(f)
                    for numero_f, ficheiros in ficheiros_por_numero.items():
                        # Substitui o(s) ficheiro(s) anteriores desta categoria para este número.
                        for nome_antigo in arquivos_avulso.get(pasta_av, {}).get(numero_f, []):
                            storage_apagar(f"irs_avulso/{ano_dados}/{pasta_av}/{numero_f}__{nome_antigo}")
                        nomes_novos = ([f"{pasta_av}.pdf"] if len(ficheiros) == 1
                                        else [f"{pasta_av}_{i}.pdf" for i in range(1, len(ficheiros) + 1)])
                        for f, nome_novo in zip(ficheiros, nomes_novos):
                            conteudo_f = f.getvalue()
                            caminho_f = f"irs_avulso/{ano_dados}/{pasta_av}/{numero_f}__{nome_novo}"
                            try:
                                storage_upload_pdf(caminho_f, conteudo_f)
                                detalhes_av.append(f"✅ {numero_f} → {caminho_f}")
                                ok_av += 1
                                # Só tenta ler o valor automaticamente quando há um SÓ ficheiro
                                # deste tipo para este cliente neste carregamento — com vários,
                                # não há forma fiável de saber qual valor usar.
                                if len(ficheiros) == 1 and pasta_av == "liquidacao":
                                    dados_liq_av = extrair_dados_liquidacao_irs(conteudo_f)
                                    upd = {}
                                    if dados_liq_av["numero_liquidacao"]:
                                        upd["Numero_Liquidacao"] = dados_liq_av["numero_liquidacao"]
                                    if dados_liq_av["valor_apurado"] is not None:
                                        upd["Valor_Apurado"] = dados_liq_av["valor_apurado"]
                                    if upd:
                                        atualizacoes_liq[numero_f] = upd
                                elif len(ficheiros) == 1 and pasta_av == "pendentes":
                                    dados_pend_av = extrair_dados_pendentes_irs(conteudo_f)
                                    if dados_pend_av["valor_pendente"] is not None:
                                        atualizacoes_pend[numero_f] = dados_pend_av["valor_pendente"]
                            except Exception as e:
                                detalhes_av.append(f"❌ {numero_f} → {caminho_f}: {e}")

                    # Aplica os valores lidos automaticamente (Liquidação/Pendentes) aos clientes afetados.
                    if atualizacoes_liq or atualizacoes_pend:
                        base_upd = base_avulso.copy()
                        for numero_f, upd in atualizacoes_liq.items():
                            idx = base_upd.index[base_upd["Numero"] == numero_f]
                            for k, v in upd.items():
                                base_upd.loc[idx, k] = v
                        for numero_f, valor_pend in atualizacoes_pend.items():
                            idx = base_upd.index[base_upd["Numero"] == numero_f]
                            base_upd.loc[idx, "Valor_Pendente"] = valor_pend
                        persistir_clientes_irs_avulso(base_upd, ano_dados)

                    msg_av = f"{ok_av} ficheiro(s) associados e guardados no arquivo."
                    if atualizacoes_liq:
                        msg_av += f" Valor/nº de liquidação lido automaticamente em {len(atualizacoes_liq)} cliente(s)."
                    if atualizacoes_pend:
                        msg_av += f" Valor pendente lido automaticamente em {len(atualizacoes_pend)} cliente(s)."
                    if sem_numero:
                        msg_av += f" Sem número no nome do ficheiro (ignorados): {', '.join(sem_numero)}"
                    st.session_state["_irs_avulso_ultimo_upload"] = {"msg": msg_av, "detalhes": detalhes_av}
                    st.rerun()

            ultimo_av = st.session_state.get("_irs_avulso_ultimo_upload")
            if ultimo_av:
                st.success(ultimo_av["msg"])
                if ultimo_av["detalhes"]:
                    with st.expander(f"Ver os {len(ultimo_av['detalhes'])} caminho(s) exatos onde foi guardado"):
                        st.text("\n".join(ultimo_av["detalhes"]))

            st.divider()
            st.markdown("**Carregamento por cliente**")
            numero_doc = st.selectbox(
                "Cliente", base_avulso["Numero"].tolist(),
                format_func=lambda n: f"{n} — {base_avulso.loc[base_avulso['Numero']==n,'Nome'].values[0]}",
                key="irs_avulso_cliente_doc",
            )

            total_docs_numero = sum(len(d.get(numero_doc, [])) for d in arquivos_avulso.values())
            if total_docs_numero:
                with st.expander(f"🗑️ Apagar todos os documentos deste cliente ({total_docs_numero})"):
                    st.caption("Apaga TODOS os documentos (IRS, Liquidação, Guia, Fatura e Pendentes) deste cliente, só para este ano.")
                    if st.button("Confirmar — apagar tudo", key=f"irs_avulso_apagar_tudo_{ano_dados}_{numero_doc}", type="primary"):
                        for pasta_x, dic_x in arquivos_avulso.items():
                            for nome_x in dic_x.get(numero_doc, []):
                                storage_apagar(f"irs_avulso/{ano_dados}/{pasta_x}/{numero_doc}__{nome_x}")
                        st.success("Documentos apagados.")
                        st.rerun()

            def _documentos_irs_avulso(col, rotulo: str, pasta: str, dicionario: dict):
                """Upload (aceita vários ficheiros) + lista com botão de apagar por
                ficheiro, para uma categoria de IRS Avulso — mesmo padrão da
                Segurança Social: carregar de novo SUBSTITUI o(s) ficheiro(s)
                anteriores; carregar vários de uma vez mantém-nos todos (numerados)."""
                with col:
                    up = st.file_uploader(rotulo, type=["pdf"], accept_multiple_files=True,
                                           key=f"irs_avulso_up_{pasta}_{ano_dados}_{numero_doc}")
                    if up:
                        ids_up = tuple(sorted(f"{f.name}_{f.size}" for f in up))
                        chave_proc = f"_irs_avulso_{pasta}_proc_{ano_dados}_{numero_doc}"
                        if st.session_state.get(chave_proc) != ids_up:
                            st.session_state[chave_proc] = ids_up
                            for nome_antigo in dicionario.get(numero_doc, []):
                                storage_apagar(f"irs_avulso/{ano_dados}/{pasta}/{numero_doc}__{nome_antigo}")
                            nomes_novos = ([f"{pasta}.pdf"] if len(up) == 1
                                            else [f"{pasta}_{i}.pdf" for i in range(1, len(up) + 1)])
                            for f, nome_novo in zip(up, nomes_novos):
                                storage_upload_pdf(f"irs_avulso/{ano_dados}/{pasta}/{numero_doc}__{nome_novo}", f.getvalue())
                            st.success(f"{len(up)} ficheiro(s) de {rotulo} guardado(s). Substituiu os anteriores.")
                            st.rerun()
                    existentes = dicionario.get(numero_doc, [])
                    if existentes:
                        for nome in existentes:
                            c_nome, c_apagar = st.columns([4, 1])
                            c_nome.caption(f"📄 {nome}")
                            if c_apagar.button("🗑️", key=f"irs_avulso_apagar_{pasta}_{ano_dados}_{numero_doc}_{nome}",
                                                help="Apagar (depois podes carregar outro em substituição)"):
                                storage_apagar(f"irs_avulso/{ano_dados}/{pasta}/{numero_doc}__{nome}")
                                st.rerun()
                    else:
                        st.caption(f"Sem {rotulo}")

            cols_av = st.columns(5)
            for col_av, (rotulo_av, pasta_av2) in zip(cols_av, PASTAS_TIPO_DOC_IRS_AVULSO.items()):
                _documentos_irs_avulso(col_av, rotulo_av, pasta_av2, arquivos_avulso.get(pasta_av2, {}))

        # --- Estado ----------------------------------------------------------
        with sub_estado:
            st.markdown("**Estado por cliente**")

            def _marca_av(pasta: str, numero: str) -> str:
                n = len(arquivos_avulso.get(pasta, {}).get(numero, []))
                if n == 0:
                    return "❌"
                return "✅" if n == 1 else f"✅ ({n})"

            linhas_estado = []
            for _, r in base_avulso.iterrows():
                linhas_estado.append({
                    "Número": r["Numero"], "NIF": r["NIF"], "Nome": r["Nome"],
                    "IRS": _marca_av("irs", r["Numero"]),
                    "Liquidação": _marca_av("liquidacao", r["Numero"]),
                    "Guia": _marca_av("guia", r["Numero"]),
                    "Fatura": _marca_av("fatura", r["Numero"]),
                    "Pendentes": _marca_av("pendentes", r["Numero"]),
                    "Nº Liquidação": r["Numero_Liquidacao"],
                    "Valor Apurado": r["Valor_Apurado"],
                    "Valor Pendente": r["Valor_Pendente"],
                    "Email Enviado": bool(r["Email_Enviado"]),
                })
            st.dataframe(pd.DataFrame(linhas_estado), use_container_width=True, hide_index=True, height=400)

        # --- Emails ------------------------------------------------------------
        with sub_emails:
            st.subheader(f"Enviar Emails — IRS Avulso {ano_dados}")
            elegiveis_av = base_avulso[base_avulso["Email"].str.strip() != ""].copy()
            sem_email_av = len(base_avulso) - len(elegiveis_av)
            if sem_email_av:
                st.caption(f"⚠️ {sem_email_av} cliente(s) sem email preenchido — não aparecem abaixo.")

            tpl_av = st.session_state.template_irs

            preview_num = st.selectbox(
                "Pré-visualizar cliente:", elegiveis_av["Numero"].tolist(),
                format_func=lambda n: f"{n} — {elegiveis_av.loc[elegiveis_av['Numero']==n,'Nome'].values[0]}",
                key="irs_avulso_preview",
            )
            docs_prev_av = []
            if preview_num:
                row_prev_av = elegiveis_av[elegiveis_av["Numero"] == preview_num].iloc[0]
                docs_prev_av = obter_documentos_irs_avulso(ano_dados, preview_num, arquivos_avulso)
                assunto_av, corpo_av = render_template_irs(tpl_av, row_prev_av)
                st.text_input("Assunto (preview)", value=assunto_av, disabled=True)
                if row_prev_av["Gestor_Email"]:
                    st.caption(f"📋 CC: {row_prev_av['Gestor_Nome'] or ''} <{row_prev_av['Gestor_Email']}>")
                else:
                    st.caption("📋 CC: — (sem gestor definido)")
                st.text_area("Corpo (preview)", value=corpo_av, height=260, disabled=True)
                st.caption("📎 Anexos: " + (", ".join(d["tipo"] for d in docs_prev_av) if docs_prev_av else "nenhum documento carregado ainda"))

                with st.expander("🔍 Diagnóstico deste cliente"):
                    if docs_prev_av:
                        for d in docs_prev_av:
                            st.text(f"{d['tipo']}   →   {d['caminho']}")
                    else:
                        st.caption("Nenhum documento encontrado para este número.")

            st.divider()
            smtp_cfg_av = escolher_conta_email("irs_avulso")

            nao_enviados_av = [n for n in elegiveis_av["Numero"] if not elegiveis_av.loc[elegiveis_av["Numero"] == n, "Email_Enviado"].values[0]]
            selecionados_av = st.multiselect(
                "Clientes selecionados para envio",
                elegiveis_av["Numero"].tolist(),
                default=nao_enviados_av,
                format_func=lambda n: f"{n} — {elegiveis_av.loc[elegiveis_av['Numero']==n,'Nome'].values[0]}"
                + ("  ✅ já enviado" if elegiveis_av.loc[elegiveis_av["Numero"] == n, "Email_Enviado"].values[0] else ""),
                key="irs_avulso_selecionados",
            )

            if st.button("🚀 Enviar Emails Selecionados (IRS Avulso)", type="primary", disabled=not selecionados_av):
                if not smtp_cfg_av["utilizador"] or not smtp_cfg_av["password"]:
                    st.error("Escolhe ou cria uma conta de email com utilizador e password preenchidos.")
                else:
                    progress_av = st.progress(0.0)
                    status_av = st.empty()
                    assinatura_av = st.session_state.params.get("assinatura_html", "")
                    sucessos_av, falhas_av = 0, 0
                    base_atualizada = base_avulso.copy()
                    for i, numero_e in enumerate(selecionados_av):
                        row_e = elegiveis_av[elegiveis_av["Numero"] == numero_e].iloc[0]
                        docs_e = obter_documentos_irs_avulso(ano_dados, numero_e, arquivos_avulso)
                        anexos_e = []
                        for d in docs_e:
                            conteudo_e = storage_download_pdf(d["caminho"])
                            if conteudo_e:
                                anexos_e.append((d["anexo"], conteudo_e))
                        assunto_e, corpo_e = render_template_irs(tpl_av, row_e)
                        try:
                            cc_gestor_e = [row_e["Gestor_Email"]] if row_e["Gestor_Email"] else []
                            enviar_email(smtp_cfg_av, row_e["Email"], assunto_e, corpo_e, anexos_e, cc=cc_gestor_e,
                                         bcc=[smtp_cfg_av["remetente"]], assinatura_html=assinatura_av)
                            base_atualizada.loc[base_atualizada["Numero"] == numero_e, "Email_Enviado"] = True
                            registar_log({
                                "data": datetime.now().strftime("%Y-%m-%d %H:%M"), "nif": row_e["NIF"] or numero_e,
                                "nome": row_e["Nome"], "pagamento": 0, "estado": "Enviado (IRS Avulso)",
                                "modulo": "IRS", "enviado_por": meu_email(),
                            })
                            sucessos_av += 1
                        except Exception as e:
                            registar_log({
                                "data": datetime.now().strftime("%Y-%m-%d %H:%M"), "nif": row_e["NIF"] or numero_e,
                                "nome": row_e["Nome"], "pagamento": 0, "estado": f"Erro (IRS Avulso): {e}",
                                "modulo": "IRS", "enviado_por": meu_email(),
                            })
                            falhas_av += 1
                        progress_av.progress((i + 1) / len(selecionados_av))
                        status_av.text(f"{i+1}/{len(selecionados_av)} — {row_e['Nome']}")
                    persistir_clientes_irs_avulso(base_atualizada, ano_dados)
                    st.success(f"Concluído: {sucessos_av} enviados, {falhas_av} com erro.")
                    st.rerun()

# --- Template de Email -------------------------------------------------
with tab_template:
    st.subheader("Template do Email de Liquidação de IRS")
    st.caption("Este template é partilhado entre os clientes de IRS normais e os de IRS Avulso.")
    tpl = st.session_state.template_irs
    editor_template_bilingue(tpl, "irs_tpl", altura=320)
    st.caption(
        "Placeholders disponíveis: {nome} {nif} {email} {ref_liquidacao} {frase_valor} {frase_pendente} {ano_dados} {ano_pagamentos}. "
        "{ref_liquidacao} já vem formatado como ', n.º de liquidação XXXX' (ou vazio, se não houver). "
        "{frase_valor} e {frase_pendente} são frases já prontas, geradas automaticamente a partir dos valores e na língua do cliente — não precisas de os escrever à mão. "
        "Alterações aqui ficam guardadas para toda a equipa."
    )

# Persistir template (guardado para toda a equipa, qualquer utilizador pode editar).
guardar_config_db(st.session_state.params, st.session_state.templates, st.session_state.template_irs)
