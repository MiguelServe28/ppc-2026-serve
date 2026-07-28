"""
Página do AIMI (Adicional ao IMI) — ao contrário do IMI, este imposto paga-se
de uma vez só por ano, com prazo até 30 de setembro. Entram os clientes com o
pisco 'AIMI'. Reaproveita a mesma tabela de Responsáveis do IMI (o gestor do
imóvel é o mesmo para os dois impostos). Mesmo padrão da Segurança
Social/IVA/IMI.
"""

from datetime import date, datetime

import pandas as pd
import streamlit as st

from common import (
    carregar_envios_db,
    carregar_responsaveis_imi_db,
    data_limite_aimi,
    docs_ss_cliente,
    editor_template_bilingue,
    enviar_email,
    escolher_conta_email,
    extrair_nif_de_filename,
    gerar_excel_estado_mensal,
    guardar_config_db,
    guardar_responsavel_imi_db,
    listar_extras_generico,
    marcar_envio_db,
    meu_email,
    montar_base_aimi,
    nomes_ficheiro_unicos,
    registar_log,
    render_template_docs,
    sanitizar_nome_ficheiro,
    storage_download_pdf,
    storage_listar,
    storage_upload_pdf,
)

st.title("🏘️ AIMI — Adicional ao IMI")
st.caption("SERVE — Contabilidade e Viabilização Empresarial")

base_aimi = montar_base_aimi()
if base_aimi.empty:
    st.info("Não há clientes com o pisco 'AIMI' ligado — ativa-o na página 'Clientes'.")
    st.stop()

ano_aimi = st.selectbox("Ano", list(range(date.today().year, date.today().year - 4, -1)), key="aimi_ano")
st.caption(f"Prazo de pagamento: até {data_limite_aimi(ano_aimi).strftime('%d/%m/%Y')} (confirma sempre a data exata na nota de cobrança/aviso da Autoridade Tributária).")
periodo = str(ano_aimi)

guias_set = {n[:-4] for n in storage_listar(f"aimi/{periodo}/guia") if n.lower().endswith(".pdf")}
extras_dict = listar_extras_generico(f"aimi/{periodo}/extra")
enviados = carregar_envios_db("aimi_dados", periodo)

base_aimi = base_aimi.reset_index(drop=True)
base_aimi["Email_Enviado"] = base_aimi["NIF"].map(lambda n: enviados.get(n, False))

# Responsável (gestor do imóvel/empresa) — mesma tabela usada pelo IMI (é o
# mesmo imóvel/gestor); não faz parte do registo central de clientes. Quando
# preenchido, o email vai para ele em vez de ir para o próprio proprietário.
responsaveis_aimi = carregar_responsaveis_imi_db()
base_aimi["Responsavel_Nome"] = base_aimi["NIF"].map(lambda n: responsaveis_aimi.get(n, {}).get("nome", ""))
base_aimi["Responsavel_Email"] = base_aimi["NIF"].map(lambda n: responsaveis_aimi.get(n, {}).get("email", ""))
base_aimi["Destinatario_Email"] = base_aimi["Responsavel_Email"].where(
    base_aimi["Responsavel_Email"].str.strip() != "", base_aimi["Email"]
)

tab_dashboard, tab_docs, tab_emails, tab_template = st.tabs(
    ["📊 Dashboard", "📎 Documentos", "✉️ Emails", "✏️ Template de Email"]
)

# --- Dashboard -----------------------------------------------------------------
with tab_dashboard:
    st.subheader("Estado dos Emails — AIMI")
    st.caption(
        "Vista rápida: para cada cliente, se o email do AIMI já foi enviado nos últimos anos. "
        "Para carregar documentos ou enviar emails, usa as abas 'Documentos' e 'Emails' (escolhe o ano no topo)."
    )

    with st.expander("✏️ Responsáveis (gestão do imóvel) — a quem vai o email de cada cliente"):
        st.caption(
            "Partilhado com o IMI — é o mesmo imóvel/gestor. Não altera o registo central de clientes. "
            "Quando preenches o Responsável (ex: uma agência de gestão), o email desse cliente passa a ir "
            "só para ele, em vez de ir para o proprietário. Deixa vazio para continuar a enviar ao próprio "
            "cliente, como sempre."
        )
        df_resp = base_aimi[["NIF", "Nome", "Responsavel_Nome", "Responsavel_Email"]].rename(
            columns={"Responsavel_Nome": "Responsável (nome)", "Responsavel_Email": "Responsável (email)"}
        )
        editado_resp = st.data_editor(
            df_resp, use_container_width=True, hide_index=True, height=360,
            disabled=["NIF", "Nome"], key="aimi_editor_responsaveis",
        )
        if st.button("💾 Guardar Responsáveis", key="aimi_guardar_responsaveis"):
            alterados_resp = 0
            for _, r in editado_resp.iterrows():
                nome_novo = (r["Responsável (nome)"] or "").strip()
                email_novo = (r["Responsável (email)"] or "").strip()
                anterior_resp = responsaveis_aimi.get(r["NIF"], {"nome": "", "email": ""})
                if nome_novo != anterior_resp["nome"] or email_novo != anterior_resp["email"]:
                    guardar_responsavel_imi_db(r["NIF"], nome_novo, email_novo)
                    alterados_resp += 1
            if alterados_resp:
                st.success(f"{alterados_resp} responsável(eis) atualizado(s).")
                st.rerun()
            else:
                st.info("Nenhuma alteração para guardar.")

    anos_dash = [ano_aimi, ano_aimi - 1, ano_aimi - 2]
    enviados_por_ano = {a: carregar_envios_db("aimi_dados", str(a)) for a in anos_dash}
    linhas_dash = []
    for _, r in base_aimi.iterrows():
        linha = {
            "N.º": r.get("Numero_Cliente", ""), "NIF": r["NIF"], "Nome": r["Nome"],
            "Responsável": r["Responsavel_Nome"] or "—",
            "Enviar para": r["Destinatario_Email"] or "⚠️ sem email",
        }
        for a in anos_dash:
            linha[str(a)] = "✅" if enviados_por_ano[a].get(r["NIF"], False) else "❌"
        linhas_dash.append(linha)
    st.dataframe(pd.DataFrame(linhas_dash), use_container_width=True, hide_index=True, height=460)

    cols_dash = st.columns(len(anos_dash))
    total_dash = len(base_aimi)
    for col, a in zip(cols_dash, anos_dash):
        col.metric(str(a), f"{sum(enviados_por_ano[a].get(n, False) for n in base_aimi['NIF'])} / {total_dash}")

# --- Documentos --------------------------------------------------------------
with tab_docs:
    st.subheader(f"Notas de cobrança — AIMI {ano_aimi}")
    st.caption("Ficam guardadas no arquivo persistente. Em massa (NIF no nome do ficheiro) ou cliente a cliente.")

    up_massa = st.file_uploader("Carregar notas de cobrança/guias PDF (nome com NIF de 9 dígitos)",
                                type=["pdf"], accept_multiple_files=True, key="aimi_up_massa")
    if up_massa:
        ids_upload = tuple(sorted(f"{f.name}_{f.size}" for f in up_massa))
        if st.session_state.get("_aimi_massa_proc") != (periodo, ids_upload):
            st.session_state["_aimi_massa_proc"] = (periodo, ids_upload)
            ok, sem_nif = 0, []
            for f in up_massa:
                nif_d = extrair_nif_de_filename(f.name)
                if nif_d:
                    storage_upload_pdf(f"aimi/{periodo}/guia/{nif_d}.pdf", f.getvalue())
                    ok += 1
                else:
                    sem_nif.append(f.name)
            msg = f"{ok} ficheiro(s) associados e guardados."
            if sem_nif:
                msg += f" Sem NIF no nome (usa o carregamento por cliente): {', '.join(sem_nif)}"
            st.success(msg)
            st.rerun()

    st.divider()
    st.markdown("**Carregamento por cliente** (inclui documentos extra)")
    nif_doc = st.selectbox(
        "Cliente",
        base_aimi["NIF"].tolist(),
        format_func=lambda n: f"{n} — {base_aimi.loc[base_aimi['NIF']==n,'Nome'].values[0]}",
        key="aimi_cliente_doc",
    )
    c1, c2 = st.columns(2)
    with c1:
        up_guia = st.file_uploader("Nota de cobrança / guia (PDF)", type=["pdf"], key=f"aimi_up_guia_{periodo}_{nif_doc}")
        if up_guia is not None:
            fid = f"{up_guia.name}_{up_guia.size}"
            if st.session_state.get(f"_aimi_guia_proc_{periodo}_{nif_doc}") != fid:
                storage_upload_pdf(f"aimi/{periodo}/guia/{nif_doc}.pdf", up_guia.getvalue())
                st.session_state[f"_aimi_guia_proc_{periodo}_{nif_doc}"] = fid
                guias_set.add(nif_doc)
        st.caption("✅ Nota no arquivo" if nif_doc in guias_set else "❌ Sem nota de cobrança")
    with c2:
        up_extras = st.file_uploader("Outros documentos (PDF, opcional)", type=["pdf"],
                                     accept_multiple_files=True, key=f"aimi_up_extra_{periodo}_{nif_doc}")
        if up_extras:
            ids_extras = tuple(sorted(f"{f.name}_{f.size}" for f in up_extras))
            if st.session_state.get(f"_aimi_extra_proc_{periodo}_{nif_doc}") != ids_extras:
                st.session_state[f"_aimi_extra_proc_{periodo}_{nif_doc}"] = ids_extras
                nomes_seguros = nomes_ficheiro_unicos([sanitizar_nome_ficheiro(f.name) for f in up_extras])
                for f, nome_seguro in zip(up_extras, nomes_seguros):
                    storage_upload_pdf(f"aimi/{periodo}/extra/{nif_doc}__{nome_seguro}", f.getvalue())
                    extras_dict.setdefault(nif_doc, []).append(nome_seguro)
                st.success(f"{len(up_extras)} documento(s) extra guardados.")
        n_extras = len(extras_dict.get(nif_doc, []))
        st.caption(f"📎 {n_extras} extra(s)" if n_extras else "Sem extras")

    st.divider()
    st.markdown("**Estado por cliente**")
    rows = []
    for _, r in base_aimi.iterrows():
        rows.append({
            "N.º": r.get("Numero_Cliente", ""), "NIF": r["NIF"], "Nome": r["Nome"],
            "Responsável": r["Responsavel_Nome"] or "—",
            "Nota de cobrança": "✅" if r["NIF"] in guias_set else "❌",
            "Extras": len(extras_dict.get(r["NIF"], [])),
            "Email Enviado": bool(r["Email_Enviado"]),
        })
    estado_df = pd.DataFrame(rows)
    editado = st.data_editor(
        estado_df,
        use_container_width=True, hide_index=True, height=360,
        disabled=["N.º", "NIF", "Nome", "Responsável", "Nota de cobrança", "Extras"],
        column_config={"Email Enviado": st.column_config.CheckboxColumn("Email Enviado")},
        key=f"aimi_estado_{periodo}",
    )
    if st.button("💾 Guardar piscos 'Email Enviado'", key="aimi_guardar_piscos"):
        for _, r in editado.iterrows():
            if bool(r["Email Enviado"]) != enviados.get(r["NIF"], False):
                marcar_envio_db("aimi_dados", r["NIF"], periodo, bool(r["Email Enviado"]))
        st.success("Estado guardado.")
        st.rerun()

    excel_aimi = gerar_excel_estado_mensal(
        f"Controlo AIMI {ano_aimi}", base_aimi, guias_set, set(), extras_dict, enviados,
        rotulo_decl="—",
    )
    st.download_button("⬇️ Descarregar Excel de Controlo (AIMI)", excel_aimi,
                       file_name=f"Controlo_AIMI_{periodo}_{date.today().isoformat()}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# --- Emails ------------------------------------------------------------------
with tab_emails:
    st.subheader(f"Enviar Emails — AIMI {ano_aimi}")

    elegiveis = base_aimi[base_aimi["Destinatario_Email"].str.strip() != ""].copy()
    sem_email = len(base_aimi) - len(elegiveis)
    if sem_email:
        st.caption(f"⚠️ {sem_email} cliente(s) sem email preenchido (nem do cliente, nem de Responsável) — não aparecem abaixo.")

    tpl = st.session_state.template_aimi

    com_docs = [n for n in elegiveis["NIF"] if n in guias_set or n in extras_dict]
    nao_enviados = [n for n in elegiveis["NIF"] if not enviados.get(n, False)]

    preview_nif = st.selectbox(
        "Pré-visualizar cliente:",
        elegiveis["NIF"].tolist(),
        format_func=lambda n: f"{n} — {elegiveis.loc[elegiveis['NIF']==n,'Nome'].values[0]}",
        key=f"aimi_preview_{periodo}",
    )

    def ctx_aimi(row):
        return {
            "ano": ano_aimi,
            "data_limite": data_limite_aimi(ano_aimi).strftime("%d/%m/%Y"),
        }

    if preview_nif:
        row = elegiveis[elegiveis["NIF"] == preview_nif].iloc[0]
        docs = docs_ss_cliente(periodo, preview_nif, guias_set, set(), extras_dict)
        assunto, corpo = render_template_docs(tpl, row, docs, ("nota de cobrança", "payment notice"), ctx_aimi(row))
        if row["Responsavel_Email"]:
            st.caption(f"📧 Vai para o Responsável: {row['Responsavel_Nome'] or ''} <{row['Responsavel_Email']}> (não vai para o proprietário)")
        else:
            st.caption(f"📧 Vai para o proprietário: {row['Nome']} <{row['Email']}>")
        st.text_input("Assunto (preview)", value=assunto, disabled=True)
        if row["Gestor_Email"]:
            st.caption(f"📋 CC: {row['Gestor_Nome'] or ''} <{row['Gestor_Email']}>  ·  Língua: {row['Lingua']}")
        else:
            st.caption(f"📋 CC: —  ·  Língua: {row['Lingua']}")
        st.text_area("Corpo (preview)", value=corpo, height=230, disabled=True)
        st.caption("📎 Anexos: " + (", ".join(docs) if docs else "nenhum documento carregado ainda"))

    st.divider()
    smtp_cfg = escolher_conta_email("aimi")

    st.markdown(f"📎 **{len(com_docs)} de {len(elegiveis)}** cliente(s) com documentos carregados este ano.")

    multiselect_key = f"aimi_selecionados_{periodo}"
    col_b1, col_b2, col_b3 = st.columns(3)
    with col_b1:
        if st.button("📎 Selecionar quem tem documentos e falta enviar", key="aimi_sel_docs"):
            st.session_state[multiselect_key] = [n for n in com_docs if n in nao_enviados]
            st.rerun()
    with col_b2:
        if st.button("☑️ Selecionar todos por enviar", key="aimi_sel_todos"):
            st.session_state[multiselect_key] = nao_enviados
            st.rerun()
    with col_b3:
        if st.button("✖️ Limpar seleção", key="aimi_sel_limpar"):
            st.session_state[multiselect_key] = []
            st.rerun()

    if multiselect_key not in st.session_state:
        st.session_state[multiselect_key] = [n for n in com_docs if n in nao_enviados]

    selecionados = st.multiselect(
        "Clientes selecionados para envio (para enviar só um, deixa só esse)",
        elegiveis["NIF"].tolist(),
        format_func=lambda n: f"{n} — {elegiveis.loc[elegiveis['NIF']==n,'Nome'].values[0]}"
        + ("" if n in com_docs else "  ⚠️ sem documentos")
        + ("  ✅ já enviado" if enviados.get(n, False) else ""),
        key=multiselect_key,
    )

    if st.button("🚀 Enviar Emails Selecionados", type="primary", disabled=not selecionados, key="aimi_enviar"):
        if not smtp_cfg["utilizador"] or not smtp_cfg["password"]:
            st.error("Escolhe ou cria uma conta de email com utilizador e password preenchidos.")
        else:
            progress = st.progress(0.0)
            status_box = st.empty()
            assinatura = st.session_state.params.get("assinatura_html", "")
            sucessos, falhas = 0, 0
            for i, nif in enumerate(selecionados):
                row = elegiveis[elegiveis["NIF"] == nif].iloc[0]
                docs = docs_ss_cliente(periodo, nif, guias_set, set(), extras_dict)
                assunto, corpo = render_template_docs(tpl, row, docs, ("nota de cobrança", "payment notice"), ctx_aimi(row))
                anexos = []
                if nif in guias_set:
                    conteudo = storage_download_pdf(f"aimi/{periodo}/guia/{nif}.pdf")
                    if conteudo:
                        anexos.append((f"AIMI_{periodo}_{nif}.pdf", conteudo))
                for nome_extra in extras_dict.get(nif, []):
                    conteudo = storage_download_pdf(f"aimi/{periodo}/extra/{nif}__{nome_extra}")
                    if conteudo:
                        anexos.append((nome_extra, conteudo))
                try:
                    cc_gestor = [row["Gestor_Email"]] if row["Gestor_Email"] else []
                    enviar_email(smtp_cfg, row["Destinatario_Email"], assunto, corpo, anexos, cc=cc_gestor,
                                 bcc=[smtp_cfg["remetente"]], assinatura_html=assinatura)
                    marcar_envio_db("aimi_dados", nif, periodo, True)
                    registar_log({
                        "data": datetime.now().strftime("%Y-%m-%d %H:%M"), "nif": nif,
                        "nome": row["Nome"], "pagamento": 0, "estado": f"Enviado ({periodo})",
                        "modulo": "AIMI", "enviado_por": meu_email(),
                    })
                    sucessos += 1
                except Exception as e:
                    registar_log({
                        "data": datetime.now().strftime("%Y-%m-%d %H:%M"), "nif": nif,
                        "nome": row["Nome"], "pagamento": 0, "estado": f"Erro ({periodo}): {e}",
                        "modulo": "AIMI", "enviado_por": meu_email(),
                    })
                    falhas += 1
                progress.progress((i + 1) / len(selecionados))
                status_box.text(f"{i+1}/{len(selecionados)} — {row['Nome']}")
            st.success(f"Concluído: {sucessos} enviados, {falhas} com erro. Estados guardados.")
            st.rerun()

# --- Template ----------------------------------------------------------------
with tab_template:
    st.subheader("Template do Email do AIMI")
    editor_template_bilingue(st.session_state.template_aimi, "aimi_tpl")
    st.caption("Placeholders disponíveis: {nome} {nif} {email} {ano} {data_limite} {lista_docs}. Alterações aqui ficam guardadas para toda a equipa.")

guardar_config_db(
    st.session_state.params, st.session_state.templates,
    st.session_state.get("template_irs"), st.session_state.get("template_ss"),
    {"iva": st.session_state.get("template_iva"), "imi": st.session_state.get("template_imi"),
     "aimi": st.session_state.get("template_aimi"), "info": st.session_state.get("template_info")},
)
