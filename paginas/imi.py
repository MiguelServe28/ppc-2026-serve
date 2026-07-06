"""
Página do IMI — envio das notas de cobrança/guias por prestação (31 de maio,
31 de agosto, 30 de novembro), com documentos extra opcionais. Entram os
clientes com o pisco 'IMI'. Mesmo padrão da Segurança Social/IVA.
"""

from datetime import date, datetime

import pandas as pd
import streamlit as st

from common import (
    PRESTACOES_IMI,
    PRESTACOES_IMI_EN,
    carregar_envios_db,
    data_limite_imi,
    docs_ss_cliente,
    editor_template_bilingue,
    enviar_email,
    escolher_conta_email,
    extrair_nif_de_filename,
    gerar_excel_estado_mensal,
    guardar_config_db,
    listar_extras_generico,
    marcar_envio_db,
    meu_email,
    montar_base_imi,
    registar_log,
    render_template_docs,
    sou_admin,
    storage_download_pdf,
    storage_listar,
    storage_upload_pdf,
)

st.title("🏠 IMI — Notas de Cobrança")
st.caption("SERVE — Contabilidade e Viabilização Empresarial")

base_imi = montar_base_imi()
if base_imi.empty:
    st.info("Não há clientes com o pisco 'IMI' ligado — ativa-o na página 'Clientes'.")
    st.stop()

col_ano, col_prest = st.columns(2)
with col_ano:
    ano_imi = st.selectbox("Ano", list(range(date.today().year, date.today().year - 4, -1)), key="imi_ano")
with col_prest:
    prestacao = st.selectbox(
        "Prestação",
        [1, 2, 3],
        format_func=lambda p: f"{PRESTACOES_IMI[p]} — até {data_limite_imi(ano_imi, p).strftime('%d/%m/%Y')}",
        key="imi_prestacao",
    )
periodo = f"{ano_imi}-P{prestacao}"

guias_set = {n[:-4] for n in storage_listar(f"imi/{periodo}/guia") if n.lower().endswith(".pdf")}
extras_dict = listar_extras_generico(f"imi/{periodo}/extra")
enviados = carregar_envios_db("imi_dados", periodo)

base_imi = base_imi.reset_index(drop=True)
base_imi["Email_Enviado"] = base_imi["NIF"].map(lambda n: enviados.get(n, False))

tab_docs, tab_emails, tab_template = st.tabs(["📎 Documentos", "✉️ Emails", "✏️ Template de Email"])

# --- Documentos --------------------------------------------------------------
with tab_docs:
    st.subheader(f"Notas de cobrança — {PRESTACOES_IMI[prestacao]} de {ano_imi}")
    st.caption("Ficam guardadas no arquivo persistente. Em massa (NIF no nome do ficheiro) ou cliente a cliente.")

    up_massa = st.file_uploader("Carregar notas de cobrança/guias PDF (nome com NIF de 9 dígitos)",
                                type=["pdf"], accept_multiple_files=True, key="imi_up_massa")
    if up_massa:
        ids_upload = tuple(sorted(f"{f.name}_{f.size}" for f in up_massa))
        if st.session_state.get("_imi_massa_proc") != (periodo, ids_upload):
            st.session_state["_imi_massa_proc"] = (periodo, ids_upload)
            ok, sem_nif = 0, []
            for f in up_massa:
                nif_d = extrair_nif_de_filename(f.name)
                if nif_d:
                    storage_upload_pdf(f"imi/{periodo}/guia/{nif_d}.pdf", f.getvalue())
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
        base_imi["NIF"].tolist(),
        format_func=lambda n: f"{n} — {base_imi.loc[base_imi['NIF']==n,'Nome'].values[0]}",
        key="imi_cliente_doc",
    )
    c1, c2 = st.columns(2)
    with c1:
        up_guia = st.file_uploader("Nota de cobrança / guia (PDF)", type=["pdf"], key=f"imi_up_guia_{periodo}_{nif_doc}")
        if up_guia is not None:
            fid = f"{up_guia.name}_{up_guia.size}"
            if st.session_state.get(f"_imi_guia_proc_{periodo}_{nif_doc}") != fid:
                storage_upload_pdf(f"imi/{periodo}/guia/{nif_doc}.pdf", up_guia.getvalue())
                st.session_state[f"_imi_guia_proc_{periodo}_{nif_doc}"] = fid
                guias_set.add(nif_doc)
        st.caption("✅ Nota no arquivo" if nif_doc in guias_set else "❌ Sem nota de cobrança")
    with c2:
        up_extras = st.file_uploader("Outros documentos (PDF, opcional)", type=["pdf"],
                                     accept_multiple_files=True, key=f"imi_up_extra_{periodo}_{nif_doc}")
        if up_extras:
            ids_extras = tuple(sorted(f"{f.name}_{f.size}" for f in up_extras))
            if st.session_state.get(f"_imi_extra_proc_{periodo}_{nif_doc}") != ids_extras:
                st.session_state[f"_imi_extra_proc_{periodo}_{nif_doc}"] = ids_extras
                for f in up_extras:
                    storage_upload_pdf(f"imi/{periodo}/extra/{nif_doc}__{f.name}", f.getvalue())
                    extras_dict.setdefault(nif_doc, []).append(f.name)
                st.success(f"{len(up_extras)} documento(s) extra guardados.")
        n_extras = len(extras_dict.get(nif_doc, []))
        st.caption(f"📎 {n_extras} extra(s)" if n_extras else "Sem extras")

    st.divider()
    st.markdown("**Estado da prestação por cliente**")
    rows = []
    for _, r in base_imi.iterrows():
        rows.append({
            "N.º": r.get("Numero_Cliente", ""), "NIF": r["NIF"], "Nome": r["Nome"],
            "Nota de cobrança": "✅" if r["NIF"] in guias_set else "❌",
            "Extras": len(extras_dict.get(r["NIF"], [])),
            "Email Enviado": bool(r["Email_Enviado"]),
        })
    estado_df = pd.DataFrame(rows)
    editado = st.data_editor(
        estado_df,
        use_container_width=True, hide_index=True, height=360,
        disabled=["N.º", "NIF", "Nome", "Nota de cobrança", "Extras"],
        column_config={"Email Enviado": st.column_config.CheckboxColumn("Email Enviado")},
        key=f"imi_estado_{periodo}",
    )
    if st.button("💾 Guardar piscos 'Email Enviado'", key="imi_guardar_piscos"):
        for _, r in editado.iterrows():
            if bool(r["Email Enviado"]) != enviados.get(r["NIF"], False):
                marcar_envio_db("imi_dados", r["NIF"], periodo, bool(r["Email Enviado"]))
        st.success("Estado guardado.")
        st.rerun()

    excel_imi = gerar_excel_estado_mensal(
        f"Controlo IMI {ano_imi} — {PRESTACOES_IMI[prestacao]}", base_imi, guias_set, set(), extras_dict, enviados,
        rotulo_decl="—",
    )
    st.download_button("⬇️ Descarregar Excel de Controlo (IMI)", excel_imi,
                       file_name=f"Controlo_IMI_{periodo}_{date.today().isoformat()}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# --- Emails ------------------------------------------------------------------
with tab_emails:
    st.subheader(f"Enviar Emails — {PRESTACOES_IMI[prestacao]} de {ano_imi}")

    elegiveis = base_imi[base_imi["Email"].str.strip() != ""].copy()
    sem_email = len(base_imi) - len(elegiveis)
    if sem_email:
        st.caption(f"⚠️ {sem_email} cliente(s) sem email preenchido — não aparecem abaixo.")

    tpl = st.session_state.template_imi

    com_docs = [n for n in elegiveis["NIF"] if n in guias_set or n in extras_dict]
    nao_enviados = [n for n in elegiveis["NIF"] if not enviados.get(n, False)]

    preview_nif = st.selectbox(
        "Pré-visualizar cliente:",
        elegiveis["NIF"].tolist(),
        format_func=lambda n: f"{n} — {elegiveis.loc[elegiveis['NIF']==n,'Nome'].values[0]}",
        key=f"imi_preview_{periodo}",
    )

    def ctx_imi(row):
        pt = row["Lingua"] != "EN"
        return {
            "ano": ano_imi,
            "prestacao": PRESTACOES_IMI[prestacao] if pt else PRESTACOES_IMI_EN[prestacao],
            "data_limite": data_limite_imi(ano_imi, prestacao).strftime("%d/%m/%Y"),
        }

    if preview_nif:
        row = elegiveis[elegiveis["NIF"] == preview_nif].iloc[0]
        docs = docs_ss_cliente(periodo, preview_nif, guias_set, set(), extras_dict)
        assunto, corpo = render_template_docs(tpl, row, docs, ("nota de cobrança", "payment notice"), ctx_imi(row))
        st.text_input("Assunto (preview)", value=assunto, disabled=True)
        if row["Gestor_Email"]:
            st.caption(f"📋 CC: {row['Gestor_Nome'] or ''} <{row['Gestor_Email']}>  ·  Língua: {row['Lingua']}")
        else:
            st.caption(f"📋 CC: —  ·  Língua: {row['Lingua']}")
        st.text_area("Corpo (preview)", value=corpo, height=230, disabled=True)
        st.caption("📎 Anexos: " + (", ".join(docs) if docs else "nenhum documento carregado ainda"))

    st.divider()
    smtp_cfg = escolher_conta_email("imi")

    st.markdown(f"📎 **{len(com_docs)} de {len(elegiveis)}** cliente(s) com documentos carregados nesta prestação.")

    multiselect_key = f"imi_selecionados_{periodo}"
    col_b1, col_b2, col_b3 = st.columns(3)
    with col_b1:
        if st.button("📎 Selecionar quem tem documentos e falta enviar", key="imi_sel_docs"):
            st.session_state[multiselect_key] = [n for n in com_docs if n in nao_enviados]
            st.rerun()
    with col_b2:
        if st.button("☑️ Selecionar todos por enviar", key="imi_sel_todos"):
            st.session_state[multiselect_key] = nao_enviados
            st.rerun()
    with col_b3:
        if st.button("✖️ Limpar seleção", key="imi_sel_limpar"):
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

    if st.button("🚀 Enviar Emails Selecionados", type="primary", disabled=not selecionados, key="imi_enviar"):
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
                assunto, corpo = render_template_docs(tpl, row, docs, ("nota de cobrança", "payment notice"), ctx_imi(row))
                anexos = []
                if nif in guias_set:
                    conteudo = storage_download_pdf(f"imi/{periodo}/guia/{nif}.pdf")
                    if conteudo:
                        anexos.append((f"IMI_{periodo}_{nif}.pdf", conteudo))
                for nome_extra in extras_dict.get(nif, []):
                    conteudo = storage_download_pdf(f"imi/{periodo}/extra/{nif}__{nome_extra}")
                    if conteudo:
                        anexos.append((nome_extra, conteudo))
                try:
                    cc_gestor = [row["Gestor_Email"]] if row["Gestor_Email"] else []
                    enviar_email(smtp_cfg, row["Email"], assunto, corpo, anexos, cc=cc_gestor, assinatura_html=assinatura)
                    marcar_envio_db("imi_dados", nif, periodo, True)
                    registar_log({
                        "data": datetime.now().strftime("%Y-%m-%d %H:%M"), "nif": nif,
                        "nome": row["Nome"], "pagamento": prestacao, "estado": f"Enviado ({periodo})",
                        "modulo": "IMI", "enviado_por": meu_email(),
                    })
                    sucessos += 1
                except Exception as e:
                    registar_log({
                        "data": datetime.now().strftime("%Y-%m-%d %H:%M"), "nif": nif,
                        "nome": row["Nome"], "pagamento": prestacao, "estado": f"Erro ({periodo}): {e}",
                        "modulo": "IMI", "enviado_por": meu_email(),
                    })
                    falhas += 1
                progress.progress((i + 1) / len(selecionados))
                status_box.text(f"{i+1}/{len(selecionados)} — {row['Nome']}")
            st.success(f"Concluído: {sucessos} enviados, {falhas} com erro. Estados guardados.")
            st.rerun()

# --- Template ----------------------------------------------------------------
with tab_template:
    st.subheader("Template do Email do IMI")
    if sou_admin():
        editor_template_bilingue(st.session_state.template_imi, "imi_tpl")
        st.caption("Placeholders disponíveis: {nome} {nif} {email} {ano} {prestacao} {data_limite} {lista_docs}.")
    else:
        st.caption("O template de email é definido pelo administrador.")
        st.text_input("Assunto (PT)", value=st.session_state.template_imi.get("assunto", ""), disabled=True)
        st.text_area("Corpo (PT)", value=st.session_state.template_imi.get("corpo", ""), height=260, disabled=True)

guardar_config_db(
    st.session_state.params, st.session_state.templates,
    st.session_state.get("template_irs"), st.session_state.get("template_ss"),
    {"iva": st.session_state.get("template_iva"), "imi": st.session_state.get("template_imi")},
)
