"""
Página do IVA — envio periódico da declaração e da guia de pagamento, com
documentos extra opcionais. Cada cliente tem regime Mensal ou Trimestral
(coluna 'Regime IVA' na página Clientes). Pagamento até dia 25 do 2.º mês
seguinte ao fim do período. Mesmo padrão da página da Segurança Social.
"""

from datetime import date, datetime

import pandas as pd
import streamlit as st

from common import (
    carregar_envios_db,
    data_limite_iva,
    docs_ss_cliente,
    editor_template_bilingue,
    enviar_email,
    escolher_conta_email,
    extrair_nif_de_filename,
    gerar_excel_estado_mensal,
    guardar_config_db,
    lista_periodos_iva,
    listar_extras_generico,
    marcar_envio_db,
    meu_email,
    montar_base_iva,
    nome_periodo_iva,
    nomes_ficheiro_unicos,
    registar_log,
    render_template_docs,
    sanitizar_nome_ficheiro,
    storage_download_pdf,
    storage_listar,
    storage_upload_pdf,
)

st.title("🧾 IVA — Declarações e Guias")
st.caption("SERVE — Contabilidade e Viabilização Empresarial")

regime = st.radio("Regime", ["Trimestral", "Mensal"], horizontal=True, key="iva_regime_escolhido",
                  help="O regime de cada cliente define-se na página 'Clientes' (coluna 'Regime IVA').")

base_iva = montar_base_iva(regime)
if base_iva.empty:
    st.info(f"Não há clientes com 'Aplica IVA' ligado e regime {regime}. Confirma os piscos e o 'Regime IVA' na página 'Clientes'.")
    st.stop()

periodos = lista_periodos_iva(regime, 12)
periodo = st.selectbox(
    "Período",
    periodos,
    index=1 if len(periodos) > 1 else 0,
    format_func=lambda p: f"{nome_periodo_iva(p)}  —  pagamento até {data_limite_iva(p).strftime('%d/%m/%Y')}",
    key=f"iva_periodo_{regime}",
)

guias_set = {n[:-4] for n in storage_listar(f"iva/{periodo}/guia") if n.lower().endswith(".pdf")}
decl_set = {n[:-4] for n in storage_listar(f"iva/{periodo}/decl") if n.lower().endswith(".pdf")}
extras_dict = listar_extras_generico(f"iva/{periodo}/extra")
enviados = carregar_envios_db("iva_dados", periodo)

base_iva = base_iva.reset_index(drop=True)
base_iva["Email_Enviado"] = base_iva["NIF"].map(lambda n: enviados.get(n, False))

tab_docs, tab_emails, tab_template = st.tabs(["📎 Documentos", "✉️ Emails", "✏️ Template de Email"])

# --- Documentos --------------------------------------------------------------
with tab_docs:
    st.subheader(f"Documentos de {nome_periodo_iva(periodo)}")
    st.caption("Ficam guardados no arquivo persistente. Em massa (NIF no nome do ficheiro) ou cliente a cliente.")

    st.markdown("**Carregamento em massa**")
    col_tipo, col_up = st.columns([1, 3])
    with col_tipo:
        tipo_doc = st.radio("Tipo de documento", ["Guia de pagamento", "Declaração periódica"], key="iva_tipo_doc")
    with col_up:
        up_massa = st.file_uploader("Carregar PDFs (nome com NIF de 9 dígitos)", type=["pdf"],
                                    accept_multiple_files=True, key="iva_up_massa")
    pasta_tipo = "guia" if tipo_doc == "Guia de pagamento" else "decl"
    if up_massa:
        ids_upload = tuple(sorted(f"{f.name}_{f.size}" for f in up_massa))
        if st.session_state.get("_iva_massa_proc") != (periodo, pasta_tipo, ids_upload):
            st.session_state["_iva_massa_proc"] = (periodo, pasta_tipo, ids_upload)
            ok, sem_nif = 0, []
            for f in up_massa:
                nif_d = extrair_nif_de_filename(f.name)
                if nif_d:
                    storage_upload_pdf(f"iva/{periodo}/{pasta_tipo}/{nif_d}.pdf", f.getvalue())
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
        base_iva["NIF"].tolist(),
        format_func=lambda n: f"{n} — {base_iva.loc[base_iva['NIF']==n,'Nome'].values[0]}",
        key="iva_cliente_doc",
    )
    c1, c2, c3 = st.columns(3)
    with c1:
        up_guia = st.file_uploader("Guia de pagamento (PDF)", type=["pdf"], key=f"iva_up_guia_{periodo}_{nif_doc}")
        if up_guia is not None:
            fid = f"{up_guia.name}_{up_guia.size}"
            if st.session_state.get(f"_iva_guia_proc_{periodo}_{nif_doc}") != fid:
                storage_upload_pdf(f"iva/{periodo}/guia/{nif_doc}.pdf", up_guia.getvalue())
                st.session_state[f"_iva_guia_proc_{periodo}_{nif_doc}"] = fid
                guias_set.add(nif_doc)
        st.caption("✅ Guia no arquivo" if nif_doc in guias_set else "❌ Sem guia")
    with c2:
        up_decl = st.file_uploader("Declaração periódica (PDF)", type=["pdf"], key=f"iva_up_decl_{periodo}_{nif_doc}")
        if up_decl is not None:
            fid = f"{up_decl.name}_{up_decl.size}"
            if st.session_state.get(f"_iva_decl_proc_{periodo}_{nif_doc}") != fid:
                storage_upload_pdf(f"iva/{periodo}/decl/{nif_doc}.pdf", up_decl.getvalue())
                st.session_state[f"_iva_decl_proc_{periodo}_{nif_doc}"] = fid
                decl_set.add(nif_doc)
        st.caption("✅ Declaração no arquivo" if nif_doc in decl_set else "❌ Sem declaração")
    with c3:
        up_extras = st.file_uploader("Outros documentos (PDF, opcional)", type=["pdf"],
                                     accept_multiple_files=True, key=f"iva_up_extra_{periodo}_{nif_doc}")
        if up_extras:
            ids_extras = tuple(sorted(f"{f.name}_{f.size}" for f in up_extras))
            if st.session_state.get(f"_iva_extra_proc_{periodo}_{nif_doc}") != ids_extras:
                st.session_state[f"_iva_extra_proc_{periodo}_{nif_doc}"] = ids_extras
                nomes_seguros = nomes_ficheiro_unicos([sanitizar_nome_ficheiro(f.name) for f in up_extras])
                for f, nome_seguro in zip(up_extras, nomes_seguros):
                    storage_upload_pdf(f"iva/{periodo}/extra/{nif_doc}__{nome_seguro}", f.getvalue())
                    extras_dict.setdefault(nif_doc, []).append(nome_seguro)
                st.success(f"{len(up_extras)} documento(s) extra guardados.")
        n_extras = len(extras_dict.get(nif_doc, []))
        st.caption(f"📎 {n_extras} extra(s)" if n_extras else "Sem extras")

    st.divider()
    st.markdown("**Estado do período por cliente**")
    rows = []
    for _, r in base_iva.iterrows():
        rows.append({
            "N.º": r.get("Numero_Cliente", ""), "NIF": r["NIF"], "Nome": r["Nome"],
            "Guia": "✅" if r["NIF"] in guias_set else "❌",
            "Declaração": "✅" if r["NIF"] in decl_set else "❌",
            "Extras": len(extras_dict.get(r["NIF"], [])),
            "Email Enviado": bool(r["Email_Enviado"]),
        })
    estado_df = pd.DataFrame(rows)
    editado = st.data_editor(
        estado_df,
        use_container_width=True, hide_index=True, height=360,
        disabled=["N.º", "NIF", "Nome", "Guia", "Declaração", "Extras"],
        column_config={"Email Enviado": st.column_config.CheckboxColumn("Email Enviado")},
        key=f"iva_estado_{periodo}",
    )
    if st.button("💾 Guardar piscos 'Email Enviado'", key="iva_guardar_piscos"):
        for _, r in editado.iterrows():
            if bool(r["Email Enviado"]) != enviados.get(r["NIF"], False):
                marcar_envio_db("iva_dados", r["NIF"], periodo, bool(r["Email Enviado"]))
        st.success("Estado guardado.")
        st.rerun()

    excel_iva = gerar_excel_estado_mensal(
        f"Controlo IVA — {nome_periodo_iva(periodo)}", base_iva, guias_set, decl_set, extras_dict, enviados,
        rotulo_decl="Declaração",
    )
    st.download_button("⬇️ Descarregar Excel de Controlo (IVA)", excel_iva,
                       file_name=f"Controlo_IVA_{periodo}_{date.today().isoformat()}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# --- Emails ------------------------------------------------------------------
with tab_emails:
    st.subheader(f"Enviar Emails — {nome_periodo_iva(periodo)}")

    elegiveis = base_iva[base_iva["Email"].str.strip() != ""].copy()
    sem_email = len(base_iva) - len(elegiveis)
    if sem_email:
        st.caption(f"⚠️ {sem_email} cliente(s) sem email preenchido — não aparecem abaixo.")

    tpl = st.session_state.template_iva
    ctx_extra = {"periodo_nome": None, "data_limite": data_limite_iva(periodo).strftime("%d/%m/%Y")}

    com_docs = [n for n in elegiveis["NIF"] if n in guias_set or n in decl_set or n in extras_dict]
    nao_enviados = [n for n in elegiveis["NIF"] if not enviados.get(n, False)]

    preview_nif = st.selectbox(
        "Pré-visualizar cliente:",
        elegiveis["NIF"].tolist(),
        format_func=lambda n: f"{n} — {elegiveis.loc[elegiveis['NIF']==n,'Nome'].values[0]}",
        key=f"iva_preview_{periodo}",
    )
    if preview_nif:
        row = elegiveis[elegiveis["NIF"] == preview_nif].iloc[0]
        docs = docs_ss_cliente(periodo, preview_nif, guias_set, decl_set, extras_dict)
        lingua_row = row["Lingua"]
        ctx = {"periodo_nome": nome_periodo_iva(periodo, lingua_row), "data_limite": data_limite_iva(periodo).strftime("%d/%m/%Y")}
        assunto, corpo = render_template_docs(tpl, row, docs, ("declaração periódica", "periodic VAT return"), ctx)
        st.text_input("Assunto (preview)", value=assunto, disabled=True)
        if row["Gestor_Email"]:
            st.caption(f"📋 CC: {row['Gestor_Nome'] or ''} <{row['Gestor_Email']}>  ·  Língua: {row['Lingua']}")
        else:
            st.caption(f"📋 CC: —  ·  Língua: {row['Lingua']}")
        st.text_area("Corpo (preview)", value=corpo, height=230, disabled=True)
        st.caption("📎 Anexos: " + (", ".join(docs) if docs else "nenhum documento carregado ainda"))

    st.divider()
    smtp_cfg = escolher_conta_email("iva")

    st.markdown(f"📎 **{len(com_docs)} de {len(elegiveis)}** cliente(s) com documentos carregados neste período.")

    multiselect_key = f"iva_selecionados_{periodo}"
    col_b1, col_b2, col_b3 = st.columns(3)
    with col_b1:
        if st.button("📎 Selecionar quem tem documentos e falta enviar", key="iva_sel_docs"):
            st.session_state[multiselect_key] = [n for n in com_docs if n in nao_enviados]
            st.rerun()
    with col_b2:
        if st.button("☑️ Selecionar todos por enviar", key="iva_sel_todos"):
            st.session_state[multiselect_key] = nao_enviados
            st.rerun()
    with col_b3:
        if st.button("✖️ Limpar seleção", key="iva_sel_limpar"):
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

    if st.button("🚀 Enviar Emails Selecionados", type="primary", disabled=not selecionados, key="iva_enviar"):
        if not smtp_cfg["utilizador"] or not smtp_cfg["password"]:
            st.error("Escolhe ou cria uma conta de email com utilizador e password preenchidos.")
        else:
            progress = st.progress(0.0)
            status_box = st.empty()
            assinatura = st.session_state.params.get("assinatura_html", "")
            sucessos, falhas = 0, 0
            for i, nif in enumerate(selecionados):
                row = elegiveis[elegiveis["NIF"] == nif].iloc[0]
                docs = docs_ss_cliente(periodo, nif, guias_set, decl_set, extras_dict)
                ctx = {"periodo_nome": nome_periodo_iva(periodo, row["Lingua"]), "data_limite": data_limite_iva(periodo).strftime("%d/%m/%Y")}
                assunto, corpo = render_template_docs(tpl, row, docs, ("declaração periódica", "periodic VAT return"), ctx)
                anexos = []
                if nif in guias_set:
                    conteudo = storage_download_pdf(f"iva/{periodo}/guia/{nif}.pdf")
                    if conteudo:
                        anexos.append((f"Guia_IVA_{periodo}_{nif}.pdf", conteudo))
                if nif in decl_set:
                    conteudo = storage_download_pdf(f"iva/{periodo}/decl/{nif}.pdf")
                    if conteudo:
                        anexos.append((f"Declaracao_IVA_{periodo}_{nif}.pdf", conteudo))
                for nome_extra in extras_dict.get(nif, []):
                    conteudo = storage_download_pdf(f"iva/{periodo}/extra/{nif}__{nome_extra}")
                    if conteudo:
                        anexos.append((nome_extra, conteudo))
                try:
                    cc_gestor = [row["Gestor_Email"]] if row["Gestor_Email"] else []
                    enviar_email(smtp_cfg, row["Email"], assunto, corpo, anexos, cc=cc_gestor, assinatura_html=assinatura)
                    marcar_envio_db("iva_dados", nif, periodo, True)
                    registar_log({
                        "data": datetime.now().strftime("%Y-%m-%d %H:%M"), "nif": nif,
                        "nome": row["Nome"], "pagamento": 0, "estado": f"Enviado ({periodo})",
                        "modulo": "IVA", "enviado_por": meu_email(),
                    })
                    sucessos += 1
                except Exception as e:
                    registar_log({
                        "data": datetime.now().strftime("%Y-%m-%d %H:%M"), "nif": nif,
                        "nome": row["Nome"], "pagamento": 0, "estado": f"Erro ({periodo}): {e}",
                        "modulo": "IVA", "enviado_por": meu_email(),
                    })
                    falhas += 1
                progress.progress((i + 1) / len(selecionados))
                status_box.text(f"{i+1}/{len(selecionados)} — {row['Nome']}")
            st.success(f"Concluído: {sucessos} enviados, {falhas} com erro. Estados guardados.")
            st.rerun()

# --- Template ----------------------------------------------------------------
with tab_template:
    st.subheader("Template do Email do IVA")
    editor_template_bilingue(st.session_state.template_iva, "iva_tpl")
    st.caption("Placeholders disponíveis: {nome} {nif} {email} {periodo_nome} {data_limite} {lista_docs}. Alterações aqui ficam guardadas para toda a equipa.")

guardar_config_db(
    st.session_state.params, st.session_state.templates,
    st.session_state.get("template_irs"), st.session_state.get("template_ss"),
    {"iva": st.session_state.get("template_iva"), "imi": st.session_state.get("template_imi")},
)
