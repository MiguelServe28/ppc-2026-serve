"""
Página da Segurança Social — envio mensal das DMR (empresas) / DRI e guias de
pagamento (trabalhadores independentes), com possibilidade de anexar outros
documentos avulsos. Tudo organizado por mês de referência: escolhe-se o mês,
carregam-se os documentos (ficam no arquivo persistente) e enviam-se os emails.
Pagamento até dia 25 do mês seguinte ao mês de referência.
"""

from datetime import datetime

import pandas as pd
import streamlit as st

from common import (
    carregar_ss_mes_db,
    data_limite_ss,
    docs_ss_cliente,
    editor_template_bilingue,
    enviar_email,
    escolher_conta_email,
    extrair_nif_de_filename,
    guardar_config_db,
    lista_meses_ss,
    listar_extras_ss,
    marcar_ss_enviado_db,
    meu_email,
    montar_base_ss,
    nome_mes,
    registar_log,
    render_template_ss,
    sou_admin,
    storage_download_pdf,
    storage_listar,
    storage_upload_pdf,
)

st.title("🏛️ Segurança Social — DMR / DRI")
st.caption("SERVE — Contabilidade e Viabilização Empresarial")

base_ss = montar_base_ss()
if base_ss.empty:
    st.info("Ainda não há clientes com o pisco 'Seg. Social' ligado — ativa-o na página 'Clientes'.")
    st.stop()

# --- Seletor do mês de referência -------------------------------------------
meses = lista_meses_ss(18)
mes = st.selectbox(
    "Mês de referência (remunerações)",
    meses,
    index=1 if len(meses) > 1 else 0,  # por omissão, o mês anterior ao atual
    format_func=lambda m: f"{nome_mes(m)}  —  pagamento até {data_limite_ss(m).strftime('%d/%m/%Y')}",
    key="ss_mes",
)

# Estado deste mês (uma leitura por tipo de documento + estado de envio)
guias_set = {n[:-4] for n in storage_listar(f"ss/{mes}/guia") if n.lower().endswith(".pdf")}
dmrs_set = {n[:-4] for n in storage_listar(f"ss/{mes}/dmr") if n.lower().endswith(".pdf")}
extras_dict = listar_extras_ss(mes)
enviados = carregar_ss_mes_db(mes)

base_ss = base_ss.reset_index(drop=True)
base_ss["Email_Enviado"] = base_ss["NIF"].map(lambda n: enviados.get(n, False))

tab_docs, tab_emails, tab_template = st.tabs(["📎 Documentos", "✉️ Emails", "✏️ Template de Email"])

# --- Documentos --------------------------------------------------------------
with tab_docs:
    st.subheader(f"Documentos de {nome_mes(mes)}")
    st.caption(
        "Os documentos ficam guardados no arquivo persistente — não se perdem ao fechar o browser. "
        "Podes carregar em massa (o NIF no nome do ficheiro associa sozinho) ou cliente a cliente."
    )

    st.markdown("**Carregamento em massa** (vários PDFs de uma vez, com o NIF no nome do ficheiro)")
    col_tipo, col_up = st.columns([1, 3])
    with col_tipo:
        tipo_doc = st.radio("Tipo de documento", ["Guia de pagamento", "DMR / DRI"], key="ss_tipo_doc")
    with col_up:
        up_massa = st.file_uploader(
            "Carregar PDFs (nome com NIF de 9 dígitos)",
            type=["pdf"], accept_multiple_files=True, key="ss_up_massa",
        )
    pasta_tipo = "guia" if tipo_doc == "Guia de pagamento" else "dmr"
    if up_massa:
        ids_upload = tuple(sorted(f"{f.name}_{f.size}" for f in up_massa))
        if st.session_state.get("_ss_massa_proc") != (mes, pasta_tipo, ids_upload):
            st.session_state["_ss_massa_proc"] = (mes, pasta_tipo, ids_upload)
            ok, sem_nif = 0, []
            for f in up_massa:
                nif_d = extrair_nif_de_filename(f.name)
                if nif_d:
                    storage_upload_pdf(f"ss/{mes}/{pasta_tipo}/{nif_d}.pdf", f.getvalue())
                    ok += 1
                else:
                    sem_nif.append(f.name)
            msg = f"{ok} ficheiro(s) associados e guardados no arquivo."
            if sem_nif:
                msg += f" Sem NIF no nome (usa o carregamento por cliente, abaixo): {', '.join(sem_nif)}"
            st.success(msg)
            st.rerun()

    st.divider()
    st.markdown("**Carregamento por cliente** (inclui outros documentos avulsos)")
    nif_doc = st.selectbox(
        "Cliente",
        base_ss["NIF"].tolist(),
        format_func=lambda n: f"{n} — {base_ss.loc[base_ss['NIF']==n,'Nome'].values[0]}",
        key="ss_cliente_doc",
    )
    c1, c2, c3 = st.columns(3)
    with c1:
        up_guia = st.file_uploader("Guia de pagamento (PDF)", type=["pdf"], key=f"ss_up_guia_{mes}_{nif_doc}")
        if up_guia is not None:
            fid = f"{up_guia.name}_{up_guia.size}"
            if st.session_state.get(f"_ss_guia_proc_{mes}_{nif_doc}") != fid:
                storage_upload_pdf(f"ss/{mes}/guia/{nif_doc}.pdf", up_guia.getvalue())
                st.session_state[f"_ss_guia_proc_{mes}_{nif_doc}"] = fid
                guias_set.add(nif_doc)
        st.caption("✅ Guia no arquivo" if nif_doc in guias_set else "❌ Sem guia")
    with c2:
        up_dmr = st.file_uploader("DMR / DRI (PDF)", type=["pdf"], key=f"ss_up_dmr_{mes}_{nif_doc}")
        if up_dmr is not None:
            fid = f"{up_dmr.name}_{up_dmr.size}"
            if st.session_state.get(f"_ss_dmr_proc_{mes}_{nif_doc}") != fid:
                storage_upload_pdf(f"ss/{mes}/dmr/{nif_doc}.pdf", up_dmr.getvalue())
                st.session_state[f"_ss_dmr_proc_{mes}_{nif_doc}"] = fid
                dmrs_set.add(nif_doc)
        st.caption("✅ DMR/DRI no arquivo" if nif_doc in dmrs_set else "❌ Sem DMR/DRI")
    with c3:
        up_extras = st.file_uploader(
            "Outros documentos (PDF, opcional)", type=["pdf"],
            accept_multiple_files=True, key=f"ss_up_extra_{mes}_{nif_doc}",
            help="Ex: outros pagamentos que queiras aproveitar para enviar no mesmo email deste mês.",
        )
        if up_extras:
            ids_extras = tuple(sorted(f"{f.name}_{f.size}" for f in up_extras))
            if st.session_state.get(f"_ss_extra_proc_{mes}_{nif_doc}") != ids_extras:
                st.session_state[f"_ss_extra_proc_{mes}_{nif_doc}"] = ids_extras
                for f in up_extras:
                    storage_upload_pdf(f"ss/{mes}/extra/{nif_doc}__{f.name}", f.getvalue())
                    extras_dict.setdefault(nif_doc, []).append(f.name)
                st.success(f"{len(up_extras)} documento(s) extra guardados.")
        n_extras = len(extras_dict.get(nif_doc, []))
        st.caption(f"📎 {n_extras} extra(s) no arquivo" if n_extras else "Sem extras")

    st.divider()
    st.markdown("**Estado do mês por cliente**")
    rows = []
    for _, r in base_ss.iterrows():
        rows.append({
            "N.º": r.get("Numero_Cliente", ""), "NIF": r["NIF"], "Nome": r["Nome"],
            "Guia": "✅" if r["NIF"] in guias_set else "❌",
            "DMR/DRI": "✅" if r["NIF"] in dmrs_set else "❌",
            "Extras": len(extras_dict.get(r["NIF"], [])),
            "Email Enviado": bool(r["Email_Enviado"]),
        })
    estado_df = pd.DataFrame(rows)
    editado = st.data_editor(
        estado_df,
        use_container_width=True,
        hide_index=True,
        height=360,
        disabled=["N.º", "NIF", "Nome", "Guia", "DMR/DRI", "Extras"],
        column_config={"Email Enviado": st.column_config.CheckboxColumn("Email Enviado")},
        key=f"ss_estado_{mes}",
    )
    if st.button("💾 Guardar piscos 'Email Enviado'"):
        for _, r in editado.iterrows():
            if bool(r["Email Enviado"]) != enviados.get(r["NIF"], False):
                marcar_ss_enviado_db(r["NIF"], mes, bool(r["Email Enviado"]))
        st.success("Estado guardado.")
        st.rerun()

# --- Emails ------------------------------------------------------------------
with tab_emails:
    st.subheader(f"Enviar Emails — {nome_mes(mes)}")

    elegiveis = base_ss[base_ss["Email"].str.strip() != ""].copy()
    sem_email = len(base_ss) - len(elegiveis)
    if sem_email:
        st.caption(f"⚠️ {sem_email} cliente(s) sem email preenchido no registo central — não aparecem abaixo.")

    tpl = st.session_state.template_ss

    com_docs = [n for n in elegiveis["NIF"] if n in guias_set or n in dmrs_set or n in extras_dict]
    nao_enviados = [n for n in elegiveis["NIF"] if not enviados.get(n, False)]

    preview_nif = st.selectbox(
        "Pré-visualizar cliente:",
        elegiveis["NIF"].tolist(),
        format_func=lambda n: f"{n} — {elegiveis.loc[elegiveis['NIF']==n,'Nome'].values[0]}",
        key=f"ss_preview_{mes}",
    )
    if preview_nif:
        row = elegiveis[elegiveis["NIF"] == preview_nif].iloc[0]
        docs = docs_ss_cliente(mes, preview_nif, guias_set, dmrs_set, extras_dict)
        assunto, corpo = render_template_ss(tpl, row, mes, docs)
        st.text_input("Assunto (preview)", value=assunto, disabled=True)
        if row["Gestor_Email"]:
            st.caption(f"📋 CC: {row['Gestor_Nome'] or ''} <{row['Gestor_Email']}>  ·  Língua: {row['Lingua']}")
        else:
            st.caption(f"📋 CC: — (sem gestor definido)  ·  Língua: {row['Lingua']}")
        st.text_area("Corpo (preview)", value=corpo, height=230, disabled=True)
        st.caption("📎 Anexos: " + (", ".join(docs) if docs else "nenhum documento carregado ainda"))

        docs_prev = docs
        ja_enviado_prev = enviados.get(preview_nif, False)

    st.divider()
    smtp_cfg = escolher_conta_email("ss")

    st.markdown(f"📎 **{len(com_docs)} de {len(elegiveis)}** cliente(s) já têm documentos carregados para este mês.")

    multiselect_key = f"ss_selecionados_{mes}"
    col_b1, col_b2, col_b3 = st.columns(3)
    with col_b1:
        if st.button("📎 Selecionar quem tem documentos e falta enviar"):
            st.session_state[multiselect_key] = [n for n in com_docs if n in nao_enviados]
            st.rerun()
    with col_b2:
        if st.button("☑️ Selecionar todos por enviar"):
            st.session_state[multiselect_key] = nao_enviados
            st.rerun()
    with col_b3:
        if st.button("✖️ Limpar seleção"):
            st.session_state[multiselect_key] = []
            st.rerun()

    if multiselect_key not in st.session_state:
        st.session_state[multiselect_key] = [n for n in com_docs if n in nao_enviados]

    selecionados = st.multiselect(
        "Clientes selecionados para envio (podes ajustar — para enviar só um, deixa só esse)",
        elegiveis["NIF"].tolist(),
        format_func=lambda n: f"{n} — {elegiveis.loc[elegiveis['NIF']==n,'Nome'].values[0]}"
        + ("" if n in com_docs else "  ⚠️ sem documentos")
        + ("  ✅ já enviado" if enviados.get(n, False) else ""),
        key=multiselect_key,
    )

    if st.button("🚀 Enviar Emails Selecionados", type="primary", disabled=not selecionados):
        if not smtp_cfg["utilizador"] or not smtp_cfg["password"]:
            st.error("Escolhe ou cria uma conta de email com utilizador e password preenchidos.")
        else:
            progress = st.progress(0.0)
            status_box = st.empty()
            assinatura = st.session_state.params.get("assinatura_html", "")
            sucessos, falhas = 0, 0
            for i, nif in enumerate(selecionados):
                row = elegiveis[elegiveis["NIF"] == nif].iloc[0]
                docs = docs_ss_cliente(mes, nif, guias_set, dmrs_set, extras_dict)
                assunto, corpo = render_template_ss(tpl, row, mes, docs)
                anexos = []
                if nif in guias_set:
                    conteudo = storage_download_pdf(f"ss/{mes}/guia/{nif}.pdf")
                    if conteudo:
                        anexos.append((f"Guia_SS_{mes}_{nif}.pdf", conteudo))
                if nif in dmrs_set:
                    conteudo = storage_download_pdf(f"ss/{mes}/dmr/{nif}.pdf")
                    if conteudo:
                        anexos.append((f"DMR_{mes}_{nif}.pdf", conteudo))
                for nome_extra in extras_dict.get(nif, []):
                    conteudo = storage_download_pdf(f"ss/{mes}/extra/{nif}__{nome_extra}")
                    if conteudo:
                        anexos.append((nome_extra, conteudo))
                try:
                    cc_gestor = [row["Gestor_Email"]] if row["Gestor_Email"] else []
                    enviar_email(smtp_cfg, row["Email"], assunto, corpo, anexos, cc=cc_gestor, assinatura_html=assinatura)
                    marcar_ss_enviado_db(nif, mes, True)
                    registar_log({
                        "data": datetime.now().strftime("%Y-%m-%d %H:%M"), "nif": nif,
                        "nome": row["Nome"], "pagamento": 0, "estado": f"Enviado ({mes})",
                        "modulo": "SS", "enviado_por": meu_email(),
                    })
                    sucessos += 1
                except Exception as e:
                    registar_log({
                        "data": datetime.now().strftime("%Y-%m-%d %H:%M"), "nif": nif,
                        "nome": row["Nome"], "pagamento": 0, "estado": f"Erro ({mes}): {e}",
                        "modulo": "SS", "enviado_por": meu_email(),
                    })
                    falhas += 1
                progress.progress((i + 1) / len(selecionados))
                status_box.text(f"{i+1}/{len(selecionados)} — {row['Nome']}")
            st.success(f"Concluído: {sucessos} enviados, {falhas} com erro. Estados guardados.")
            st.rerun()

    if st.session_state.log_envio:
        st.markdown("### Log de Envios")
        log_df = pd.DataFrame(st.session_state.log_envio)
        if "modulo" in log_df.columns:
            log_df = log_df[log_df["modulo"] == "SS"] if st.checkbox("Mostrar só Segurança Social", value=True, key="ss_log_filtro") else log_df
        st.dataframe(log_df, use_container_width=True, height=220)

# --- Template ----------------------------------------------------------------
with tab_template:
    st.subheader("Template do Email da Segurança Social")
    if sou_admin():
        editor_template_bilingue(st.session_state.template_ss, "ss_tpl")
        st.caption(
            "Placeholders disponíveis: {nome} {nif} {email} {mes_nome} {data_limite} {lista_docs}. "
            "{mes_nome} e {data_limite} são calculados a partir do mês escolhido; {lista_docs} lista "
            "automaticamente os documentos anexados a cada cliente (guia, DMR, extras)."
        )
    else:
        st.caption("O template de email é definido pelo administrador.")
        st.text_input("Assunto (PT)", value=st.session_state.template_ss.get("assunto", ""), disabled=True)
        st.text_area("Corpo (PT)", value=st.session_state.template_ss.get("corpo", ""), height=260, disabled=True)

# Persistir template caso o admin o tenha editado (RLS bloqueia gestores).
guardar_config_db(
    st.session_state.params, st.session_state.templates,
    st.session_state.get("template_irs"), st.session_state.get("template_ss"),
)
