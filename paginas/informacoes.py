"""
Página "Informações" — módulo genérico para avisos e documentos avulsos que
não pertencem a nenhum imposto específico (ex: aviso do valor trimestral da
Segurança Social dos Trabalhadores Independentes, ou qualquer outro
comunicado pontual). Não tem "pisco" de obrigação fiscal: funciona sobre
TODOS os clientes do registo central, filtrados/escolhidos diretamente aqui.

Cada envio agrupa-se por "período" — um texto livre definido por quem usa
(ex: "2026-T3 — SS Trab. Independentes") — dentro do qual se pode indicar,
por cliente: um valor (opcional), uma mensagem livre (opcional) e vários
documentos anexos (opcionais, vários por cliente, aditivos, tal como "Outros
documentos" nos outros módulos).
"""

from datetime import datetime

import pandas as pd
import streamlit as st

from common import (
    DEFAULT_TEMPLATE_INFO,
    carregar_info_db,
    editor_template_bilingue,
    enviar_email,
    escolher_conta_email,
    extrair_label_extra,
    extrair_nif_de_filename,
    formatar_valor,
    guardar_config_db,
    guardar_info_valores_db,
    listar_extras_generico,
    listar_periodos_info,
    marcar_envio_db,
    meu_email,
    montar_base_info,
    nomes_ficheiro_unicos,
    obter_documentos_info,
    registar_log,
    render_template_info,
    sanitizar_nome_ficheiro,
    sanitizar_pasta,
    storage_apagar,
    storage_download_pdf,
    storage_upload_pdf,
)

st.title("ℹ️ Informações")
st.caption("SERVE — Contabilidade e Viabilização Empresarial")
st.caption(
    "Módulo genérico para avisos e documentos avulsos que não pertencem a nenhum imposto específico "
    "(ex: valor trimestral da Segurança Social dos Trabalhadores Independentes)."
)

base_completa = montar_base_info()
if base_completa.empty:
    st.info("Ainda não há clientes registados — cria-os na página 'Clientes'.")
    st.stop()

# --- Período -------------------------------------------------------------------
periodos_existentes = listar_periodos_info()
col_p1, col_p2 = st.columns([2, 1])
with col_p2:
    escolha = st.selectbox(
        "Reabrir um período anterior",
        ["(novo período)"] + periodos_existentes,
        key="info_periodo_escolha",
    )
with col_p1:
    periodo = st.text_input(
        "Período / identificador desta informação",
        value="" if escolha == "(novo período)" else escolha,
        placeholder="ex: 2026-T3 — SS Trabalhadores Independentes",
        key=f"info_periodo_texto_{escolha}",
    ).strip()

if not periodo:
    st.info(
        "Define um período/identificador para continuar (ex: '2026-T3 — SS Trabalhadores Independentes'). "
        "Serve só para agrupar este envio — usa um texto diferente para cada assunto/campanha."
    )
    st.stop()

# O período é texto livre (pode ter espaços, acentos, "|", etc.) — guarda-se
# tal e qual na base de dados e no email, mas para o CAMINHO no Storage
# precisa de uma versão "limpa" (o Storage do Supabase rejeita esses
# caracteres com erro "Invalid Key"). periodo_pasta é só para isso.
periodo_pasta = sanitizar_pasta(periodo)

# --- Filtro de destinatários -----------------------------------------------------
with st.expander("🔎 Filtrar clientes elegíveis para este período", expanded=False):
    col_f1, col_f2, col_f3, col_f4 = st.columns(4)
    f_empresa = col_f1.checkbox("Empresa", value=True, key="info_f_empresa")
    f_al = col_f2.checkbox("AL", value=True, key="info_f_al")
    f_ti = col_f3.checkbox("Trab. Independente", value=True, key="info_f_ti")
    f_rf = col_f4.checkbox("Rep. Fiscal", value=True, key="info_f_rf")

filtros_ativos = []
if f_empresa:
    filtros_ativos.append("Tipo_Empresa")
if f_al:
    filtros_ativos.append("Tipo_AL")
if f_ti:
    filtros_ativos.append("Tipo_Trab_Independente")
if f_rf:
    filtros_ativos.append("Tipo_Rep_Fiscal")

base_info = base_completa[base_completa[filtros_ativos].any(axis=1)].copy() if filtros_ativos else base_completa.copy()

if base_info.empty:
    st.warning("Nenhum cliente corresponde aos filtros escolhidos.")
    st.stop()

extras_dict = listar_extras_generico(f"info/{periodo_pasta}/doc")
estado_periodo = carregar_info_db(periodo)

base_info = base_info.reset_index(drop=True)
base_info["Valor"] = base_info["NIF"].map(lambda n: estado_periodo.get(n, {}).get("valor", 0.0))
base_info["Mensagem"] = base_info["NIF"].map(lambda n: estado_periodo.get(n, {}).get("mensagem", ""))
base_info["Email_Enviado"] = base_info["NIF"].map(lambda n: estado_periodo.get(n, {}).get("email_enviado", False))

tab_valores, tab_docs, tab_emails, tab_template = st.tabs(
    ["👥 Clientes e Valores", "📎 Documentos", "✉️ Emails", "✏️ Template de Email"]
)

# --- Clientes e Valores ----------------------------------------------------------
with tab_valores:
    st.subheader(f"Clientes e valores — {periodo}")
    st.caption(
        "Preenche o Valor (opcional, ex: valor trimestral a pagar) e/ou a Mensagem (texto livre) "
        "para cada cliente. Fica guardado só para este período — outro período pode ter valores diferentes."
    )
    tabela = base_info.rename(columns={"Numero_Cliente": "N.º"})[
        ["N.º", "NIF", "Nome", "Valor", "Mensagem", "Email_Enviado"]
    ]
    editado = st.data_editor(
        tabela,
        use_container_width=True,
        hide_index=True,
        height=420,
        disabled=["N.º", "NIF", "Nome", "Email_Enviado"],
        column_config={
            "Valor": st.column_config.NumberColumn("Valor (€)", format="%.2f", step=0.01, min_value=0.0),
            "Mensagem": st.column_config.TextColumn("Mensagem (opcional)"),
            "Email_Enviado": st.column_config.CheckboxColumn("Email Enviado"),
        },
        key=f"info_editor_{periodo}",
    )
    if st.button("💾 Guardar valores e mensagens", key="info_guardar_valores"):
        guardar_info_valores_db(editado, periodo)
        for _, r in editado.iterrows():
            if bool(r["Email_Enviado"]) != estado_periodo.get(r["NIF"], {}).get("email_enviado", False):
                marcar_envio_db("info_dados", r["NIF"], periodo, bool(r["Email_Enviado"]))
        st.success("Guardado.")
        st.rerun()

# --- Documentos --------------------------------------------------------------------
with tab_docs:
    st.subheader(f"Documentos — {periodo}")
    st.caption(
        "Anexos avulsos, por cliente. Cada nome de ficheiro é um documento à parte (não substitui os "
        "outros) — tal como 'Outros documentos' nos outros módulos. Podes carregar em massa (o NIF no "
        "nome do ficheiro associa sozinho) ou cliente a cliente."
    )

    st.markdown("**Carregamento em massa**")
    up_massa = st.file_uploader(
        "Carregar PDFs (nome a começar pelo NIF de 9 dígitos, ex: '267894449_Aviso.pdf')",
        type=["pdf"], accept_multiple_files=True, key=f"info_up_massa_{periodo}",
    )
    if up_massa:
        ids_upload = tuple(sorted(f"{f.name}_{f.size}" for f in up_massa))
        if st.session_state.get("_info_massa_proc") != (periodo, ids_upload):
            st.session_state["_info_massa_proc"] = (periodo, ids_upload)
            ok, sem_nif, detalhes, falhas_upload = 0, [], [], []
            por_nif = {}
            for f in up_massa:
                nif_d = extrair_nif_de_filename(f.name)
                if not nif_d:
                    sem_nif.append(f.name)
                    continue
                por_nif.setdefault(nif_d, []).append(f)
            for nif_d, ficheiros in por_nif.items():
                labels = [extrair_label_extra(f.name, nif_d) for f in ficheiros]
                nomes_novos = nomes_ficheiro_unicos(labels)
                for f, nome_novo in zip(ficheiros, nomes_novos):
                    caminho_novo = f"info/{periodo_pasta}/doc/{nif_d}__{nome_novo}"
                    try:
                        storage_upload_pdf(caminho_novo, f.getvalue())
                        detalhes.append(f"✅ {nif_d} → {caminho_novo}")
                        ok += 1
                    except Exception as e:
                        falhas_upload.append(f"❌ {nif_d} → {caminho_novo}: {e}")
            msg = f"{ok} ficheiro(s) associados e guardados no arquivo."
            if sem_nif:
                msg += f" Sem NIF no nome (usa o carregamento por cliente, abaixo): {', '.join(sem_nif)}"
            st.session_state["_info_ultimo_upload_massa"] = {"msg": msg, "detalhes": detalhes, "falhas": falhas_upload}
            st.rerun()

    ultimo_upload = st.session_state.get("_info_ultimo_upload_massa")
    if ultimo_upload:
        st.success(ultimo_upload["msg"])
        if ultimo_upload["detalhes"]:
            with st.expander(f"Ver os {len(ultimo_upload['detalhes'])} caminho(s) exatos onde foi guardado"):
                st.text("\n".join(ultimo_upload["detalhes"]))
        if ultimo_upload["falhas"]:
            st.error("Estes ficheiros FALHARAM ao guardar:\n" + "\n".join(ultimo_upload["falhas"]))

    st.divider()
    st.markdown("**Carregamento por cliente**")
    nif_doc = st.selectbox(
        "Cliente",
        base_info["NIF"].tolist(),
        format_func=lambda n: f"{n} — {base_info.loc[base_info['NIF']==n,'Nome'].values[0]}",
        key="info_cliente_doc",
    )
    up = st.file_uploader(
        "Documentos (PDF, vários)", type=["pdf"], accept_multiple_files=True,
        key=f"info_up_{periodo}_{nif_doc}",
    )
    if up:
        ids_up = tuple(sorted(f"{f.name}_{f.size}" for f in up))
        chave_proc = f"_info_proc_{periodo}_{nif_doc}"
        if st.session_state.get(chave_proc) != ids_up:
            st.session_state[chave_proc] = ids_up
            nomes_novos = nomes_ficheiro_unicos([sanitizar_nome_ficheiro(f.name) for f in up])
            for f, nome_novo in zip(up, nomes_novos):
                storage_upload_pdf(f"info/{periodo_pasta}/doc/{nif_doc}__{nome_novo}", f.getvalue())
            st.success(f"{len(up)} documento(s) guardado(s).")
            st.rerun()
    existentes = extras_dict.get(nif_doc, [])
    if existentes:
        for nome in existentes:
            c_nome, c_apagar = st.columns([4, 1])
            c_nome.caption(f"📄 {nome}")
            if c_apagar.button("🗑️", key=f"info_apagar_{periodo}_{nif_doc}_{nome}",
                                help="Apagar este documento"):
                storage_apagar(f"info/{periodo_pasta}/doc/{nif_doc}__{nome}")
                st.rerun()
    else:
        st.caption("Sem documentos para este cliente neste período.")

    st.divider()
    st.markdown("**Estado do período por cliente**")
    rows = []
    for _, r in base_info.iterrows():
        rows.append({
            "N.º": r.get("Numero_Cliente", ""), "NIF": r["NIF"], "Nome": r["Nome"],
            "Valor": f"{formatar_valor(r['Valor'])} €" if r["Valor"] else "—",
            "Documentos": len(extras_dict.get(r["NIF"], [])),
            "Email Enviado": bool(r["Email_Enviado"]),
        })
    st.dataframe(pd.DataFrame(rows), use_container_width=True, height=360, hide_index=True)

# --- Emails --------------------------------------------------------------------
with tab_emails:
    st.subheader(f"Enviar Emails — {periodo}")

    elegiveis = base_info[base_info["Email"].str.strip() != ""].copy()
    sem_email = len(base_info) - len(elegiveis)
    if sem_email:
        st.caption(f"⚠️ {sem_email} cliente(s) sem email preenchido no registo central — não aparecem abaixo.")

    tpl = st.session_state.template_info

    com_conteudo = [
        n for n in elegiveis["NIF"]
        if n in extras_dict or estado_periodo.get(n, {}).get("valor") or estado_periodo.get(n, {}).get("mensagem")
    ]
    nao_enviados = [n for n in elegiveis["NIF"] if not estado_periodo.get(n, {}).get("email_enviado", False)]

    preview_nif = st.selectbox(
        "Pré-visualizar cliente:",
        elegiveis["NIF"].tolist(),
        format_func=lambda n: f"{n} — {elegiveis.loc[elegiveis['NIF']==n,'Nome'].values[0]}",
        key=f"info_preview_{periodo}",
    )
    if preview_nif:
        row = elegiveis[elegiveis["NIF"] == preview_nif].iloc[0]
        docs = obter_documentos_info(periodo_pasta, preview_nif, extras_dict)
        assunto, corpo = render_template_info(tpl, row, periodo, row["Valor"], row["Mensagem"], docs)
        st.text_input("Assunto (preview)", value=assunto, disabled=True)
        if row["Gestor_Email"]:
            st.caption(f"📋 CC: {row['Gestor_Nome'] or ''} <{row['Gestor_Email']}>  ·  Língua: {row['Lingua']}")
        else:
            st.caption(f"📋 CC: — (sem gestor definido)  ·  Língua: {row['Lingua']}")
        st.text_area("Corpo (preview)", value=corpo, height=260, disabled=True)
        st.caption("📎 Anexos: " + (", ".join(d["tipo"] for d in docs) if docs else "nenhum"))
        if not row["Valor"] and not row["Mensagem"] and not docs:
            st.caption("⚠️ Este cliente ainda não tem valor, mensagem nem documentos definidos para este período.")

    st.divider()
    smtp_cfg = escolher_conta_email("info")

    st.markdown(f"📎 **{len(com_conteudo)} de {len(elegiveis)}** cliente(s) já têm conteúdo definido para este período.")

    multiselect_key = f"info_selecionados_{periodo}"
    col_b1, col_b2, col_b3 = st.columns(3)
    with col_b1:
        if st.button("📎 Selecionar quem tem conteúdo e falta enviar"):
            st.session_state[multiselect_key] = [n for n in com_conteudo if n in nao_enviados]
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
        st.session_state[multiselect_key] = [n for n in com_conteudo if n in nao_enviados]

    selecionados = st.multiselect(
        "Clientes selecionados para envio (podes ajustar — para enviar só um, deixa só esse)",
        elegiveis["NIF"].tolist(),
        format_func=lambda n: f"{n} — {elegiveis.loc[elegiveis['NIF']==n,'Nome'].values[0]}"
        + ("" if n in com_conteudo else "  ⚠️ sem valor/mensagem/documentos")
        + ("  ✅ já enviado" if estado_periodo.get(n, {}).get("email_enviado", False) else ""),
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
                docs = obter_documentos_info(periodo_pasta, nif, extras_dict)
                anexos = []
                for d in docs:
                    conteudo = storage_download_pdf(d["caminho"])
                    if conteudo:
                        anexos.append((d["anexo"], conteudo))
                assunto, corpo = render_template_info(tpl, row, periodo, row["Valor"], row["Mensagem"], docs)
                try:
                    cc_gestor = [row["Gestor_Email"]] if row["Gestor_Email"] else []
                    enviar_email(smtp_cfg, row["Email"], assunto, corpo, anexos, cc=cc_gestor,
                                 bcc=[smtp_cfg["remetente"]], assinatura_html=assinatura)
                    marcar_envio_db("info_dados", nif, periodo, True)
                    registar_log({
                        "data": datetime.now().strftime("%Y-%m-%d %H:%M"), "nif": nif,
                        "nome": row["Nome"], "pagamento": 0, "estado": f"Enviado ({periodo})",
                        "modulo": "Informações", "enviado_por": meu_email(),
                    })
                    sucessos += 1
                except Exception as e:
                    registar_log({
                        "data": datetime.now().strftime("%Y-%m-%d %H:%M"), "nif": nif,
                        "nome": row["Nome"], "pagamento": 0, "estado": f"Erro ({periodo}): {e}",
                        "modulo": "Informações", "enviado_por": meu_email(),
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
            log_df = log_df[log_df["modulo"] == "Informações"] if st.checkbox(
                "Mostrar só Informações", value=True, key="info_log_filtro") else log_df
        st.dataframe(log_df, use_container_width=True, height=220)

# --- Template ----------------------------------------------------------------
with tab_template:
    st.subheader("Template do Email de Informações")
    if st.button("🔄 Repor template padrão", key="info_tpl_reset"):
        st.session_state.template_info = DEFAULT_TEMPLATE_INFO.copy()
        st.rerun()
    editor_template_bilingue(st.session_state.template_info, "info_tpl")
    st.caption(
        "Placeholders disponíveis: {nome} {nif} {email} {periodo} {valor} {mensagem} {lista_docs}. "
        "{valor} só aparece se preencheres um valor para o cliente (fica vazio caso contrário); "
        "{mensagem} é o texto livre escrito para o cliente na aba 'Clientes e Valores'; {lista_docs} "
        "lista automaticamente os documentos anexados, se houver. Alterações aqui ficam guardadas "
        "para toda a equipa."
    )

guardar_config_db(
    st.session_state.params, st.session_state.templates,
    st.session_state.get("template_irs"), st.session_state.get("template_ss"),
    {"iva": st.session_state.get("template_iva"), "imi": st.session_state.get("template_imi"),
     "info": st.session_state.get("template_info")},
)
