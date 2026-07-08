"""
Página da Segurança Social — envio mensal de DMR, DRI e outros documentos
avulsos (ex: IUC, retenções). Tudo organizado por mês de referência:
escolhe-se o mês, carregam-se os documentos (ficam no arquivo persistente,
identificados pelo NIF no nome do ficheiro) e enviam-se os emails. Quando
possível, o valor de cada documento é lido automaticamente do PDF e listado
no corpo do email. Pagamento até dia 25 do mês seguinte ao mês de referência.
"""

from datetime import datetime

import pandas as pd
import streamlit as st

from common import (
    carregar_anexos_e_valores_ss,
    carregar_ss_mes_db,
    data_limite_ss,
    editor_template_bilingue,
    gerar_excel_estado_mensal,
    enviar_email,
    escolher_conta_email,
    extrair_label_extra,
    extrair_nif_de_filename,
    guardar_config_db,
    lista_meses_ss,
    listar_extras_generico,
    marcar_ss_enviado_db,
    meu_email,
    montar_base_ss,
    nome_mes,
    nomes_ficheiro_unicos,
    obter_documentos_ss,
    registar_log,
    render_template_ss,
    sanitizar_nome_ficheiro,
    storage_apagar,
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

# Estado deste mês — cada categoria é {nif: [nome1.pdf, nome2.pdf, ...]},
# porque agora um cliente pode ter mais do que um ficheiro na mesma categoria
# (ex: 2 DMRs). O "Sim/Não" nas tabelas usa "nif in dicionario", que continua
# a funcionar normalmente com dicts (verifica só as chaves).
dmrs_dict = listar_extras_generico(f"ss/{mes}/dmr")
dris_dict = listar_extras_generico(f"ss/{mes}/dri")
retencoes_dict = listar_extras_generico(f"ss/{mes}/retencoes")
iuc_dict = listar_extras_generico(f"ss/{mes}/iuc")
extras_dict = listar_extras_generico(f"ss/{mes}/extra")
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
    st.caption("Se um cliente tiver 2 documentos da mesma categoria, distingue-os no nome do ficheiro (ex: '267894449_Entidade A.pdf').")
    col_tipo, col_up = st.columns([1, 3])
    with col_tipo:
        tipo_doc = st.radio(
            "Tipo de documento",
            ["DMR", "DRI", "Retenções", "IUC", "Outros documentos"],
            key="ss_tipo_doc",
        )
    with col_up:
        up_massa = st.file_uploader(
            "Carregar PDFs (nome a começar pelo NIF de 9 dígitos — o resto do nome pode ser "
            "o que quiseres, ex: '267894449_IUC.pdf', para depois identificares o documento)",
            type=["pdf"], accept_multiple_files=True, key="ss_up_massa",
        )
    PASTAS_TIPO_DOC = {"DMR": "dmr", "DRI": "dri", "Retenções": "retencoes", "IUC": "iuc", "Outros documentos": "extra"}
    if up_massa:
        ids_upload = tuple(sorted(f"{f.name}_{f.size}" for f in up_massa))
        if st.session_state.get("_ss_massa_proc") != (mes, tipo_doc, ids_upload):
            st.session_state["_ss_massa_proc"] = (mes, tipo_doc, ids_upload)
            pasta = PASTAS_TIPO_DOC[tipo_doc]
            ok, sem_nif = 0, []
            itens = []  # (ficheiro, nif) por ordem, para depois desambiguar nomes repetidos
            labels_por_nif = {}
            for f in up_massa:
                nif_d = extrair_nif_de_filename(f.name)
                if not nif_d:
                    sem_nif.append(f.name)
                    continue
                label = extrair_label_extra(f.name, nif_d)
                if tipo_doc != "Outros documentos" and label == "Documento.pdf":
                    label = f"{tipo_doc}.pdf"  # só veio o NIF no nome -> usa o próprio tipo como etiqueta
                itens.append((f, nif_d))
                labels_por_nif.setdefault(nif_d, []).append(label)
            # Se dois ficheiros do MESMO cliente derem o mesmo nome (ex: duas DMRs
            # ambas só com o NIF no nome), desambigua-se para nenhum se perder.
            labels_unicos_por_nif = {nif: iter(nomes_ficheiro_unicos(labels)) for nif, labels in labels_por_nif.items()}
            for f, nif_d in itens:
                label = next(labels_unicos_por_nif[nif_d])
                storage_upload_pdf(f"ss/{mes}/{pasta}/{nif_d}__{label}", f.getvalue())
                ok += 1
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
    def _gestor_documentos(col, rotulo: str, pasta: str, dicionario: dict, opcional: bool = False):
        """Upload (aceita vários ficheiros) + lista com botão de apagar, para
        uma categoria de documento (DMR, DRI, Retenções, IUC ou Outros) de um
        cliente/mês. Suporta mais do que um ficheiro por cliente (ex: 2 DMRs)
        e substituir um documento (apaga o antigo e carrega de novo)."""
        with col:
            rotulo_campo = f"{rotulo} (PDF{', opcional' if opcional else ''})"
            up = st.file_uploader(rotulo_campo, type=["pdf"], accept_multiple_files=True,
                                   key=f"ss_up_{pasta}_{mes}_{nif_doc}")
            if up:
                ids_up = tuple(sorted(f"{f.name}_{f.size}" for f in up))
                if st.session_state.get(f"_ss_{pasta}_proc_{mes}_{nif_doc}") != ids_up:
                    st.session_state[f"_ss_{pasta}_proc_{mes}_{nif_doc}"] = ids_up
                    nomes_seguros = nomes_ficheiro_unicos([sanitizar_nome_ficheiro(f.name) for f in up])
                    for f, nome_seguro in zip(up, nomes_seguros):
                        storage_upload_pdf(f"ss/{mes}/{pasta}/{nif_doc}__{nome_seguro}", f.getvalue())
                    st.success(f"{len(up)} ficheiro(s) de {rotulo} guardado(s).")
                    st.rerun()
            existentes = dicionario.get(nif_doc, [])
            if existentes:
                for nome in existentes:
                    c_nome, c_apagar = st.columns([4, 1])
                    c_nome.caption(f"📄 {nome}")
                    if c_apagar.button("🗑️", key=f"apagar_{pasta}_{mes}_{nif_doc}_{nome}",
                                        help="Apagar (depois podes carregar outro ficheiro em substituição)"):
                        storage_apagar(f"ss/{mes}/{pasta}/{nif_doc}__{nome}")
                        st.rerun()
            else:
                st.caption(f"Sem {rotulo}")

    c1, c2, c3, c4, c5 = st.columns(5)
    _gestor_documentos(c1, "DMR", "dmr", dmrs_dict)
    _gestor_documentos(c2, "DRI", "dri", dris_dict)
    _gestor_documentos(c3, "Retenções", "retencoes", retencoes_dict)
    _gestor_documentos(c4, "IUC", "iuc", iuc_dict)
    _gestor_documentos(c5, "Outros documentos", "extra", extras_dict, opcional=True)

    st.divider()
    st.markdown("**Estado do mês por cliente**")
    rows = []
    for _, r in base_ss.iterrows():
        rows.append({
            "N.º": r.get("Numero_Cliente", ""), "NIF": r["NIF"], "Nome": r["Nome"],
            "DMR": "✅" if r["NIF"] in dmrs_dict else "❌",
            "DRI": "✅" if r["NIF"] in dris_dict else "❌",
            "Retenções": "✅" if r["NIF"] in retencoes_dict else "❌",
            "IUC": "✅" if r["NIF"] in iuc_dict else "❌",
            "Extras": len(extras_dict.get(r["NIF"], [])),
            "Email Enviado": bool(r["Email_Enviado"]),
        })
    estado_df = pd.DataFrame(rows)
    editado = st.data_editor(
        estado_df,
        use_container_width=True,
        hide_index=True,
        height=360,
        disabled=["N.º", "NIF", "Nome", "DMR", "DRI", "Retenções", "IUC", "Extras"],
        column_config={"Email Enviado": st.column_config.CheckboxColumn("Email Enviado")},
        key=f"ss_estado_{mes}",
    )
    if st.button("💾 Guardar piscos 'Email Enviado'"):
        for _, r in editado.iterrows():
            if bool(r["Email Enviado"]) != enviados.get(r["NIF"], False):
                marcar_ss_enviado_db(r["NIF"], mes, bool(r["Email Enviado"]))
        st.success("Estado guardado.")
        st.rerun()

    excel_ss = gerar_excel_estado_mensal(
        f"Controlo Segurança Social — {nome_mes(mes)}", base_ss, dmrs_dict, dris_dict, extras_dict, enviados,
        rotulo_guia="DMR", rotulo_decl="DRI",
        extra_categorias=[("Retenções", retencoes_dict), ("IUC", iuc_dict)],
    )
    st.download_button(
        "⬇️ Descarregar Excel de Controlo (SS)", excel_ss,
        file_name=f"Controlo_SS_{mes}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# --- Emails ------------------------------------------------------------------
with tab_emails:
    st.subheader(f"Enviar Emails — {nome_mes(mes)}")

    elegiveis = base_ss[base_ss["Email"].str.strip() != ""].copy()
    sem_email = len(base_ss) - len(elegiveis)
    if sem_email:
        st.caption(f"⚠️ {sem_email} cliente(s) sem email preenchido no registo central — não aparecem abaixo.")

    tpl = st.session_state.template_ss

    com_docs = [
        n for n in elegiveis["NIF"]
        if n in dmrs_dict or n in dris_dict or n in retencoes_dict or n in iuc_dict or n in extras_dict
    ]
    nao_enviados = [n for n in elegiveis["NIF"] if not enviados.get(n, False)]

    preview_nif = st.selectbox(
        "Pré-visualizar cliente:",
        elegiveis["NIF"].tolist(),
        format_func=lambda n: f"{n} — {elegiveis.loc[elegiveis['NIF']==n,'Nome'].values[0]}",
        key=f"ss_preview_{mes}",
    )
    if preview_nif:
        row = elegiveis[elegiveis["NIF"] == preview_nif].iloc[0]
        docs = obter_documentos_ss(mes, preview_nif, dmrs_dict, dris_dict, retencoes_dict, iuc_dict, extras_dict)
        _, valores_prev = carregar_anexos_e_valores_ss(docs) if docs else ([], [])
        assunto, corpo = render_template_ss(tpl, row, mes, docs, valores_prev)
        st.text_input("Assunto (preview)", value=assunto, disabled=True)
        if row["Gestor_Email"]:
            st.caption(f"📋 CC: {row['Gestor_Nome'] or ''} <{row['Gestor_Email']}>  ·  Língua: {row['Lingua']}")
        else:
            st.caption(f"📋 CC: — (sem gestor definido)  ·  Língua: {row['Lingua']}")
        st.text_area("Corpo (preview)", value=corpo, height=260, disabled=True)
        st.caption("📎 Anexos: " + (", ".join(d["tipo"] for d in docs) if docs else "nenhum documento carregado ainda"))
        if not valores_prev and docs:
            st.caption("ℹ️ Não foi possível ler automaticamente o valor de nenhum destes documentos — confirma manualmente se precisares de indicar montantes.")

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
                docs = obter_documentos_ss(mes, nif, dmrs_dict, dris_dict, retencoes_dict, iuc_dict, extras_dict)
                anexos, valores = carregar_anexos_e_valores_ss(docs)
                assunto, corpo = render_template_ss(tpl, row, mes, docs, valores)
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
    editor_template_bilingue(st.session_state.template_ss, "ss_tpl")
    st.caption(
        "Placeholders disponíveis: {nome} {nif} {email} {mes_nome} {data_limite} {lista_docs} {lista_valores}. "
        "{mes_nome} e {data_limite} são calculados a partir do mês escolhido; {lista_docs} lista "
        "automaticamente os documentos anexados a cada cliente (DMR, DRI, extras); {lista_valores} é o bloco "
        "com os montantes lidos automaticamente de cada documento (ex: 'DMR - 45,00 €'), quando for possível "
        "lê-los — fica vazio se não conseguir ler nenhum valor. Alterações aqui ficam guardadas para toda a equipa."
    )

# Persistir template (guardado para toda a equipa, qualquer utilizador pode editar).
guardar_config_db(
    st.session_state.params, st.session_state.templates,
    st.session_state.get("template_irs"), st.session_state.get("template_ss"),
)
