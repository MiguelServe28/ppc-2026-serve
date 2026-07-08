"""
Página do PPC — Pagamentos por Conta. Tudo o que é específico deste imposto
vive aqui: dados de PPC por cliente, guias (guardadas de forma persistente no
Supabase Storage), emails e exportação. Os parâmetros de cálculo editam-se na
página 'Configurações' (admin).
"""

from datetime import date, datetime

import pandas as pd
import streamlit as st

from common import (
    PPC_COLS,
    calcular_ppc,
    editor_template_bilingue,
    enviar_email,
    escolher_conta_email,
    extrair_nif_de_filename,
    formatar_valor,
    gerar_excel_ppc,
    guardar_config_db,
    meu_email,
    montar_base_ppc,
    persistir_ppc,
    registar_log,
    render_template,
    sou_admin,
    storage_download_pdf,
    storage_listar,
    storage_upload_pdf,
)

p = st.session_state.params
ano_dados = p.get("ano_dados", 2025)
ano_pag = p.get("ano_pagamentos", 2026)

st.title(f"💶 PPC — Pagamentos por Conta {ano_pag}")
st.caption("SERVE — Contabilidade e Viabilização Empresarial")

with st.sidebar:
    st.header("Parâmetros de Cálculo (PPC)")
    st.write(f"Limiar Volume de Negócios: **{p['limiar_volume']:,.2f} €**")
    st.write(f"Taxa ≤ limiar / > limiar: **{p['taxa_baixa']:.0%} / {p['taxa_alta']:.0%}**")
    st.write(f"Limite de dispensa: **{p['limite_dispensa']:,.2f} €**")
    st.divider()
    st.write(f"Data limite 1.º Pagamento: **{p['data1'].strftime('%d/%m/%Y')}**")
    st.write(f"Data limite 2.º Pagamento: **{p['data2'].strftime('%d/%m/%Y')}**")
    st.write(f"Data limite 3.º Pagamento: **{p['data3'].strftime('%d/%m/%Y')}**")
    st.divider()
    if sou_admin():
        st.caption("✏️ Edita estes parâmetros na página 'Configurações'.")
    else:
        st.caption("Estes parâmetros são definidos pelo administrador e aplicam-se a todos os clientes.")
    st.caption("Fórmula: Total PPC = (Coleta IRC − Retenções) × Taxa, repartido em 3 prestações iguais, cada uma arredondada por excesso para euro (art. 105.º CIRC). Dispensa se (Coleta − Retenções) < limite definido.")

df_calc = calcular_ppc(montar_base_ppc(), st.session_state.params)


def nifs_com_guia_storage(n_pag: int) -> set:
    """NIFs com guia carregada no Storage para este pagamento."""
    return {nome[:-4] for nome in storage_listar(f"ppc/{n_pag}") if nome.lower().endswith(".pdf")}


tab_visao, tab_dados, tab_guias, tab_emails, tab_export = st.tabs(
    ["📊 Visão PPC", "🧮 Dados e Cálculo", "📎 Guias", "✉️ Emails", "⬇️ Exportar"]
)

# --- Visão PPC ---------------------------------------------------------
with tab_visao:
    if df_calc.empty:
        st.info("Ainda não há clientes com 'Aplica PPC' ligado. Ativa esse interruptor na página 'Clientes'.")
    else:
        elegiveis_dash = df_calc[~df_calc["Dispensado"]]
        c1, c2, c3 = st.columns(3)
        c1.metric("Clientes PPC", len(df_calc))
        c2.metric("Dispensados", int(df_calc["Dispensado"].sum()))
        c3.metric("Elegíveis para Pagamento", len(elegiveis_dash))

        st.divider()
        st.markdown("### Estado por Pagamento")
        cols = st.columns(3)
        resumo = []
        for i, n in enumerate([1, 2, 3]):
            total = len(elegiveis_dash)
            guias_storage = nifs_com_guia_storage(n)
            com_guia = int(elegiveis_dash["NIF"].isin(guias_storage).sum()) if total else 0
            enviados = int(elegiveis_dash[f"Email{n}_Enviado"].sum()) if total else 0
            pendentes = total - enviados
            resumo.append({"Pagamento": f"{n}.º", "Total Elegíveis": total, "Guia no Arquivo": com_guia,
                            "Emails Enviados": enviados, "Pendentes": pendentes})
            with cols[i]:
                st.markdown(f"**{n}.º Pagamento**")
                st.metric("Enviados", enviados, delta=f"-{pendentes} pendentes" if pendentes else "Completo", delta_color="inverse" if pendentes else "off")
                st.progress(enviados / total if total else 0)

        st.divider()
        st.markdown("### Tabela Resumo")
        st.dataframe(pd.DataFrame(resumo), use_container_width=True, hide_index=True)

        st.divider()
        st.markdown("### Clientes com Pagamentos Pendentes")
        pag_filtro = st.selectbox("Ver pendentes de:", [1, 2, 3], format_func=lambda x: f"{x}.º Pagamento", key="dash_filtro")
        pendentes_df = elegiveis_dash[~elegiveis_dash[f"Email{pag_filtro}_Enviado"]][["NIF", "Nome", "Email", f"Pag{pag_filtro}"]]
        if pendentes_df.empty:
            st.success(f"Todos os clientes elegíveis já receberam o {pag_filtro}.º pagamento. 🎉")
        else:
            st.dataframe(pendentes_df, use_container_width=True, height=300, hide_index=True)

        st.caption("💾 Tudo fica guardado de forma persistente: os dados no Supabase e as guias em PDF no arquivo (Storage) — já não se perdem ao fechar o browser.")

# --- Dados e Cálculo -----------------------------------------------------
with tab_dados:
    st.subheader("Dados de PPC")
    base_ppc = montar_base_ppc()
    if base_ppc.empty:
        st.info("Ainda não há clientes com 'Aplica PPC' ligado. Ativa esse interruptor na página 'Clientes'.")
    else:
        st.caption(f"Preenche aqui o Volume de Negócios, Coleta IRC e Retenções de {ano_dados} dos clientes elegíveis para PPC. Os piscos de Guia Emitida / Email Enviado são marcados automaticamente pelas abas 'Guias' e 'Emails', mas também os podes marcar/desmarcar aqui à mão.")
        editor_ppc = st.data_editor(
            base_ppc[["NIF", "Nome"] + [c for c in PPC_COLS if c != "NIF"]],
            use_container_width=True,
            hide_index=True,
            disabled=["NIF", "Nome"],
            column_config={
                "Volume": st.column_config.NumberColumn(f"Volume Negócios {ano_dados} (campo 411)", format="%.2f"),
                "Coleta": st.column_config.NumberColumn(f"Coleta IRC {ano_dados} (campo 351)", format="%.2f"),
                "Retencoes": st.column_config.NumberColumn(f"Retenções {ano_dados} (campo 359)", format="%.2f"),
                "Guia1_Emitida": st.column_config.CheckboxColumn("Guia 1 Emitida"),
                "Guia2_Emitida": st.column_config.CheckboxColumn("Guia 2 Emitida"),
                "Guia3_Emitida": st.column_config.CheckboxColumn("Guia 3 Emitida"),
                "Email1_Enviado": st.column_config.CheckboxColumn("Email 1 Enviado"),
                "Email2_Enviado": st.column_config.CheckboxColumn("Email 2 Enviado"),
                "Email3_Enviado": st.column_config.CheckboxColumn("Email 3 Enviado"),
            },
            key="editor_ppc_dados",
        )
        if st.button("💾 Guardar dados de PPC"):
            persistir_ppc(editor_ppc[PPC_COLS])
            st.success("Dados de PPC guardados.")
            st.rerun()

    st.divider()
    st.subheader("Resultado do Cálculo")
    if df_calc.empty:
        st.info("Ainda não há clientes elegíveis para PPC.")
    else:
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Nº Clientes", len(df_calc))
        c2.metric("Nº Dispensados", int(df_calc["Dispensado"].sum()))
        c3.metric("Total PPC a cobrar", formatar_valor(df_calc.loc[~df_calc["Dispensado"], "Total_PPC"].sum()) + " €")
        c4.metric("Valor médio / cliente", formatar_valor(df_calc.loc[~df_calc["Dispensado"], "Total_PPC"].mean() if (~df_calc["Dispensado"]).any() else 0) + " €")

        def highlight_dispensado(row):
            return ["background-color: #E2EFDA" if row["Dispensado"] else "" for _ in row]

        show_cols = ["NIF", "Nome", "Volume", "Coleta", "Retencoes",
                     "Base_Calculo", "Taxa", "Total_PPC", "Dispensado", "Pag1", "Pag2", "Pag3"]
        st.dataframe(
            df_calc[show_cols].style.apply(highlight_dispensado, axis=1).format(
                {"Volume": "{:,.2f}", "Coleta": "{:,.2f}", "Retencoes": "{:,.2f}",
                 "Base_Calculo": "{:,.2f}", "Taxa": "{:.0%}", "Total_PPC": "{:,.2f}",
                 "Pag1": "{:,.2f}", "Pag2": "{:,.2f}", "Pag3": "{:,.2f}"}
            ),
            use_container_width=True,
            height=420,
        )
        st.caption("Validar o cálculo contra o simulador da OCC em alguns casos reais antes de confiar 100% na fórmula.")

# --- Guias -----------------------------------------------------------------
with tab_guias:
    st.subheader("Associar Guias (PDF) aos Clientes")
    st.caption(
        "Carrega os PDFs das guias — se o nome do ficheiro contiver o NIF (9 dígitos), a associação é automática "
        "e a guia fica guardada no arquivo persistente (não se perde ao fechar o browser). "
        "Ficheiros sem NIF no nome ficam à espera de associação manual abaixo."
    )

    n_pag = st.selectbox("A que pagamento correspondem estas guias?", [1, 2, 3], format_func=lambda x: f"{x}.º Pagamento")
    up_guias = st.file_uploader("Carregar guias PDF", type=["pdf"], accept_multiple_files=True, key="up_guias")

    if up_guias:
        ids_upload = tuple(sorted(f"{f.name}_{f.size}" for f in up_guias))
        if st.session_state.get("_guias_ppc_processadas") != (n_pag, ids_upload):
            st.session_state["_guias_ppc_processadas"] = (n_pag, ids_upload)
            auto, manuais = 0, 0
            for f in up_guias:
                nif_detetado = extrair_nif_de_filename(f.name)
                conteudo = f.getvalue()
                if nif_detetado:
                    storage_upload_pdf(f"ppc/{n_pag}/{nif_detetado}.pdf", conteudo)
                    auto += 1
                else:
                    st.session_state.guias_por_associar[(n_pag, f.name)] = conteudo
                    manuais += 1
            msg = f"{auto} guia(s) associadas automaticamente e guardadas no arquivo."
            if manuais:
                msg += f" {manuais} ficheiro(s) sem NIF no nome — associa-os manualmente abaixo."
            st.success(msg)

    guias_storage = nifs_com_guia_storage(n_pag)

    pendentes_associar = [(k, v) for k, v in st.session_state.guias_por_associar.items() if k[0] == n_pag]
    if pendentes_associar and not df_calc.empty:
        st.markdown("**Associação manual**")
        st.caption("Estes ficheiros não tinham NIF no nome. Escolhe o cliente a quem cada um pertence — ao associar, a guia fica guardada no arquivo.")
        opcoes_ficheiro = {k[1]: k for k, _ in pendentes_associar}
        col_a, col_b, col_c = st.columns([2, 2, 1])
        with col_a:
            ficheiro_escolhido = st.selectbox("Ficheiro", list(opcoes_ficheiro.keys()), key="manual_ficheiro")
        with col_b:
            cliente_escolhido = st.selectbox(
                "Cliente correto",
                df_calc["NIF"].tolist(),
                format_func=lambda n: f"{n} — {df_calc.loc[df_calc['NIF']==n,'Nome'].values[0]}",
                key="manual_cliente",
            )
        with col_c:
            st.write("")
            st.write("")
            if st.button("Associar", key="btn_associar_manual"):
                chave = opcoes_ficheiro[ficheiro_escolhido]
                conteudo = st.session_state.guias_por_associar.pop(chave)
                storage_upload_pdf(f"ppc/{n_pag}/{cliente_escolhido}.pdf", conteudo)
                st.success(f"'{ficheiro_escolhido}' associado a {cliente_escolhido} e guardado no arquivo.")
                st.rerun()

    if not df_calc.empty:
        st.markdown("**Estado das guias por cliente:**")
        rows = []
        for _, r in df_calc.iterrows():
            tem_guia = r["NIF"] in guias_storage
            rows.append({"NIF": r["NIF"], "Nome": r["Nome"], "Guia no arquivo": "✅" if tem_guia else "❌",
                         f"Guia{n_pag}_Emitida (registo manual)": r[f"Guia{n_pag}_Emitida"]})
        st.dataframe(pd.DataFrame(rows), use_container_width=True, height=300)

        if st.button(f"Marcar Guia {n_pag} como Emitida para todos os clientes com guia no arquivo"):
            df_full = montar_base_ppc()
            for idx, r in df_full.iterrows():
                if r["NIF"] in guias_storage:
                    df_full.at[idx, f"Guia{n_pag}_Emitida"] = True
            persistir_ppc(df_full[PPC_COLS])
            st.success("Estado atualizado e guardado.")
            st.rerun()

# --- Emails ------------------------------------------------------------
with tab_emails:
    st.subheader("Gerar e Enviar Emails")

    if df_calc.empty:
        st.info("Ainda não há clientes elegíveis para PPC.")
    else:
        n_pag_email = st.selectbox("Qual pagamento?", [1, 2, 3], format_func=lambda x: f"{x}.º Pagamento", key="n_pag_email")
        tpl = st.session_state.templates[n_pag_email]

        with st.expander("✏️ Editar template deste email (PT e EN)"):
            editor_template_bilingue(tpl, f"ppc_tpl_{n_pag_email}", altura=300)
            st.caption("Placeholders disponíveis: {nome} {nif} {email} {pag1} {pag2} {pag3} {total} {data1} {data2} {data3} {ano_dados} {ano_pagamentos}. Cada cliente recebe na língua definida no registo central (coluna 'Língua'). Alterações aqui ficam guardadas para toda a equipa.")

        elegiveis = df_calc[~df_calc["Dispensado"]].copy()
        elegiveis = elegiveis[elegiveis["Email"].str.strip() != ""]

        st.markdown(f"**{len(elegiveis)} clientes elegíveis** (não dispensados, com email preenchido).")

        guias_storage_email = nifs_com_guia_storage(n_pag_email)

        preview_nif = st.selectbox("Pré-visualizar cliente:", elegiveis["NIF"].tolist() if not elegiveis.empty else [])
        if preview_nif:
            row = elegiveis[elegiveis["NIF"] == preview_nif].iloc[0]
            assunto, corpo = render_template(tpl, row, st.session_state.params)
            st.text_input("Assunto (preview)", value=assunto, disabled=True)
            if row["Gestor_Email"]:
                st.caption(f"📋 CC: {row['Gestor_Nome'] or ''} <{row['Gestor_Email']}>")
            else:
                st.caption("📋 CC: — (sem gestor definido para este cliente)")
            st.text_area("Corpo (preview)", value=corpo, height=250, disabled=True)
            tem_guia = row["NIF"] in guias_storage_email
            st.write("📎 Guia no arquivo:", "✅ Sim" if tem_guia else "❌ Não carregada (aba Guias)")

        st.divider()
        smtp_cfg = escolher_conta_email("ppc")

        com_guia = [n for n in elegiveis["NIF"].tolist() if n in guias_storage_email]
        sem_guia = [n for n in elegiveis["NIF"].tolist() if n not in com_guia]
        nao_enviados = [
            n for n in elegiveis["NIF"].tolist()
            if not df_calc.loc[df_calc["NIF"] == n, f"Email{n_pag_email}_Enviado"].iloc[0]
        ]

        st.markdown(f"📎 **{len(com_guia)} de {len(elegiveis)} clientes elegíveis já têm guia no arquivo** para este pagamento.")
        if sem_guia:
            st.caption(f"Sem guia no arquivo (vão ser enviados sem anexo): {len(sem_guia)} cliente(s) — associa-os na aba 'Guias'.")

        multiselect_key = f"selecionados_email_{n_pag_email}"

        col_btn1, col_btn2, col_btn3 = st.columns(3)
        with col_btn1:
            if st.button("📎 Selecionar só quem tem guia no arquivo"):
                st.session_state[multiselect_key] = [n for n in com_guia if n in nao_enviados]
                st.rerun()
        with col_btn2:
            if st.button("☑️ Selecionar todos os elegíveis"):
                st.session_state[multiselect_key] = nao_enviados
                st.rerun()
        with col_btn3:
            if st.button("✖️ Limpar seleção"):
                st.session_state[multiselect_key] = []
                st.rerun()

        if multiselect_key not in st.session_state:
            st.session_state[multiselect_key] = [n for n in com_guia if n in nao_enviados]

        selecionados = st.multiselect(
            "Clientes selecionados para envio (podes ajustar manualmente)",
            elegiveis["NIF"].tolist(),
            format_func=lambda n: f"{n} — {elegiveis.loc[elegiveis['NIF']==n,'Nome'].values[0]}" + ("" if n in com_guia else "  ⚠️ sem guia"),
            key=multiselect_key,
        )

        if st.button("🚀 Enviar Emails Selecionados", type="primary", disabled=not selecionados):
            if not smtp_cfg["utilizador"] or not smtp_cfg["password"]:
                st.error("Escolhe ou cria uma conta de email com utilizador e password preenchidos.")
            else:
                progress = st.progress(0.0)
                status_box = st.empty()
                df_full = montar_base_ppc()
                assinatura = st.session_state.params.get("assinatura_html", "")
                sucessos, falhas = 0, 0
                for i, nif in enumerate(selecionados):
                    row = elegiveis[elegiveis["NIF"] == nif].iloc[0]
                    assunto, corpo = render_template(tpl, row, st.session_state.params)
                    anexos = []
                    guia_bytes = storage_download_pdf(f"ppc/{n_pag_email}/{nif}.pdf") if nif in guias_storage_email else None
                    if guia_bytes:
                        anexos.append((f"Guia_{n_pag_email}Pagamento_{nif}.pdf", guia_bytes))
                    try:
                        cc_gestor = [row["Gestor_Email"]] if row["Gestor_Email"] else []
                        enviar_email(smtp_cfg, row["Email"], assunto, corpo, anexos, cc=cc_gestor,
                                     bcc=[smtp_cfg["remetente"]], assinatura_html=assinatura)
                        idx = df_full.index[df_full["NIF"] == nif][0]
                        df_full.at[idx, f"Email{n_pag_email}_Enviado"] = True
                        registar_log(
                            {"data": datetime.now().strftime("%Y-%m-%d %H:%M"), "nif": nif,
                             "nome": row["Nome"], "pagamento": n_pag_email, "estado": "Enviado",
                             "modulo": "PPC", "enviado_por": meu_email()}
                        )
                        sucessos += 1
                    except Exception as e:
                        registar_log(
                            {"data": datetime.now().strftime("%Y-%m-%d %H:%M"), "nif": nif,
                             "nome": row["Nome"], "pagamento": n_pag_email, "estado": f"Erro: {e}",
                             "modulo": "PPC", "enviado_por": meu_email()}
                        )
                        falhas += 1
                    progress.progress((i + 1) / len(selecionados))
                    status_box.text(f"{i+1}/{len(selecionados)} — {row['Nome']}")
                persistir_ppc(df_full[PPC_COLS])
                st.success(f"Concluído: {sucessos} enviados, {falhas} com erro. Estados guardados.")
                st.rerun()

        if st.session_state.log_envio:
            st.markdown("### Log de Envios")
            log_df = pd.DataFrame(st.session_state.log_envio)
            if "modulo" in log_df.columns:
                modulos = ["Todos"] + sorted(m for m in log_df["modulo"].unique() if m)
                filtro_mod = st.selectbox("Filtrar por módulo", modulos, key="log_filtro_ppc")
                if filtro_mod != "Todos":
                    log_df = log_df[log_df["modulo"] == filtro_mod]
            st.dataframe(log_df, use_container_width=True, height=250)

# --- Exportar --------------------------------------------------------------
with tab_export:
    st.subheader("Exportar Folha de Controlo (PPC)")
    st.caption("📌 Exportar é apenas um download — os teus dados continuam guardados na app depois disto.")
    if df_calc.empty:
        st.info("Ainda não há clientes elegíveis para PPC.")
    else:
        excel_bytes = gerar_excel_ppc(df_calc, st.session_state.params)
        st.download_button(
            "⬇️ Descarregar Excel de Controlo (com fórmulas e estados)",
            data=excel_bytes,
            file_name=f"Controlo_PPC_{ano_pag}_{date.today().isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.caption("O Excel inclui todos os cálculos, os estados de Guia Emitida / Email Enviado e fica destacado a verde para clientes dispensados.")

        if st.session_state.log_envio:
            log_csv = pd.DataFrame(st.session_state.log_envio).to_csv(index=False, sep=";")
            st.download_button("⬇️ Descarregar log de envios (CSV)", log_csv, file_name="log_envios.csv", mime="text/csv")

# Persistir templates caso o admin os tenha editado nesta execução (RLS bloqueia gestores).
guardar_config_db(st.session_state.params, st.session_state.templates, st.session_state.get("template_irs"))
