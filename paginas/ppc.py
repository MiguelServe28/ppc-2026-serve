"""
Página do PPC — Pagamentos por Conta. Tudo o que é específico deste imposto
vive aqui: parâmetros de cálculo, dados de PPC por cliente, guias, emails e
exportação. Só aparece quando esta página é selecionada no menu lateral.
"""

from datetime import date, datetime

import pandas as pd
import streamlit as st

from common import (
    PPC_COLS,
    calcular_ppc,
    enviar_email,
    escolher_conta_email,
    extrair_nif_de_filename,
    gerar_excel_ppc,
    guardar_config_db,
    montar_base_ppc,
    persistir_ppc,
    registar_log,
    render_template,
    sou_admin,
)

st.title("💶 PPC — Pagamentos por Conta 2026")
st.caption("SERVE — Contabilidade e Viabilização Empresarial")

with st.sidebar:
    st.header("Parâmetros de Cálculo (PPC)")
    p = st.session_state.params
    if sou_admin():
        p["limiar_volume"] = st.number_input("Limiar Volume de Negócios (€)", value=float(p["limiar_volume"]), step=10000.0)
        p["taxa_baixa"] = st.number_input("Taxa se Volume ≤ limiar", value=float(p["taxa_baixa"]), step=0.01, format="%.2f")
        p["taxa_alta"] = st.number_input("Taxa se Volume > limiar", value=float(p["taxa_alta"]), step=0.01, format="%.2f")
        p["limite_dispensa"] = st.number_input("Limite de dispensa (€)", value=float(p["limite_dispensa"]), step=10.0)
        st.divider()
        p["data1"] = st.date_input("Data limite 1.º Pagamento", value=p["data1"])
        p["data2"] = st.date_input("Data limite 2.º Pagamento", value=p["data2"])
        p["data3"] = st.date_input("Data limite 3.º Pagamento", value=p["data3"])
    else:
        st.caption("Estes parâmetros são definidos pelo administrador e aplicam-se a todos os clientes.")
        st.write(f"Limiar Volume de Negócios: **{p['limiar_volume']:,.2f} €**")
        st.write(f"Taxa ≤ limiar / > limiar: **{p['taxa_baixa']:.0%} / {p['taxa_alta']:.0%}**")
        st.write(f"Limite de dispensa: **{p['limite_dispensa']:,.2f} €**")
        st.divider()
        st.write(f"Data limite 1.º Pagamento: **{p['data1'].strftime('%d/%m/%Y')}**")
        st.write(f"Data limite 2.º Pagamento: **{p['data2'].strftime('%d/%m/%Y')}**")
        st.write(f"Data limite 3.º Pagamento: **{p['data3'].strftime('%d/%m/%Y')}**")
    st.divider()
    st.caption("Fórmula: Total PPC = (Coleta IRC − Retenções) × Taxa, repartido em 3 prestações iguais, cada uma arredondada por excesso para euro (art. 105.º CIRC). Dispensa se (Coleta − Retenções) < limite definido.")

df_calc = calcular_ppc(montar_base_ppc(), st.session_state.params)

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
            com_guia = int(sum((row["NIF"], n) in st.session_state.guias for _, row in elegiveis_dash.iterrows()))
            enviados = int(elegiveis_dash[f"Email{n}_Enviado"].sum()) if total else 0
            pendentes = total - enviados
            resumo.append({"Pagamento": f"{n}.º", "Total Elegíveis": total, "Guia Anexada (sessão atual)": com_guia,
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

        st.caption("💾 Os dados ficam guardados de forma persistente no Supabase. As guias em PDF carregadas na aba 'Guias' só existem durante a sessão atual — o estado 'Guia Emitida' fica guardado, mas o ficheiro em si tens de recarregar se voltares noutro dia.")

# --- Dados e Cálculo -----------------------------------------------------
with tab_dados:
    st.subheader("Dados de PPC")
    base_ppc = montar_base_ppc()
    if base_ppc.empty:
        st.info("Ainda não há clientes com 'Aplica PPC' ligado. Ativa esse interruptor na página 'Clientes'.")
    else:
        st.caption("Preenche aqui o Volume de Negócios, Coleta IRC e Retenções de 2025 dos clientes elegíveis para PPC. Os estados de guia/email são geridos nas abas 'Guias' e 'Emails'.")
        editor_ppc = st.data_editor(
            base_ppc[["NIF", "Nome"] + [c for c in PPC_COLS if c != "NIF"]],
            use_container_width=True,
            hide_index=True,
            disabled=["NIF", "Nome", "Guia1_Emitida", "Guia2_Emitida", "Guia3_Emitida",
                      "Email1_Enviado", "Email2_Enviado", "Email3_Enviado"],
            column_config={
                "Volume_2025": st.column_config.NumberColumn("Volume Negócios 2025 (campo 411)", format="%.2f"),
                "Coleta_2025": st.column_config.NumberColumn("Coleta IRC 2025 (campo 351)", format="%.2f"),
                "Retencoes_2025": st.column_config.NumberColumn("Retenções 2025 (campo 359)", format="%.2f"),
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
        c3.metric("Total PPC a cobrar", f"{df_calc.loc[~df_calc['Dispensado'], 'Total_PPC'].sum():,.2f} €".replace(",", "X").replace(".", ",").replace("X", "."))
        c4.metric("Valor médio / cliente", f"{df_calc.loc[~df_calc['Dispensado'], 'Total_PPC'].mean() if (~df_calc['Dispensado']).any() else 0:,.2f} €".replace(",", "X").replace(".", ",").replace("X", "."))

        def highlight_dispensado(row):
            return ["background-color: #E2EFDA" if row["Dispensado"] else "" for _ in row]

        show_cols = ["NIF", "Nome", "Volume_2025", "Coleta_2025", "Retencoes_2025",
                     "Base_Calculo", "Taxa", "Total_PPC", "Dispensado", "Pag1", "Pag2", "Pag3"]
        st.dataframe(
            df_calc[show_cols].style.apply(highlight_dispensado, axis=1).format(
                {"Volume_2025": "{:,.2f}", "Coleta_2025": "{:,.2f}", "Retencoes_2025": "{:,.2f}",
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
    st.caption("Carregue os PDFs das guias — se o nome do ficheiro contiver o NIF (9 dígitos), a associação é automática. Caso contrário, associe manualmente abaixo.")

    n_pag = st.selectbox("A que pagamento correspondem estas guias?", [1, 2, 3], format_func=lambda x: f"{x}.º Pagamento")
    up_guias = st.file_uploader("Carregar guias PDF", type=["pdf"], accept_multiple_files=True, key="up_guias")

    if up_guias:
        for f in up_guias:
            nif_detetado = extrair_nif_de_filename(f.name)
            st.session_state.guias[(nif_detetado or f.name, n_pag)] = (f.name, f.read())
        st.success(f"{len(up_guias)} ficheiro(s) carregado(s).")

    if not df_calc.empty:
        clientes_nifs = set(df_calc["NIF"].tolist())
        chaves_deste_pagamento = [k for k in st.session_state.guias.keys() if k[1] == n_pag]

        if chaves_deste_pagamento:
            st.markdown("**Associação manual / correção**")
            st.caption("Escolhe um ficheiro carregado e o cliente a quem pertence. Útil se o nome do PDF não tinha o NIF, ou se a associação automática ficou errada.")

            opcoes_ficheiro = {
                f"{st.session_state.guias[k][0]}"
                + (f"  (atualmente: sem cliente associado)" if k[0] not in clientes_nifs else f"  (atualmente: {k[0]})"):
                k
                for k in chaves_deste_pagamento
            }
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
                    chave_antiga = opcoes_ficheiro[ficheiro_escolhido]
                    filename, filebytes = st.session_state.guias.pop(chave_antiga)
                    st.session_state.guias[(cliente_escolhido, n_pag)] = (filename, filebytes)
                    st.success(f"'{filename}' associado a {cliente_escolhido}.")
                    st.rerun()

    if not df_calc.empty:
        st.markdown("**Estado das guias por cliente:**")
        rows = []
        for _, r in df_calc.iterrows():
            tem_guia = (r["NIF"], n_pag) in st.session_state.guias
            rows.append({"NIF": r["NIF"], "Nome": r["Nome"], "Guia carregada": "✅" if tem_guia else "❌",
                         f"Guia{n_pag}_Emitida (registo manual)": r[f"Guia{n_pag}_Emitida"]})
        st.dataframe(pd.DataFrame(rows), use_container_width=True, height=300)

        if st.button(f"Marcar Guia {n_pag} como Emitida para todos os clientes com PDF carregado"):
            df_full = montar_base_ppc()
            for idx, r in df_full.iterrows():
                if (r["NIF"], n_pag) in st.session_state.guias:
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

        with st.expander("✏️ Editar template deste email"):
            if sou_admin():
                tpl["assunto"] = st.text_input("Assunto", value=tpl["assunto"], key=f"assunto_{n_pag_email}")
                tpl["corpo"] = st.text_area("Corpo", value=tpl["corpo"], height=300, key=f"corpo_{n_pag_email}")
                st.caption("Placeholders disponíveis: {nome} {nif} {email} {pag1} {pag2} {pag3} {total} {data1} {data2} {data3}")
            else:
                st.caption("Os templates de email são definidos pelo administrador.")
                st.text_input("Assunto", value=tpl["assunto"], disabled=True)
                st.text_area("Corpo", value=tpl["corpo"], height=300, disabled=True)

        elegiveis = df_calc[~df_calc["Dispensado"]].copy()
        elegiveis = elegiveis[elegiveis["Email"].str.strip() != ""]

        st.markdown(f"**{len(elegiveis)} clientes elegíveis** (não dispensados, com email preenchido).")

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
            tem_guia = (row["NIF"], n_pag_email) in st.session_state.guias
            st.write("📎 Guia anexada:", "✅ Sim" if tem_guia else "❌ Não carregada (aba Guias)")

        st.divider()
        smtp_cfg = escolher_conta_email("ppc")

        com_guia = [n for n in elegiveis["NIF"].tolist() if (n, n_pag_email) in st.session_state.guias]
        sem_guia = [n for n in elegiveis["NIF"].tolist() if n not in com_guia]
        nao_enviados = [
            n for n in elegiveis["NIF"].tolist()
            if not df_calc.loc[df_calc["NIF"] == n, f"Email{n_pag_email}_Enviado"].iloc[0]
        ]

        st.markdown(f"📎 **{len(com_guia)} de {len(elegiveis)} clientes elegíveis já têm guia anexada** para este pagamento.")
        if sem_guia:
            st.caption(f"Sem guia anexada (não vão poder ser enviados com anexo): {len(sem_guia)} cliente(s) — associa-os na aba 'Guias'.")

        multiselect_key = f"selecionados_email_{n_pag_email}"

        col_btn1, col_btn2, col_btn3 = st.columns(3)
        with col_btn1:
            if st.button("📎 Selecionar só quem tem guia anexada"):
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
                sucessos, falhas = 0, 0
                for i, nif in enumerate(selecionados):
                    row = elegiveis[elegiveis["NIF"] == nif].iloc[0]
                    assunto, corpo = render_template(tpl, row, st.session_state.params)
                    anexos = []
                    guia = st.session_state.guias.get((nif, n_pag_email))
                    if guia:
                        anexos.append(guia)
                    try:
                        cc_gestor = [row["Gestor_Email"]] if row["Gestor_Email"] else []
                        enviar_email(smtp_cfg, row["Email"], assunto, corpo, anexos, cc=cc_gestor)
                        idx = df_full.index[df_full["NIF"] == nif][0]
                        df_full.at[idx, f"Email{n_pag_email}_Enviado"] = True
                        registar_log(
                            {"data": datetime.now().strftime("%Y-%m-%d %H:%M"), "nif": nif,
                             "nome": row["Nome"], "pagamento": n_pag_email, "estado": "Enviado"}
                        )
                        sucessos += 1
                    except Exception as e:
                        registar_log(
                            {"data": datetime.now().strftime("%Y-%m-%d %H:%M"), "nif": nif,
                             "nome": row["Nome"], "pagamento": n_pag_email, "estado": f"Erro: {e}"}
                        )
                        falhas += 1
                    progress.progress((i + 1) / len(selecionados))
                    status_box.text(f"{i+1}/{len(selecionados)} — {row['Nome']}")
                persistir_ppc(df_full[PPC_COLS])
                st.success(f"Concluído: {sucessos} enviados, {falhas} com erro. Estados guardados.")
                st.rerun()

        if st.session_state.log_envio:
            st.markdown("### Log de Envios")
            st.dataframe(pd.DataFrame(st.session_state.log_envio), use_container_width=True, height=250)

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
            file_name=f"Controlo_PPC_2026_{date.today().isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.caption("O Excel inclui todos os cálculos, os estados de Guia Emitida / Email Enviado e fica destacado a verde para clientes dispensados.")

        if st.session_state.log_envio:
            log_csv = pd.DataFrame(st.session_state.log_envio).to_csv(index=False, sep=";")
            st.download_button("⬇️ Descarregar log de envios (CSV)", log_csv, file_name="log_envios_ppc.csv", mime="text/csv")

# Persistir os parâmetros/templates do PPC caso o admin os tenha editado nesta execução (RLS bloqueia gestores).
guardar_config_db(st.session_state.params, st.session_state.templates, st.session_state.get("template_irs"))
