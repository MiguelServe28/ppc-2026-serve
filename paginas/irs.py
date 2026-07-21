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

import re
from datetime import date, datetime

import pandas as pd
import streamlit as st

from common import (
    DEFAULT_TEMPLATE_IRS,
    IRS_COLS,
    PASTAS_TIPO_DOC_IRS_AVULSO,
    carregar_clientes_irs_avulso_db,
    clean_clientes_df,
    clean_irs_avulso_df,
    clean_irs_df,
    detetar_categoria_irs_avulso,
    editor_template_bilingue,
    enviar_email,
    escolher_conta_email,
    extrair_dados_liquidacao_irs,
    extrair_dados_pendentes_irs,
    extrair_nif_de_filename,
    extrair_numero_de_filename,
    formatar_valor,
    gerar_excel_estado_mensal,
    guardar_config_db,
    ler_ficheiro_importacao,
    listar_extras_generico,
    marcar_irs_enviado_db,
    meu_email,
    migrar_documentos_irs_antigos,
    montar_base_irs,
    obter_documentos_irs,
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
    "Carrega os documentos (IRS, Liquidação, Guia, Fatura, Listagem de Pendentes) em massa ou "
    "cliente a cliente — a categoria é reconhecida automaticamente pelo nome do ficheiro e associada "
    "pelo NIF. O email lista os documentos anexados, sem mencionar valores. Tudo fica guardado no "
    "arquivo — não se perde ao fechar o browser."
)

base_irs = montar_base_irs()

tab_importar, tab_visao, tab_docs, tab_emails, tab_avulso, tab_template = st.tabs(
    ["📥 Importar Clientes", "📊 Estado", "📎 Documentos", "✉️ Emails", "🔢 IRS Avulso (por número)", "✏️ Template de Email"]
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

# --- Estado -----------------------------------------------------------
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

        arquivos_irs = {
            pasta: listar_extras_generico(f"irs/{ano_dados}/{pasta}")
            for pasta in PASTAS_TIPO_DOC_IRS_AVULSO.values()
        }

        def _marca_irs(pasta: str, nif: str) -> str:
            n = len(arquivos_irs.get(pasta, {}).get(nif, []))
            if n == 0:
                return "❌"
            return "✅" if n == 1 else f"✅ ({n})"

        linhas_estado_irs = []
        for _, r in mostrados.iterrows():
            linhas_estado_irs.append({
                "N.º": r.get("Numero_Cliente", ""), "NIF": r["NIF"], "Nome": r["Nome"],
                "IRS": _marca_irs("irs", r["NIF"]),
                "Liquidação": _marca_irs("liquidacao", r["NIF"]),
                "Guia": _marca_irs("guia", r["NIF"]),
                "Fatura": _marca_irs("fatura", r["NIF"]),
                "Pendentes": _marca_irs("pendentes", r["NIF"]),
                "Incluído na Avença": bool(r["Incluido_Avenca"]),
                "Email Enviado": bool(r["Email_Enviado"]),
            })
        st.caption("✏️ Podes marcar/desmarcar diretamente os piscos 'Incluído na Avença' e 'Email Enviado' — carrega em Guardar no fim.")
        editado = st.data_editor(
            pd.DataFrame(linhas_estado_irs),
            use_container_width=True,
            hide_index=True,
            height=400,
            disabled=["N.º", "NIF", "Nome", "IRS", "Liquidação", "Guia", "Fatura", "Pendentes"],
            column_config={
                "Incluído na Avença": st.column_config.CheckboxColumn("Incluído na Avença"),
                "Email Enviado": st.column_config.CheckboxColumn("Email Enviado"),
            },
            key=f"editor_visao_irs_{filtro_tipo}",
        )
        if st.button("💾 Guardar piscos"):
            atual = clean_irs_df(pd.DataFrame(st.session_state.irs_dados))
            novo = atual.copy()
            for _, r in editado.iterrows():
                nif_p = r["NIF"]
                if nif_p in set(novo["NIF"]):
                    novo.loc[novo["NIF"] == nif_p, "Incluido_Avenca"] = bool(r["Incluído na Avença"])
                    novo.loc[novo["NIF"] == nif_p, "Email_Enviado"] = bool(r["Email Enviado"])
                else:
                    nova_linha = pd.DataFrame([{
                        "NIF": nif_p, "Numero_Liquidacao": "", "Valor_Apurado": 0.0, "Valor_Pendente": 0.0,
                        "Incluido_Avenca": bool(r["Incluído na Avença"]), "Email_Enviado": bool(r["Email Enviado"]),
                    }])
                    novo = pd.concat([novo, nova_linha], ignore_index=True)
            persistir_irs(novo)
            st.success("Piscos guardados.")
            st.rerun()

        c1, c2 = st.columns(2)
        c1.metric("Clientes IRS (no filtro)", len(mostrados))
        c2.metric("Emails Enviados", f"{int(mostrados['Email_Enviado'].sum())} / {len(mostrados)}")

        st.divider()
        enviados_irs_excel = dict(zip(mostrados["NIF"], mostrados["Email_Enviado"]))
        excel_irs = gerar_excel_estado_mensal(
            f"Controlo IRS — {ano_dados}", mostrados,
            arquivos_irs.get("irs", {}), arquivos_irs.get("liquidacao", {}), {}, enviados_irs_excel,
            rotulo_guia="IRS", rotulo_decl="Liquidação",
            extra_categorias=[("Guia", arquivos_irs.get("guia", {})), ("Fatura", arquivos_irs.get("fatura", {})),
                               ("Pendentes", arquivos_irs.get("pendentes", {}))],
        )
        st.download_button(
            "⬇️ Descarregar Excel de Controlo (IRS)",
            data=excel_irs,
            file_name=f"Controlo_IRS_{ano_dados}_{date.today().isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.caption("Clientes com email já enviado ficam destacados a verde no Excel.")

# --- Documentos ---------------------------------------------------------
with tab_docs:
    if base_irs.empty:
        st.info("Ainda não há clientes com 'Aplica IRS' ligado — importa-os na aba 'Importar Clientes' ou ativa o interruptor na página 'Clientes'.")
    else:
        arquivos_irs_docs = {
            pasta: listar_extras_generico(f"irs/{ano_dados}/{pasta}")
            for pasta in PASTAS_TIPO_DOC_IRS_AVULSO.values()
        }

        with st.expander("🔁 Migrar documentos antigos (Guia/Fatura carregados na versão anterior desta página)"):
            st.caption(
                "Antes, a Guia e a Fatura ficavam guardadas de forma diferente (uma pasta por cliente). "
                "Se já tinhas carregado alguma antes desta atualização, usa este botão uma vez para as "
                "trazer para o sítio novo — não apaga nem duplica nada, só reorganiza."
            )
            if st.button("Migrar agora", key="irs_migrar_antigos"):
                migrados = migrar_documentos_irs_antigos(ano_dados)
                st.success(f"Migrados: {migrados['guia']} guia(s), {migrados['fatura']} fatura(s).")
                st.rerun()

        total_docs_ano_irs = sum(len(nomes) for dic in arquivos_irs_docs.values() for nomes in dic.values())
        if total_docs_ano_irs:
            with st.expander(f"🗑️ Apagar TODOS os documentos de TODOS os clientes de {ano_dados} ({total_docs_ano_irs})"):
                st.caption(
                    "Apaga de uma vez todos os documentos (IRS, Liquidação, Guia, Fatura e Pendentes) de "
                    "todos os clientes de IRS, só para este ano. Não apaga os clientes em si, só os ficheiros."
                )
                if st.button(f"Confirmar — apagar os {total_docs_ano_irs} documento(s) de {ano_dados}",
                             key=f"irs_apagar_tudo_ano_{ano_dados}", type="primary"):
                    for pasta_x, dic_x in arquivos_irs_docs.items():
                        for nif_x, nomes_x in dic_x.items():
                            for nome_x in nomes_x:
                                storage_apagar(f"irs/{ano_dados}/{pasta_x}/{nif_x}__{nome_x}")
                    st.success("Todos os documentos deste ano foram apagados.")
                    st.rerun()
            st.divider()

        st.markdown("**Carregamento em massa** (o nome do ficheiro deve começar pelo NIF de 9 dígitos, ex: '123456789 - Guia - Miguel Silva.pdf')")
        st.caption(
            "A categoria (IRS, Liquidação, Guia, Fatura ou Listagem de Pendentes) é reconhecida "
            "automaticamente pelo resto do nome do ficheiro — não precisas de escolher antes, podes "
            "carregar tudo misturado de uma vez. Carregar um novo ficheiro para um cliente/categoria "
            "SUBSTITUI o(s) ficheiro(s) que já lá estavam dessa categoria — não precisas de apagar antes "
            "de corrigir. Se quiseres mesmo mais do que um ficheiro do mesmo tipo para o mesmo cliente, "
            "carrega-os todos DE UMA VEZ (ficam numerados)."
        )
        up_massa_irs = st.file_uploader(
            "Carregar PDFs (nome a começar pelo NIF; a categoria é adivinhada pelo resto do nome)",
            type=["pdf"], accept_multiple_files=True, key="irs_up_massa",
        )
        if up_massa_irs:
            ids_up_irs = tuple(sorted(f"{f.name}_{f.size}" for f in up_massa_irs))
            if st.session_state.get("_irs_massa_proc") != (ano_dados, ids_up_irs):
                st.session_state["_irs_massa_proc"] = (ano_dados, ids_up_irs)
                ok_irs, sem_nif_irs, sem_categoria_irs, detalhes_irs = 0, [], [], []
                ficheiros_por_grupo_irs = {}
                for f in up_massa_irs:
                    nif_f = extrair_nif_de_filename(f.name)
                    if not nif_f:
                        sem_nif_irs.append(f.name)
                        continue
                    pasta_f = detetar_categoria_irs_avulso(f.name)
                    if not pasta_f:
                        sem_categoria_irs.append(f.name)
                        continue
                    ficheiros_por_grupo_irs.setdefault((nif_f, pasta_f), []).append(f)
                for (nif_f, pasta_ir), ficheiros in ficheiros_por_grupo_irs.items():
                    for nome_antigo in arquivos_irs_docs.get(pasta_ir, {}).get(nif_f, []):
                        storage_apagar(f"irs/{ano_dados}/{pasta_ir}/{nif_f}__{nome_antigo}")
                    nomes_novos = ([f"{pasta_ir}.pdf"] if len(ficheiros) == 1
                                    else [f"{pasta_ir}_{i}.pdf" for i in range(1, len(ficheiros) + 1)])
                    for f, nome_novo in zip(ficheiros, nomes_novos):
                        caminho_f = f"irs/{ano_dados}/{pasta_ir}/{nif_f}__{nome_novo}"
                        try:
                            storage_upload_pdf(caminho_f, f.getvalue())
                            detalhes_irs.append(f"✅ {nif_f} ({pasta_ir}) → {caminho_f}")
                            ok_irs += 1
                        except Exception as e:
                            detalhes_irs.append(f"❌ {nif_f} ({pasta_ir}) → {caminho_f}: {e}")
                msg_irs = f"{ok_irs} ficheiro(s) associados e guardados no arquivo."
                if sem_nif_irs:
                    msg_irs += f" Sem NIF no nome (ignorados): {', '.join(sem_nif_irs)}"
                if sem_categoria_irs:
                    msg_irs += f" Categoria não reconhecida no nome (usa o carregamento por cliente, abaixo): {', '.join(sem_categoria_irs)}"
                st.session_state["_irs_ultimo_upload_massa"] = {"msg": msg_irs, "detalhes": detalhes_irs}
                st.rerun()

        ultimo_irs = st.session_state.get("_irs_ultimo_upload_massa")
        if ultimo_irs:
            st.success(ultimo_irs["msg"])
            if ultimo_irs["detalhes"]:
                with st.expander(f"Ver os {len(ultimo_irs['detalhes'])} caminho(s) exatos onde foi guardado"):
                    st.text("\n".join(ultimo_irs["detalhes"]))

        st.divider()
        st.markdown("**Carregamento por cliente**")
        nif_doc_irs = st.selectbox(
            "Cliente", base_irs["NIF"].tolist(),
            format_func=lambda n: f"{n} — {base_irs.loc[base_irs['NIF']==n,'Nome'].values[0]}",
            key="irs_cliente_doc",
        )
        total_docs_cliente_irs = sum(len(d.get(nif_doc_irs, [])) for d in arquivos_irs_docs.values())
        if total_docs_cliente_irs:
            with st.expander(f"🗑️ Apagar todos os documentos deste cliente ({total_docs_cliente_irs})"):
                st.caption("Apaga TODOS os documentos (IRS, Liquidação, Guia, Fatura e Pendentes) deste cliente, só para este ano.")
                if st.button("Confirmar — apagar tudo", key=f"irs_apagar_tudo_cliente_{ano_dados}_{nif_doc_irs}", type="primary"):
                    for pasta_x, dic_x in arquivos_irs_docs.items():
                        for nome_x in dic_x.get(nif_doc_irs, []):
                            storage_apagar(f"irs/{ano_dados}/{pasta_x}/{nif_doc_irs}__{nome_x}")
                    st.success("Documentos apagados.")
                    st.rerun()

        st.caption(
            "📂 Carrega de uma vez TODOS os documentos deste cliente, misturados — a categoria "
            "(IRS, Liquidação, Guia, Fatura ou Pendentes) é reconhecida automaticamente pelo nome de cada "
            "ficheiro (não precisas de separar por tipo nem de escolher o NIF, já está selecionado acima). "
            "Carregar de novo substitui o que já lá estava dessa categoria para este cliente."
        )
        up_cliente_irs = st.file_uploader(
            "Carregar ficheiros deste cliente (todos de uma vez)",
            type=["pdf"], accept_multiple_files=True, key=f"irs_up_cliente_{ano_dados}_{nif_doc_irs}",
        )
        if up_cliente_irs:
            ids_up_cli = tuple(sorted(f"{f.name}_{f.size}" for f in up_cliente_irs))
            chave_proc_cli = f"_irs_up_cliente_proc_{ano_dados}_{nif_doc_irs}"
            if st.session_state.get(chave_proc_cli) != ids_up_cli:
                st.session_state[chave_proc_cli] = ids_up_cli
                grupos_cli = {}
                sem_categoria_cli = []
                for f in up_cliente_irs:
                    pasta_f = detetar_categoria_irs_avulso(f.name)
                    if not pasta_f:
                        sem_categoria_cli.append(f.name)
                        continue
                    grupos_cli.setdefault(pasta_f, []).append(f)
                for pasta_ir, ficheiros in grupos_cli.items():
                    for nome_antigo in arquivos_irs_docs.get(pasta_ir, {}).get(nif_doc_irs, []):
                        storage_apagar(f"irs/{ano_dados}/{pasta_ir}/{nif_doc_irs}__{nome_antigo}")
                    nomes_novos = ([f"{pasta_ir}.pdf"] if len(ficheiros) == 1
                                    else [f"{pasta_ir}_{i}.pdf" for i in range(1, len(ficheiros) + 1)])
                    for f, nome_novo in zip(ficheiros, nomes_novos):
                        storage_upload_pdf(f"irs/{ano_dados}/{pasta_ir}/{nif_doc_irs}__{nome_novo}", f.getvalue())
                msg_cli = f"{sum(len(v) for v in grupos_cli.values())} ficheiro(s) guardados para este cliente."
                if sem_categoria_cli:
                    msg_cli += f" Categoria não reconhecida no nome (não guardados — usa a correção manual abaixo): {', '.join(sem_categoria_cli)}"
                st.session_state["_irs_ultimo_upload_cliente"] = msg_cli
                st.rerun()

        ultimo_cli_irs = st.session_state.pop("_irs_ultimo_upload_cliente", None)
        if ultimo_cli_irs:
            st.success(ultimo_cli_irs)

        def _documentos_irs(col, rotulo: str, pasta: str, dicionario: dict):
            """Upload (aceita vários ficheiros) + lista com botão de apagar por
            ficheiro, para uma categoria do IRS normal — mesmo padrão já usado
            na Segurança Social e no IRS Avulso: carregar de novo SUBSTITUI o(s)
            ficheiro(s) anteriores; carregar vários de uma vez mantém-nos todos
            (numerados)."""
            with col:
                up = st.file_uploader(rotulo, type=["pdf"], accept_multiple_files=True,
                                       key=f"irs_up_{pasta}_{ano_dados}_{nif_doc_irs}")
                if up:
                    ids_up = tuple(sorted(f"{f.name}_{f.size}" for f in up))
                    chave_proc = f"_irs_{pasta}_proc_{ano_dados}_{nif_doc_irs}"
                    if st.session_state.get(chave_proc) != ids_up:
                        st.session_state[chave_proc] = ids_up
                        for nome_antigo in dicionario.get(nif_doc_irs, []):
                            storage_apagar(f"irs/{ano_dados}/{pasta}/{nif_doc_irs}__{nome_antigo}")
                        nomes_novos = ([f"{pasta}.pdf"] if len(up) == 1
                                        else [f"{pasta}_{i}.pdf" for i in range(1, len(up) + 1)])
                        for f, nome_novo in zip(up, nomes_novos):
                            storage_upload_pdf(f"irs/{ano_dados}/{pasta}/{nif_doc_irs}__{nome_novo}", f.getvalue())
                        st.success(f"{len(up)} ficheiro(s) de {rotulo} guardado(s). Substituiu os anteriores.")
                        st.rerun()
                existentes = dicionario.get(nif_doc_irs, [])
                if existentes:
                    for nome in existentes:
                        c_nome, c_apagar = st.columns([4, 1])
                        c_nome.caption(f"📄 {nome}")
                        if c_apagar.button("🗑️", key=f"irs_apagar_{pasta}_{ano_dados}_{nif_doc_irs}_{nome}",
                                            help="Apagar (depois podes carregar outro em substituição)"):
                            storage_apagar(f"irs/{ano_dados}/{pasta}/{nif_doc_irs}__{nome}")
                            st.rerun()
                else:
                    st.caption(f"Sem {rotulo}")

        with st.expander("✏️ Correção manual — carregar/ver por categoria, uma a uma"):
            cols_irs = st.columns(5)
            for col_irs, (rotulo_irs, pasta_irs2) in zip(cols_irs, PASTAS_TIPO_DOC_IRS_AVULSO.items()):
                _documentos_irs(col_irs, rotulo_irs, pasta_irs2, arquivos_irs_docs.get(pasta_irs2, {}))

# --- Emails --------------------------------------------------------------
with tab_emails:
    if base_irs.empty:
        st.info("Ainda não há clientes com 'Aplica IRS' ligado — importa-os na aba 'Importar Clientes' ou ativa o interruptor na página 'Clientes'.")
    else:
        st.subheader(f"Enviar Emails — IRS {ano_dados}")
        arquivos_irs_email = {
            pasta: listar_extras_generico(f"irs/{ano_dados}/{pasta}")
            for pasta in PASTAS_TIPO_DOC_IRS_AVULSO.values()
        }
        elegiveis_irs = base_irs[base_irs["Email"].str.strip() != ""].copy()
        sem_email_irs = len(base_irs) - len(elegiveis_irs)
        if sem_email_irs:
            st.caption(f"⚠️ {sem_email_irs} cliente(s) sem email preenchido — não aparecem abaixo.")

        tpl_irs = st.session_state.template_irs

        preview_nif_irs = st.selectbox(
            "Pré-visualizar cliente:", elegiveis_irs["NIF"].tolist(),
            format_func=lambda n: f"{n} — {elegiveis_irs.loc[elegiveis_irs['NIF']==n,'Nome'].values[0]}",
            key="irs_preview_email",
        )
        docs_prev_irs = []
        if preview_nif_irs:
            row_prev_irs = elegiveis_irs[elegiveis_irs["NIF"] == preview_nif_irs].iloc[0]
            docs_prev_irs = obter_documentos_irs(ano_dados, preview_nif_irs, arquivos_irs_email)
            assunto_irs, corpo_irs = render_template_irs(tpl_irs, row_prev_irs, docs_prev_irs)
            st.text_input("Assunto (preview)", value=assunto_irs, disabled=True)
            if row_prev_irs["Gestor_Email"]:
                st.caption(f"📋 CC: {row_prev_irs['Gestor_Nome'] or ''} <{row_prev_irs['Gestor_Email']}>")
            else:
                st.caption("📋 CC: — (sem gestor definido)")
            st.text_area("Corpo (preview)", value=corpo_irs, height=260, disabled=True)
            st.caption("📎 Anexos: " + (", ".join(d["tipo"] for d in docs_prev_irs) if docs_prev_irs else "nenhum documento carregado ainda"))

            with st.expander("🔍 Diagnóstico deste cliente"):
                if docs_prev_irs:
                    for d in docs_prev_irs:
                        st.text(f"{d['tipo']}   →   {d['caminho']}")
                else:
                    st.caption("Nenhum documento encontrado para este cliente.")

        st.divider()
        smtp_cfg_irs = escolher_conta_email("irs")

        nao_enviados_irs = [n for n in elegiveis_irs["NIF"] if not elegiveis_irs.loc[elegiveis_irs["NIF"] == n, "Email_Enviado"].values[0]]
        com_docs_irs = [n for n in elegiveis_irs["NIF"] if any(n in dic for dic in arquivos_irs_email.values())]

        st.markdown(f"📎 **{len(com_docs_irs)} de {len(elegiveis_irs)}** cliente(s) já têm documentos carregados para IRS {ano_dados}.")

        multiselect_key_irs = f"irs_selecionados_{ano_dados}"
        col_b1_irs, col_b2_irs, col_b3_irs = st.columns(3)
        with col_b1_irs:
            if st.button("📎 Selecionar quem tem documentos e falta enviar", key="irs_sel_docs"):
                st.session_state[multiselect_key_irs] = [n for n in com_docs_irs if n in nao_enviados_irs]
                st.rerun()
        with col_b2_irs:
            if st.button("☑️ Selecionar todos por enviar", key="irs_sel_todos"):
                st.session_state[multiselect_key_irs] = nao_enviados_irs
                st.rerun()
        with col_b3_irs:
            if st.button("✖️ Limpar seleção", key="irs_sel_limpar"):
                st.session_state[multiselect_key_irs] = []
                st.rerun()

        if multiselect_key_irs not in st.session_state:
            st.session_state[multiselect_key_irs] = [n for n in com_docs_irs if n in nao_enviados_irs]

        selecionados_irs = st.multiselect(
            "Clientes selecionados para envio (podes ajustar — para enviar só um, deixa só esse)",
            elegiveis_irs["NIF"].tolist(),
            format_func=lambda n: f"{n} — {elegiveis_irs.loc[elegiveis_irs['NIF']==n,'Nome'].values[0]}"
            + ("" if n in com_docs_irs else "  ⚠️ sem documentos")
            + ("  ✅ já enviado" if elegiveis_irs.loc[elegiveis_irs["NIF"] == n, "Email_Enviado"].values[0] else ""),
            key=multiselect_key_irs,
        )

        if st.button("🚀 Enviar Emails Selecionados", type="primary", disabled=not selecionados_irs):
            if not smtp_cfg_irs["utilizador"] or not smtp_cfg_irs["password"]:
                st.error("Escolhe ou cria uma conta de email com utilizador e password preenchidos.")
            else:
                progress_irs = st.progress(0.0)
                status_irs = st.empty()
                assinatura_irs = st.session_state.params.get("assinatura_html", "")
                sucessos_irs, falhas_irs = 0, 0
                for i, nif_e in enumerate(selecionados_irs):
                    row_e_irs = elegiveis_irs[elegiveis_irs["NIF"] == nif_e].iloc[0]
                    docs_e_irs = obter_documentos_irs(ano_dados, nif_e, arquivos_irs_email)
                    anexos_e_irs = []
                    for d in docs_e_irs:
                        conteudo_e_irs = storage_download_pdf(d["caminho"])
                        if conteudo_e_irs:
                            anexos_e_irs.append((d["anexo"], conteudo_e_irs))
                    assunto_e_irs, corpo_e_irs = render_template_irs(tpl_irs, row_e_irs, docs_e_irs)
                    try:
                        cc_gestor_e_irs = [row_e_irs["Gestor_Email"]] if row_e_irs["Gestor_Email"] else []
                        enviar_email(smtp_cfg_irs, row_e_irs["Email"], assunto_e_irs, corpo_e_irs, anexos_e_irs,
                                     cc=cc_gestor_e_irs, bcc=[smtp_cfg_irs["remetente"]], assinatura_html=assinatura_irs)
                        marcar_irs_enviado_db(nif_e, ano_dados, True)
                        registar_log({
                            "data": datetime.now().strftime("%Y-%m-%d %H:%M"), "nif": nif_e,
                            "nome": row_e_irs["Nome"], "pagamento": 0, "estado": "Enviado",
                            "modulo": "IRS", "enviado_por": meu_email(),
                        })
                        sucessos_irs += 1
                    except Exception as e:
                        registar_log({
                            "data": datetime.now().strftime("%Y-%m-%d %H:%M"), "nif": nif_e,
                            "nome": row_e_irs["Nome"], "pagamento": 0, "estado": f"Erro: {e}",
                            "modulo": "IRS", "enviado_por": meu_email(),
                        })
                        falhas_irs += 1
                    progress_irs.progress((i + 1) / len(selecionados_irs))
                    status_irs.text(f"{i+1}/{len(selecionados_irs)} — {row_e_irs['Nome']}")
                st.success(f"Concluído: {sucessos_irs} enviados, {falhas_irs} com erro. Estados guardados.")
                st.rerun()

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
                # Reconhece a coluna do número mesmo com nomes tipo "N.º", "Nº", "N°",
                # "Número" — e também variações corrompidas por problemas de encoding
                # em Excel (ex: "N.Âº"), normalizando para só as letras ASCII antes
                # de comparar, em vez de exigir o texto exato.
                def _normalizar_col(c):
                    return re.sub(r"[^A-Za-z]", "", str(c)).lower()
                nomes_numero = {"n", "no", "num", "numero", "nmero"}
                bruto_av = bruto_av.rename(columns={
                    c: "Numero" for c in bruto_av.columns if _normalizar_col(c) in nomes_numero
                })
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
            COLS_EDIT_AV = ["Numero", "NIF", "Nome", "Email", "Lingua", "Gestor_Nome", "Gestor_Email",
                            "Numero_Liquidacao", "Valor_Apurado", "Valor_Pendente"]
            st.markdown(
                "**Clientes já importados** (edita Nome/Email/Língua/Gestor diretamente se precisares — "
                "e também o Nº de Liquidação/Valor Apurado/Valor Pendente, para os casos em que a leitura "
                "automática do PDF não encontrou o valor: sem isto preenchido, o email não afirma nada "
                "sobre o valor, em vez de arriscar dizer 'sem valor a pagar' por engano)"
            )
            tabela_edit_av = base_avulso[COLS_EDIT_AV].copy()
            # Número como inteiro só aqui na tabela (não na base) — para o clique no
            # cabeçalho da coluna ordenar numericamente (1, 2, 10) em vez de por texto
            # (1, 10, 100, 101...). É convertido de volta a texto antes de gravar.
            tabela_edit_av["Numero"] = pd.to_numeric(tabela_edit_av["Numero"], errors="coerce").astype("Int64")
            edit_av = st.data_editor(
                tabela_edit_av,
                use_container_width=True, hide_index=True, height=300,
                disabled=["Numero"],
                column_config={
                    "Valor_Apurado": st.column_config.NumberColumn("Valor Apurado (€)", format="%.2f", step=0.01),
                    "Valor_Pendente": st.column_config.NumberColumn("Valor Pendente (€)", format="%.2f", step=0.01),
                },
                key=f"editor_irs_avulso_clientes_{ano_dados}",
            )
            if st.button("💾 Guardar alterações aos clientes de IRS avulso"):
                edit_av_gravar = edit_av.copy()
                edit_av_gravar["Numero"] = edit_av_gravar["Numero"].astype(str)
                restante = base_avulso.drop(columns=COLS_EDIT_AV[1:]).merge(
                    edit_av_gravar, on="Numero", how="left"
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
            total_docs_ano = sum(len(nomes) for dic in arquivos_avulso.values() for nomes in dic.values())
            if total_docs_ano:
                with st.expander(f"🗑️ Apagar TODOS os documentos de TODOS os clientes de {ano_dados} ({total_docs_ano})"):
                    st.caption(
                        "Apaga de uma vez todos os documentos (IRS, Liquidação, Guia, Fatura e Pendentes) "
                        "de todos os clientes de IRS Avulso, só para este ano — útil para recomeçar limpo "
                        "antes de re-testar um carregamento em massa. Não apaga os clientes em si, só os "
                        "ficheiros."
                    )
                    if st.button(f"Confirmar — apagar os {total_docs_ano} documento(s) de {ano_dados}",
                                 key=f"irs_avulso_apagar_tudo_ano_{ano_dados}", type="primary"):
                        for pasta_x, dic_x in arquivos_avulso.items():
                            for numero_x, nomes_x in dic_x.items():
                                for nome_x in nomes_x:
                                    storage_apagar(f"irs_avulso/{ano_dados}/{pasta_x}/{numero_x}__{nome_x}")
                        st.success("Todos os documentos deste ano foram apagados.")
                        st.rerun()
                st.divider()

            st.markdown("**Carregamento em massa** (o nome do ficheiro deve começar pelo número, ex: '1 - Guia - Miguel Silva.pdf')")
            st.caption(
                "A categoria (IRS, Liquidação, Guia, Fatura ou Listagem de Pendentes) é reconhecida "
                "automaticamente pelo resto do nome do ficheiro — não precisas de escolher antes, podes "
                "carregar tudo misturado de uma vez. Carregar um novo ficheiro para um cliente/categoria "
                "SUBSTITUI o(s) ficheiro(s) que já lá estavam dessa categoria — não precisas de apagar "
                "antes de corrigir. Se quiseres mesmo mais do que um ficheiro do mesmo tipo para o mesmo "
                "cliente, carrega-os todos DE UMA VEZ (ficam numerados)."
            )
            up_massa_av = st.file_uploader(
                "Carregar PDFs (nome a começar pelo número; a categoria é adivinhada pelo resto do nome)",
                type=["pdf"], accept_multiple_files=True, key="irs_avulso_up_massa",
            )
            if up_massa_av:
                ids_up_av = tuple(sorted(f"{f.name}_{f.size}" for f in up_massa_av))
                if st.session_state.get("_irs_avulso_massa_proc") != (ano_dados, ids_up_av):
                    st.session_state["_irs_avulso_massa_proc"] = (ano_dados, ids_up_av)
                    ok_av, sem_numero, sem_categoria, detalhes_av = 0, [], [], []
                    atualizacoes_liq = {}   # numero -> {"Numero_Liquidacao":..., "Valor_Apurado":...}
                    atualizacoes_pend = {}  # numero -> valor_pendente
                    ficheiros_por_grupo = {}  # (numero, pasta) -> [ficheiros]
                    for f in up_massa_av:
                        numero_f = extrair_numero_de_filename(f.name)
                        if not numero_f:
                            sem_numero.append(f.name)
                            continue
                        pasta_f = detetar_categoria_irs_avulso(f.name)
                        if not pasta_f:
                            sem_categoria.append(f.name)
                            continue
                        ficheiros_por_grupo.setdefault((numero_f, pasta_f), []).append(f)
                    for (numero_f, pasta_av), ficheiros in ficheiros_por_grupo.items():
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
                                detalhes_av.append(f"✅ {numero_f} ({pasta_av}) → {caminho_f}")
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
                                detalhes_av.append(f"❌ {numero_f} ({pasta_av}) → {caminho_f}: {e}")

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
                    if sem_categoria:
                        msg_av += f" Categoria não reconhecida no nome (usa o carregamento por cliente, abaixo): {', '.join(sem_categoria)}"
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
                    # Número como inteiro (não texto) para poderes clicar no cabeçalho
                    # da coluna e ordenar numericamente (1, 2, 10...), não por texto
                    # (que puxava o "10" para antes do "2").
                    "Número": int(r["Numero"]) if str(r["Numero"]).isdigit() else r["Numero"],
                    "NIF": r["NIF"], "Nome": r["Nome"],
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
                assunto_av, corpo_av = render_template_irs(tpl_av, row_prev_av, docs_prev_av)
                st.text_input("Assunto (preview)", value=assunto_av, disabled=True)
                if row_prev_av["Gestor_Email"]:
                    st.caption(f"📋 CC: {row_prev_av['Gestor_Nome'] or ''} <{row_prev_av['Gestor_Email']}>")
                else:
                    st.caption("📋 CC: — (sem gestor definido)")
                st.text_area("Corpo (preview)", value=corpo_av, height=260, disabled=True)
                st.caption("📎 Anexos: " + (", ".join(d["tipo"] for d in docs_prev_av) if docs_prev_av else "nenhum documento carregado ainda"))
                if not row_prev_av["Numero_Liquidacao"] and not row_prev_av["Valor_Apurado"]:
                    st.caption("⚠️ Ainda não há valor apurado nem n.º de liquidação confirmados para este cliente — o email não vai afirmar nada sobre o valor (fica em branco). Se já carregaste a Liquidação e não leu automaticamente, corrige o valor à mão na tabela da aba 'Importar' (em baixo, 'Clientes já importados').")

                with st.expander("🔍 Diagnóstico deste cliente"):
                    if docs_prev_av:
                        for d in docs_prev_av:
                            st.text(f"{d['tipo']}   →   {d['caminho']}")
                    else:
                        st.caption("Nenhum documento encontrado para este número.")

            st.divider()
            smtp_cfg_av = escolher_conta_email("irs_avulso")

            nao_enviados_av = [n for n in elegiveis_av["Numero"] if not elegiveis_av.loc[elegiveis_av["Numero"] == n, "Email_Enviado"].values[0]]
            com_docs_av = [n for n in elegiveis_av["Numero"] if any(n in dic for dic in arquivos_avulso.values())]

            st.markdown(f"📎 **{len(com_docs_av)} de {len(elegiveis_av)}** cliente(s) já têm documentos carregados para IRS Avulso {ano_dados}.")

            multiselect_key_av = f"irs_avulso_selecionados_{ano_dados}"
            col_b1_av, col_b2_av, col_b3_av = st.columns(3)
            with col_b1_av:
                if st.button("📎 Selecionar quem tem documentos e falta enviar", key="irs_avulso_sel_docs"):
                    st.session_state[multiselect_key_av] = [n for n in com_docs_av if n in nao_enviados_av]
                    st.rerun()
            with col_b2_av:
                if st.button("☑️ Selecionar todos por enviar", key="irs_avulso_sel_todos"):
                    st.session_state[multiselect_key_av] = nao_enviados_av
                    st.rerun()
            with col_b3_av:
                if st.button("✖️ Limpar seleção", key="irs_avulso_sel_limpar"):
                    st.session_state[multiselect_key_av] = []
                    st.rerun()

            if multiselect_key_av not in st.session_state:
                st.session_state[multiselect_key_av] = [n for n in com_docs_av if n in nao_enviados_av]

            selecionados_av = st.multiselect(
                "Clientes selecionados para envio (podes ajustar — para enviar só um, deixa só esse)",
                elegiveis_av["Numero"].tolist(),
                format_func=lambda n: f"{n} — {elegiveis_av.loc[elegiveis_av['Numero']==n,'Nome'].values[0]}"
                + ("" if n in com_docs_av else "  ⚠️ sem documentos")
                + ("  ✅ já enviado" if elegiveis_av.loc[elegiveis_av["Numero"] == n, "Email_Enviado"].values[0] else ""),
                key=multiselect_key_av,
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
                        assunto_e, corpo_e = render_template_irs(tpl_av, row_e, docs_e)
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
    if st.button("🔄 Repor template padrão", key="irs_tpl_reset",
                 help="Substitui o texto atual pelo modelo mais recente (ex: lista de documentos anexados, no IRS Avulso)."):
        st.session_state.template_irs = DEFAULT_TEMPLATE_IRS.copy()
        st.rerun()
    tpl = st.session_state.template_irs
    editor_template_bilingue(tpl, "irs_tpl", altura=320)
    st.caption(
        "Placeholders disponíveis: {nome} {nif} {email} {ref_liquidacao} {frase_valor} {frase_pendente} {lista_docs} {ano_dados} {ano_pagamentos}. "
        "{ref_liquidacao} já vem formatado como ', n.º de liquidação XXXX' (ou vazio, se não houver). "
        "{frase_valor} e {frase_pendente} são frases já prontas, geradas automaticamente a partir dos valores e na língua do cliente — não precisas de os escrever à mão. "
        "{lista_docs} no IRS normal é a frase genérica sobre a guia; no IRS Avulso lista automaticamente os documentos "
        "carregados para cada cliente (IRS, Liquidação, Guia, Fatura, Pendentes), em formato de lista por pontos. "
        "Alterações aqui ficam guardadas para toda a equipa."
    )

# Persistir template (guardado para toda a equipa, qualquer utilizador pode editar).
guardar_config_db(st.session_state.params, st.session_state.templates, st.session_state.template_irs)
