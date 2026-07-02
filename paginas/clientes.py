"""
Clientes — registo central, partilhado por toda a plataforma. Não tem nada
específico de nenhum imposto: NIF, contactos, gestor, tipos e os interruptores
que decidem em que páginas de imposto cada cliente aparece.
"""

import pandas as pd
import streamlit as st

from common import (
    PPC_COLS,
    PPC_NUM_COLS,
    carregar_ppc_db,
    clean_clientes_df,
    clean_ppc_df,
    ler_ficheiro_importacao,
    meu_email,
    persistir_clientes,
    persistir_clientes_parcial,
    persistir_ppc,
    sou_admin,
)

st.title("📋 Registo Central de Clientes")
st.caption(
    "Esta tabela é partilhada por toda a plataforma. Os 'Tipos' são só categorização. "
    "Os interruptores 'Aplica' são manuais e decidem em que módulos (PPC, IVA, IMI, IRS, Segurança Social) "
    "este cliente aparece — um cliente pode ter vários ligados ao mesmo tempo."
)
if not sou_admin():
    st.caption(f"Estás a ver apenas os clientes atribuídos a ti ({meu_email()}).")

col1, col2 = st.columns([2, 1])
with col1:
    up = st.file_uploader(
        "Importar CSV ou Excel (colunas mínimas: NIF, Nome, Email, Gestor_Nome, Gestor_Email; "
        "opcionalmente também Volume_2025, Coleta_2025, Retencoes_2025 para já entrarem no PPC)",
        type=["csv", "xlsx"],
    )
    if up is not None:
        try:
            bruto = ler_ficheiro_importacao(up)
            tem_dados_ppc = any(c in bruto.columns for c in PPC_NUM_COLS)

            novo_clientes = clean_clientes_df(bruto)
            if tem_dados_ppc:
                novo_clientes["Aplica_PPC"] = True  # importação clássica de PPC -> liga logo o interruptor

            novo_ppc = clean_ppc_df(bruto) if tem_dados_ppc else pd.DataFrame(columns=PPC_COLS)

            modo = st.radio("Modo de importação", ["Substituir tudo", "Adicionar aos existentes"], horizontal=True, key="modo_import")
            if st.button("Confirmar importação"):
                if modo == "Substituir tudo":
                    persistir_clientes(novo_clientes)
                else:
                    persistir_clientes(
                        clean_clientes_df(pd.concat([st.session_state.clientes, novo_clientes], ignore_index=True))
                        .drop_duplicates(subset="NIF", keep="last")
                    )
                # Sincroniza (pode ter havido eliminações em cascata de ppc_dados) e funde os novos dados de PPC.
                st.session_state.ppc_dados = carregar_ppc_db()
                if not novo_ppc.empty:
                    persistir_ppc(
                        clean_ppc_df(pd.concat([st.session_state.ppc_dados, novo_ppc], ignore_index=True))
                        .drop_duplicates(subset="NIF", keep="last")
                    )
                st.success(f"{len(novo_clientes)} clientes importados e guardados.")
                st.rerun()
        except Exception as e:
            st.error(f"Erro ao importar: {e}")
with col2:
    template_csv = pd.DataFrame(
        [{"NIF": "500123456", "Nome": "Empresa Exemplo, Lda.", "Email": "geral@exemplo.pt",
          "Gestor_Nome": "Ana Gestora", "Gestor_Email": "ana@serve.pt",
          "Volume_2025": 10000, "Coleta_2025": 2000, "Retencoes_2025": 200}]
    ).to_csv(index=False, sep=";")
    st.download_button("📥 Template CSV", template_csv, file_name="template_clientes.csv", mime="text/csv")

st.divider()
FILTROS_IMPOSTO = {
    "Todos": None,
    "Só PPC": "Aplica_PPC",
    "Só IVA": "Aplica_IVA",
    "Só IMI": "Aplica_IMI",
    "Só IRS": "Aplica_IRS",
    "Só Segurança Social": "Aplica_SS",
    "Sem nenhum imposto atribuído": "__nenhum__",
}
filtro_escolhido = st.selectbox(
    "Mostrar",
    list(FILTROS_IMPOSTO.keys()),
    key="filtro_clientes",
    help="Filtra a tabela abaixo — útil para veres só os clientes de um imposto (ex: os que só têm IRS) sem os teres misturados visualmente com os restantes.",
)
todos_clientes = clean_clientes_df(st.session_state.clientes)
coluna_filtro = FILTROS_IMPOSTO[filtro_escolhido]
if coluna_filtro is None:
    clientes_mostrados = todos_clientes
elif coluna_filtro == "__nenhum__":
    nenhum_aplica = ~(todos_clientes["Aplica_PPC"] | todos_clientes["Aplica_IVA"] | todos_clientes["Aplica_IMI"]
                       | todos_clientes["Aplica_IRS"] | todos_clientes["Aplica_SS"])
    clientes_mostrados = todos_clientes[nenhum_aplica]
else:
    clientes_mostrados = todos_clientes[todos_clientes[coluna_filtro]]

nifs_visiveis_antes = set(clientes_mostrados["NIF"])
st.caption(f"A mostrar {len(clientes_mostrados)} de {len(todos_clientes)} cliente(s) no total.")

st.markdown("**Tabela de clientes** — pode editar diretamente, adicionar ou apagar linhas.")
col_config = {
    "Tipo_Empresa": st.column_config.CheckboxColumn("Empresa"),
    "Tipo_AL": st.column_config.CheckboxColumn("Aloj. Local"),
    "Tipo_Trab_Independente": st.column_config.CheckboxColumn("Trab. Independente"),
    "Tipo_Rep_Fiscal": st.column_config.CheckboxColumn("Repr. Fiscal"),
    "Aplica_PPC": st.column_config.CheckboxColumn("PPC"),
    "Aplica_IVA": st.column_config.CheckboxColumn("IVA"),
    "Aplica_IMI": st.column_config.CheckboxColumn("IMI"),
    "Aplica_IRS": st.column_config.CheckboxColumn("IRS"),
    "Aplica_SS": st.column_config.CheckboxColumn("Seg. Social"),
}
if sou_admin():
    col_config["Gestor_Nome"] = st.column_config.TextColumn("Gestor (nome)")
    col_config["Gestor_Email"] = st.column_config.TextColumn("Gestor (email, vai em CC)")
else:
    col_config["Gestor_Nome"] = st.column_config.TextColumn("Gestor (nome)", disabled=True)
    col_config["Gestor_Email"] = st.column_config.TextColumn("Gestor (email)", disabled=True)

edited = st.data_editor(
    clientes_mostrados,
    num_rows="dynamic",
    use_container_width=True,
    column_config=col_config,
    key="editor_clientes",
)
if st.button("💾 Guardar alterações à tabela"):
    edited_final = clean_clientes_df(edited)
    if coluna_filtro not in (None, "__nenhum__"):
        # Uma linha nova adicionada enquanto este filtro está ativo assume-se que é para este imposto.
        novas = ~edited_final["NIF"].isin(nifs_visiveis_antes)
        edited_final.loc[novas, coluna_filtro] = True
    persistir_clientes_parcial(edited_final, nifs_visiveis_antes)
    st.success("Tabela atualizada e guardada — os dados ficam gravados mesmo depois de fechares o browser.")
    st.rerun()
