"""
Gestores — só visível para admin. Criar/gerir contas de login. Comum a toda a
plataforma (não é específico de nenhum imposto).
"""

import pandas as pd
import streamlit as st

from common import SUPABASE_SERVICE_KEY, get_admin_client, get_client, sou_admin

if not sou_admin():
    st.error("Esta página é exclusiva do administrador.")
    st.stop()

st.title("👥 Contas de Gestor")
st.caption("Cria e gere as contas de login dos gestores da SERVE. Depois de criares a conta, atribui os clientes a este gestor na página 'Clientes' (campo Gestor_Email igual ao email de login abaixo).")

client = get_client()
perfis_resp = client.table("perfis").select("*").order("email").execute()
perfis_lista = perfis_resp.data or []

if perfis_lista:
    df_perfis = pd.DataFrame(perfis_lista)[["id", "email", "nome", "role"]].rename(
        columns={"email": "Email", "nome": "Nome", "role": "Papel"}
    )
    if not SUPABASE_SERVICE_KEY:
        st.dataframe(df_perfis.drop(columns=["id"]), use_container_width=True, hide_index=True)
        st.caption("Falta configurar SUPABASE_SERVICE_KEY para poderes mudar o Papel de uma conta já criada.")
    else:
        st.caption("✏️ Podes mudar o Papel diretamente na tabela (admin ↔ gestor) — carrega em Guardar no fim.")
        editado_perfis = st.data_editor(
            df_perfis,
            use_container_width=True,
            hide_index=True,
            column_config={
                "id": None,
                "Papel": st.column_config.SelectboxColumn("Papel", options=["gestor", "admin"], required=True),
            },
            disabled=["Email", "Nome"],
            key="editor_perfis",
        )
        if st.button("💾 Guardar papéis"):
            admin_client_papeis = get_admin_client()
            alterados = 0
            for _, r in editado_perfis.iterrows():
                original = df_perfis.loc[df_perfis["id"] == r["id"], "Papel"].values[0]
                if r["Papel"] != original:
                    admin_client_papeis.table("perfis").update({"role": r["Papel"]}).eq("id", r["id"]).execute()
                    alterados += 1
            if alterados:
                st.success(f"{alterados} conta(s) atualizada(s).")
                st.rerun()
            else:
                st.info("Nenhuma alteração para guardar.")
else:
    st.info("Ainda não há contas registadas além da tua.")

st.divider()
st.markdown("### Criar nova conta de gestor")

if not SUPABASE_SERVICE_KEY:
    st.warning(
        "Falta configurar SUPABASE_SERVICE_KEY em Settings → Secrets para poderes criar contas "
        "diretamente pela app. Ver GUIA_SUPABASE.md, secção 'Adicionar gestores'."
    )
else:
    with st.form("form_novo_gestor"):
        novo_nome = st.text_input("Nome do gestor")
        novo_email = st.text_input("Email de login")
        nova_pass = st.text_input("Password inicial (o gestor pode alterá-la depois)", type="password")
        novo_role = st.selectbox("Papel", ["gestor", "admin"])
        submitted = st.form_submit_button("Criar conta")
        if submitted:
            if not novo_email or not nova_pass:
                st.error("Preenche pelo menos o email e a password.")
            else:
                try:
                    admin_client = get_admin_client()
                    criado = admin_client.auth.admin.create_user(
                        {
                            "email": novo_email,
                            "password": nova_pass,
                            "email_confirm": True,
                            "user_metadata": {"nome": novo_nome},
                        }
                    )
                    admin_client.table("perfis").upsert(
                        {"id": criado.user.id, "email": novo_email, "nome": novo_nome, "role": novo_role}
                    ).execute()
                    st.success(f"Conta criada para {novo_email}. Já pode entrar na app com esta password.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Não foi possível criar a conta: {e}")
