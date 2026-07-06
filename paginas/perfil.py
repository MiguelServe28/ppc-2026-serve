"""
O Meu Perfil — cada utilizador (admin ou gestor) pode ver os seus dados,
alterar o nome e mudar a própria password, sem precisar do administrador.
"""

import streamlit as st

from common import get_client, sou_admin

st.title("👤 O Meu Perfil")

perfil = st.session_state.perfil
papel = "Administrador" if perfil.get("role") == "admin" else "Gestor"
st.caption(f"Sessão iniciada como **{perfil.get('email')}** · {papel}")

st.divider()
st.markdown("### Nome")
novo_nome = st.text_input("Nome apresentado na plataforma", value=perfil.get("nome") or "")
if st.button("💾 Guardar nome"):
    try:
        get_client().table("perfis").update({"nome": novo_nome.strip()}).eq("id", perfil["id"]).execute()
        st.session_state.perfil["nome"] = novo_nome.strip()
        st.success("Nome atualizado.")
        st.rerun()
    except Exception as e:
        st.error(f"Não foi possível atualizar o nome: {e}")

st.divider()
st.markdown("### Alterar Password")
st.caption("A nova password passa a ser usada no próximo login. Mínimo de 6 caracteres.")
with st.form("form_mudar_password", clear_on_submit=True):
    nova_pass = st.text_input("Nova password", type="password")
    confirmar = st.text_input("Confirmar nova password", type="password")
    if st.form_submit_button("🔑 Alterar password", type="primary"):
        if len(nova_pass) < 6:
            st.error("A password deve ter pelo menos 6 caracteres.")
        elif nova_pass != confirmar:
            st.error("As passwords não coincidem.")
        else:
            try:
                get_client().auth.update_user({"password": nova_pass})
                st.success("Password alterada com sucesso — usa-a no próximo login.")
            except Exception as e:
                st.error(f"Não foi possível alterar a password: {e}")

if sou_admin():
    st.divider()
    st.caption("💡 Como administrador, crias e geres as contas dos gestores na página 'Gestores'. Cada gestor pode mudar a própria password nesta página.")
