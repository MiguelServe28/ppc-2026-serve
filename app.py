"""
Gestão Fiscal SERVE — ponto de entrada da plataforma.

Isto só trata do login e da navegação entre páginas. Cada imposto/módulo vive
na sua própria página dentro de paginas/, e só carrega quando é selecionado no
menu lateral — nada de PPC aparece no Dashboard, por exemplo.

Correr com:  streamlit run app.py

Configuração necessária em .streamlit/secrets.toml (local) ou em
Settings → Secrets (Streamlit Community Cloud):

    SUPABASE_URL = "https://XXXXXXXX.supabase.co"
    SUPABASE_ANON_KEY = "a chave 'anon public' / 'publishable' do projeto"
    SUPABASE_SERVICE_KEY = "a chave 'service_role' / 'secret' do projeto"   # opcional, só o admin precisa
    SMTP_ENC_KEY = "chave Fernet para encriptar as passwords de email"      # ver GUIA, atualização v5

Ver GUIA_SUPABASE.md para o processo completo passo a passo.
"""

import os

import streamlit as st

from common import init_state, requer_login, sou_admin

st.set_page_config(page_title="Gestão Fiscal SERVE", layout="wide", page_icon="📁")

# Logotipo no topo do menu lateral — basta pôr um ficheiro "logo.png" na raiz
# do projeto (junto ao app.py) e ele aparece automaticamente.
if os.path.exists("logo.png"):
    st.logo("logo.png")

# Login obrigatório antes de mostrar qualquer página (mostra o ecrã de login e para, se necessário).
requer_login()

# Estado partilhado (clientes, dados de PPC, parâmetros, templates, log) — carregado uma vez por sessão.
init_state()

paginas = {
    "Plataforma": [
        st.Page("paginas/dashboard.py", title="Dashboard", icon="📊", default=True),
        st.Page("paginas/clientes.py", title="Clientes", icon="📋"),
        st.Page("paginas/perfil.py", title="O Meu Perfil", icon="👤"),
    ],
    "Impostos": [
        st.Page("paginas/ppc.py", title="PPC", icon="💶"),
        st.Page("paginas/iva.py", title="IVA", icon="🧾"),
        st.Page("paginas/irs.py", title="IRS", icon="🗂️"),
        st.Page("paginas/imi.py", title="IMI", icon="🏠"),
        st.Page("paginas/ss.py", title="Segurança Social", icon="🏛️"),
        st.Page("paginas/informacoes.py", title="Informações", icon="ℹ️"),
    ],
}

if sou_admin():
    paginas["Plataforma"].append(st.Page("paginas/gestores.py", title="Gestores", icon="👥"))
    paginas["Plataforma"].append(st.Page("paginas/configuracoes.py", title="Configurações", icon="⚙️"))

pagina_atual = st.navigation(paginas)
pagina_atual.run()
