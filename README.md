# Gestão de Pagamentos por Conta 2026 — SERVE

App Streamlit que automatiza o processo completo dos Pagamentos por Conta:
1. Importar/gerir clientes (bulk CSV/Excel ou edição manual)
2. Calcular automaticamente o PPC (fórmula do art. 105.º/106.º CIRC, replicada do simulador da OCC)
3. Associar as guias em PDF a cada cliente
4. Gerar e enviar os emails (com a guia em anexo) diretamente por SMTP
5. Exportar a folha de controlo em Excel a qualquer momento

## Instalação

```bash
pip install -r requirements.txt
```

## Arrancar a app

```bash
streamlit run app.py
```

Abre automaticamente no browser em `http://localhost:8501`.

## Fluxo de utilização

### 1. Clientes
- Importa os clientes por CSV/Excel (usa o template disponível na app: colunas `NIF, Nome, Email, Volume_2025, Coleta_2025, Retencoes_2025`) ou edita diretamente na tabela.
- `Volume_2025` = campo 411 (Q11 Mod22) · `Coleta_2025` = campo 351 (Q10 Mod22) · `Retencoes_2025` = campo 359 (Q10 Mod22).

### 2. Cálculo PPC
- Vês logo o total a cobrar, nº de dispensados e o detalhe por cliente.
- **Recomendado:** valida a fórmula contra o simulador da OCC em 5-10 casos reais antes de confiares 100% nela (foi testada com o teu exemplo — Coleta 2.000€, Retenções 200€, Volume 10.000€ → Total 1.440€, 480/480/480 — e bateu certo).

### 3. Guias
- Carrega os PDFs das guias emitidas na AT. Se o nome do ficheiro tiver o NIF (9 dígitos), a associação ao cliente é automática.
- Marca em lote as guias como "Emitidas".

### 4. Emails
- Escolhe qual pagamento (1º/2º/3º) — cada um tem um template de email diferente, editável na app.
- Pré-visualiza o email antes de enviar.
- Configura o SMTP (a password nunca fica guardada, só é usada durante o envio):
  - **Gmail**: `smtp.gmail.com`, porta 587, TLS ligado — precisa de uma [App Password](https://myaccount.google.com/apppasswords), não a password normal.
  - **Outlook/Office365**: `smtp.office365.com`, porta 587, TLS ligado.
  - Outro fornecedor: consulta as definições SMTP do teu domínio de email.
- Seleciona os clientes (por omissão, só aparecem os que ainda não têm o email desse pagamento marcado como enviado) e envia. A guia correspondente é anexada automaticamente se tiver sido carregada na aba anterior.
- Fica um log de envios (sucesso/erro) por cliente.

### 5. Exportar
- A qualquer momento, descarrega o Excel de controlo completo (mesmo formato entregue anteriormente, com fórmulas calculadas e os estados de guia/email).

## Notas importantes

- **Os dados só existem durante a sessão do browser** (não há base de dados). Antes de fechar, exporta o Excel de controlo para não perderes o trabalho. Se quiseres persistência entre sessões, dá para adicionar depois (SQLite ou ligação ao TOConline).
- **Validação da fórmula**: a fórmula está correta para o exemplo testado, mas a lei fiscal tem exceções (ex: grupos sujeitos a RETGS, alterações legislativas). Confirma sempre uma amostra no simulador oficial da OCC antes de confiar cegamente.
- **Guias em PDF**: a app não interage diretamente com o Portal das Finanças (não existe API pública para isso) — o carregamento das guias continua a ser manual, mas a associação ao cliente e o anexo ao email ficam automáticos.
