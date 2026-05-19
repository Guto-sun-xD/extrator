# Extrator PDF — Documentos Financeiros Brasileiros

Digitaliza automaticamente boletos, contas (luz, água, gás, internet, telefone),
notas fiscais e comprovantes em PDF usando visão computacional do **Gemini 2.5 Flash**,
e organiza tudo em uma planilha Excel formatada.

Disponível em dois modos: **interface web (Streamlit)** e **linha de comando**.

---

## Interface Web (Streamlit)

### Rodando localmente

```bash
# 1. Clone ou baixe o projeto
cd extrator-pdf-web

# 2. Crie e ative o ambiente virtual
python -m venv .venv
.venv\Scripts\activate          # Windows
# source .venv/bin/activate     # macOS/Linux

# 3. Instale as dependências
pip install -r requirements.txt

# 4. Configure a chave da API
copy .streamlit\secrets.toml.example .streamlit\secrets.toml
# Edite .streamlit/secrets.toml e preencha GEMINI_API_KEY

# 5. Inicie o app
streamlit run app.py
```

O app abre automaticamente em `http://localhost:8501`.

### Deploy no Streamlit Cloud

1. **Suba o código para o GitHub** (certifique-se de que `.streamlit/secrets.toml` está no `.gitignore`)
2. Acesse [share.streamlit.io](https://share.streamlit.io) e faça login com sua conta GitHub
3. Clique em **New app** → selecione o repositório e o branch
4. Em **Main file path**, coloque `app.py`
5. Clique em **Advanced settings → Secrets** e adicione:
   ```toml
   GEMINI_API_KEY = "sua_chave_real_aqui"
   ```
6. Clique em **Deploy** — o Streamlit Cloud instala as dependências e publica o app automaticamente

A cada `git push` para o branch configurado, o app é atualizado automaticamente.

---

## Linha de Comando (CLI)

### Instalação

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt

copy .env.example .env
# Edite .env e preencha GEMINI_API_KEY
```

### Uso

```bash
# 1. Coloque os PDFs em entrada/
# 2. Execute:
python extrator.py
# 3. Planilha gerada em saida/
# 4. PDFs renomeados em processados/
```

---

## Estrutura do projeto

```
extrator-pdf-web/
├── app.py                      ← interface web Streamlit
├── extrator.py                 ← lógica de extração (CLI + base do web)
├── requirements.txt
├── .env                        ← chave para uso via CLI (não versionado)
├── .env.example
├── .streamlit/
│   ├── config.toml             ← tema visual do Streamlit
│   └── secrets.toml            ← chave para uso web (não versionado)
│   └── secrets.toml.example    ← template (versionado)
├── entrada/                    ← PDFs para o CLI
├── processados/                ← PDFs renomeados pelo CLI
│   └── nao_identificados/
└── saida/                      ← planilhas Excel geradas
```

---

## Tipos de documento reconhecidos

| Tipo | Exemplos |
|------|----------|
| `boleto_bancario` | Boletos de cobrança bancária |
| `conta_consumo` | Energia, água, gás, internet, telefone, streaming |
| `nota_fiscal` | NF-e, NFS-e, NFCe |
| `comprovante_pagamento` | PIX, TED, transferência |
| `recibo` | Recibos de pagamento/quitação |
| `outros` | Qualquer outro documento financeiro |

---

## Planilha Excel gerada

O arquivo `extrato_AAAA-MM_HHMMSS.xlsx` contém:

- **Resumo** — estatísticas gerais, gráficos e próximos vencimentos em destaque
- **Todos** — tabela completa com todos os campos extraídos
- **Boletos** — somente boletos bancários
- **Contas** — somente contas de consumo
- **Notas Fiscais** — somente NFs
- **Outros** — comprovantes, recibos e demais

### Formatação

- Cabeçalhos azul escuro (`#1F3864`) com texto branco
- Linhas alternadas branco/azul claro
- Vencimentos em vermelho (vencidos) ou amarelo (próximos 30 dias)
- Valores em `R$ #.##0,00` e datas em `DD/MM/AAAA`

---

## Tratamento de erros

| Situação | Ação |
|----------|------|
| PDF corrompido / ilegível | Registra erro; não bloqueia os demais |
| Resposta do Gemini inválida | Registra erro; continua o lote |
| API key ausente | Exibe instrução clara e para |
| Rate limit (erro 429) | Retry automático com espera progressiva (60s, 120s) |

---

## Variáveis de ambiente / secrets

| Variável | Onde configurar | Descrição |
|----------|-----------------|-----------|
| `GEMINI_API_KEY` | `.env` (CLI) ou `.streamlit/secrets.toml` (web) | Chave da API do Google AI Studio |

Obtenha sua chave em: <https://aistudio.google.com/app/apikey>
