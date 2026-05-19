# Pianetto — Dashboard Contas a Receber (Streamlit)

Dashboard **público** das contas a receber da Pianetto Transportes
(11 filiais — CNPJ raiz 43.976.512), com janela configurável (7 a 90 dias).

> ⚠️ **Atenção LGPD**: este dashboard é público (sem senha). Por padrão, o
> sidebar mantém o toggle **"Mascarar CPF/CNPJ"** ligado e oferece a opção
> **"Anonimizar clientes"** para substituir o nome por um rótulo genérico.
> Se a base contiver clientes pessoa física (CPF), avalie deixar a
> anonimização ligada por padrão ou restringir o acesso.

## Estrutura do app

- **KPIs**:
  - Valor total dos títulos · Já recebido · Saldo a receber
  - Vence hoje · Próximos 7/14/21 dias
- **Por Dia**: barras empilhadas (recebido + saldo) + tabela
- **Por Cliente**: ranking Top N + pizza + tabela com % e acumulado + export Excel
- **Dia × Cliente**: heatmap
- **Por Filial**: barras + pizza
- **Detalhe**: 1 linha por título, com export Excel

## Rodar localmente

```bash
cd streamlit_pianetto_receber
python -m venv .venv
.venv\Scripts\activate     # Windows
pip install -r requirements.txt

# Crie .streamlit/secrets.toml a partir do .example e preencha a senha do banco
copy .streamlit\secrets.toml.example .streamlit\secrets.toml

streamlit run streamlit_app.py
```

## Deploy no Streamlit Community Cloud

### 1) Subir para o GitHub — **tudo de uma vez**

Use o script `deploy.ps1` (PowerShell — Windows). Ele cuida de:

- conferir que `.streamlit/secrets.toml` não vai pro Git
- instalar o `gh` CLI via `winget` se faltar (com confirmação)
- `gh auth login` se ainda não autenticado
- `git init` + `commit` + `gh repo create --private --push`

```powershell
cd streamlit_pianetto_receber
.\deploy.ps1                                  # repo "pianetto-contas-receber", privado
# ou:
.\deploy.ps1 -RepoName "meu-repo"
.\deploy.ps1 -RepoName "meu-repo" -Public     # repositorio publico no GitHub
.\deploy.ps1 -SkipPush                        # so o commit local; voce empurra depois
```

> **IMPORTANTE**: `.streamlit/secrets.toml` está no `.gitignore` — o próprio
> `deploy.ps1` aborta se detectar que esse arquivo entrou no staging. Só o
> `.example` vai pro repo.

#### Subindo manualmente (sem o script)

```bash
cd streamlit_pianetto_receber
git init -b main
git add .
git commit -m "Pianetto - dashboard contas a receber"

# crie o repo no GitHub e:
git remote add origin https://github.com/<seu-usuario>/pianetto-contas-receber.git
git push -u origin main
```

### 2) Criar o app em share.streamlit.io

1. Acesse https://share.streamlit.io e faça login com o GitHub.
2. Clique **New app** → selecione o repositório, branch `main`, arquivo
   principal `streamlit_app.py`.
3. Em **Advanced settings → Secrets**, cole:

```toml
[postgres]
host = "187.77.242.106"
port = 6432
dbname = "jhn"
user = "jhn_ro"
password = "<senha do banco>"
sslmode = "disable"
```

4. Clique **Deploy**. O app fica disponível em
   `https://<seu-usuario>-pianetto-contas-receber.streamlit.app`.
5. Em **Settings → Sharing**, deixe **"Anyone with the link"** marcado para
   manter o acesso público.

### 3) Compartilhar a URL

Qualquer pessoa com a URL acessa o dashboard sem login. Considere:

- Manter os toggles de privacidade ligados.
- Ativar o cache (já vem com TTL=15 min) para reduzir custo de queries.
- Monitorar o uso pelos analytics do Streamlit Cloud.

## Segurança

- Credenciais do banco ficam **só em `st.secrets`** (Streamlit Cloud) ou em
  `.streamlit/secrets.toml` local (ignorado pelo Git).
- O usuário do banco é `jhn_ro` (read-only) — não é possível alterar dados.
- Cache de 15 min reduz carga no banco.
- `enableXsrfProtection = true` em `.streamlit/config.toml`.

## Estrutura de arquivos

```
streamlit_pianetto_receber/
├── streamlit_app.py            # app principal
├── requirements.txt
├── README.md
├── .gitignore                  # secrets.toml fora do git
└── .streamlit/
    ├── config.toml
    └── secrets.toml.example    # template; o real NUNCA vai pro git
```
