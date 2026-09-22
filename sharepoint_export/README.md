# SharePoint -> "Base Nova.xlsx" (automacao Windows)

Automatiza, de ponta a ponta:

1. Abre o **Microsoft Edge** (seu perfil do dia a dia, janela **maximizada**) via Playwright.
2. Entra no SharePoint e clica em **Export to Excel**.
3. Intercepta o download (que sai com nome aleatorio), renomeia para **`query.iqy`**
   e coloca na **Area de Trabalho**.
4. Le a URL de dentro do `query.iqy`, monta a consulta web no **Excel** e aguarda
   a carga da base (comeca na hora; sem espera fixa).
5. Aplica **AutoFit Column Width** e depois **AutoFit Row Height** na area de dados.
6. Salva como **`Base Nova.xlsx`** na Area de Trabalho e **fecha o Excel**.

A Area de Trabalho e a pasta Downloads sao descobertas pelo Windows (Known Folders),
entao funciona mesmo com elas redirecionadas para o **OneDrive**.

## Instalacao (uma unica vez)

Com Python 3.10+ instalado, de duplo clique em:

```
Instalar (rodar 1 vez).bat
```

Ou, manualmente:

```bat
python -m pip install -r requirements.txt
python -m playwright install msedge
```

> O Excel e o Edge precisam estar instalados na maquina (padrao no Windows corporativo).

## Uso no dia a dia

Duplo clique em:

```
Rodar Automacao.bat
```

Ao final, a `Base Nova.xlsx` fica na Area de Trabalho e o Excel e fechado automaticamente.

## Configuracao

Tudo o que muda de ambiente esta em `config.py`:

| Item | Variavel |
|---|---|
| URL do SharePoint | `SHAREPOINT_URL` |
| Rotulo do botao | `EXPORT_MENU_ITEM` |
| Perfil do Edge | `EDGE_USER_DATA_DIR`, `EDGE_PROFILE_DIRECTORY` |
| Usar copia do perfil | `USE_PROFILE_COPY` |
| Teto da carga da base | `DATA_LOAD_TIMEOUT_SECONDS` (padrao 600s) |
| Pastas (auto-detectadas) | `DESKTOP_DIR`, `DOWNLOADS_DIR` (None = automatico) |
| Excel visivel durante o processo | `EXCEL_VISIBLE` (padrao False, mais rapido) |
| Nome final do arquivo | `OUTPUT_XLSX_NAME` |
| Manter Excel aberto ao final | `KEEP_EXCEL_OPEN` (padrao False) |

### Sobre o perfil do Edge

O Windows nao deixa dois processos usarem o mesmo perfil do Edge ao mesmo tempo.
Por isso o padrao e `USE_PROFILE_COPY = True`: o script trabalha sobre uma **copia**
do seu perfil (mantendo cookies e SSO da empresa), sem exigir que voce feche o Edge.
A copia e criada em `%LOCALAPPDATA%\SharePointExportAutomation\EdgeProfile` e
reaproveitada nas proximas execucoes.

Se preferir usar o perfil original, coloque `USE_PROFILE_COPY = False` — nesse caso
o Edge precisa estar **totalmente fechado** antes de rodar.

## Primeira execucao

Rode uma vez com calma e acompanhe a janela:

- Se o SharePoint pedir login, faca o login nessa janela; a sessao fica salva na copia
  do perfil e nas proximas vezes o fluxo roda sozinho.
- Se o Excel exibir a barra amarela de seguranca ao abrir o `query.iqy`, clique em
  **Habilitar conteudo** uma vez e marque o site do SharePoint como local confiavel
  (Arquivo > Opcoes > Central de Confiabilidade > Locais Confiaveis). Depois disso o
  script nao para mais nesse ponto.

## Problemas comuns

| Sintoma | Causa provavel | Solucao |
|---|---|---|
| "Nao encontrei o item 'Export to Excel'" | Pagina abriu deslogada ou o rotulo mudou | Rode com a janela visivel, faca login; ajuste `EXPORT_MENU_ITEM` |
| "O download nao foi concluido" | Exportacao demorou mais que o limite | Aumente `DOWNLOAD_WAIT_SECONDS` |
| "A base nao carregou" | Excel bloqueou a conexao externa | Abra o `query.iqy` manualmente uma vez e habilite o conteudo |
| Arquivo nao aparece no Desktop | Desktop redirecionado | Ja tratado; se preciso, fixe `DESKTOP_DIR` em `config.py` |
| Edge nao abre | Perfil travado | Mantenha `USE_PROFILE_COPY = True` |
