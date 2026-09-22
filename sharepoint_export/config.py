# -*- coding: utf-8 -*-
"""
Configuracoes do automatizador.
Edite SOMENTE este arquivo para adaptar a rotina ao seu ambiente.
"""

import os

# --------------------------------------------------------------------------- #
# SharePoint
# --------------------------------------------------------------------------- #
SHAREPOINT_URL = (
    "https://ericsson.sharepoint.com/sites/LegalServicesCULatamSouth/SitePages/Home.aspx"
)

# Texto do item de menu que dispara a exportacao.
EXPORT_MENU_ITEM = "Export to Excel"

# --------------------------------------------------------------------------- #
# Navegador (Microsoft Edge pessoal / do dia a dia)
# --------------------------------------------------------------------------- #
# Canal do Playwright: "msedge" usa o Microsoft Edge instalado na maquina.
BROWSER_CHANNEL = "msedge"

# Perfil do Edge do dia a dia (mantem logins/SSO da empresa).
EDGE_USER_DATA_DIR = os.path.join(
    os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Edge", "User Data"
)
# Nome do perfil dentro do User Data ("Default", "Profile 1", ...).
EDGE_PROFILE_DIRECTORY = "Default"

# O Playwright nao consegue abrir um perfil que ja esteja em uso por outro
# processo do Edge. Com True, o script trabalha sobre uma COPIA do perfil
# (mais seguro e nao exige fechar o Edge). Com False, usa o perfil original
# e exige que o Edge esteja totalmente fechado.
USE_PROFILE_COPY = True

# Refaz a copia do perfil a cada execucao. Por padrao False: a copia e criada
# uma unica vez (na 1a execucao voce faz login nela) e depois e reaproveitada,
# o que evita conflito com o Edge aberto, que trava arquivos como o "Cookies".
PROFILE_RESYNC_EACH_RUN = False

# Onde a copia de trabalho do perfil fica guardada (reaproveitada entre runs).
PROFILE_COPY_DIR = os.path.join(
    os.environ.get("LOCALAPPDATA", os.path.expanduser("~")),
    "SharePointExportAutomation",
    "EdgeProfile",
)

# --------------------------------------------------------------------------- #
# Arquivos
# --------------------------------------------------------------------------- #
DOWNLOADS_DIR = os.path.join(os.path.expanduser("~"), "Downloads")
DESKTOP_DIR = os.path.join(os.path.expanduser("~"), "Desktop")

IQY_FILENAME = "query.iqy"
OUTPUT_XLSX_NAME = "Base Nova.xlsx"

# Extensoes aceitas como resultado do "Export to Excel".
ACCEPTED_DOWNLOAD_SUFFIXES = (".iqy", ".xlsx", ".xls", ".csv")

# --------------------------------------------------------------------------- #
# Tempos (segundos)
# --------------------------------------------------------------------------- #
PAGE_LOAD_TIMEOUT_MS = 120_000       # carregamento da pagina do SharePoint
DOWNLOAD_WAIT_SECONDS = 180          # espera pelo arquivo exportado
DATA_LOAD_WAIT_SECONDS = 180         # os "3 minutos" de carga da base no Excel
DATA_LOAD_EXTRA_GRACE_SECONDS = 120  # folga adicional se a query ainda atualizar

# Mantem o navegador aberto ao final (util para depurar).
KEEP_BROWSER_OPEN = False

# Mantem o Excel visivel/aberto ao final com a "Base Nova.xlsx" pronta.
KEEP_EXCEL_OPEN = True
