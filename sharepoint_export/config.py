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
BROWSER_CHANNEL = "msedge"

EDGE_USER_DATA_DIR = os.path.join(
    os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Edge", "User Data"
)
EDGE_PROFILE_DIRECTORY = "Default"

# Trabalha sobre uma COPIA do perfil (nao exige fechar o Edge).
USE_PROFILE_COPY = True

# A copia e criada uma unica vez e reaproveitada. True refaz a cada execucao.
PROFILE_RESYNC_EACH_RUN = False

PROFILE_COPY_DIR = os.path.join(
    os.environ.get("LOCALAPPDATA", os.path.expanduser("~")),
    "SharePointExportAutomation",
    "EdgeProfile",
)

# --------------------------------------------------------------------------- #
# Pastas
# --------------------------------------------------------------------------- #
# Deixe None para o script descobrir sozinho as pastas reais do Windows.
# Isso cobre o caso de Desktop/Downloads redirecionados para o OneDrive.
# Se quiser fixar, coloque o caminho completo entre aspas, por exemplo:
#   DESKTOP_DIR = r"C:\Users\enguicu\OneDrive - Ericsson\Desktop"
DESKTOP_DIR = None
DOWNLOADS_DIR = None

IQY_FILENAME = "query.iqy"
OUTPUT_XLSX_NAME = "Base Nova.xlsx"

# Extensoes aceitas como resultado do "Export to Excel".
ACCEPTED_DOWNLOAD_SUFFIXES = (".iqy", ".xlsx", ".xls", ".csv")

# --------------------------------------------------------------------------- #
# Tempos (segundos)
# --------------------------------------------------------------------------- #
PAGE_LOAD_TIMEOUT_MS = 90_000     # teto para a pagina do SharePoint responder
EXPORT_SEARCH_SECONDS = 60        # teto para achar o botao (sai assim que aparece)
DOWNLOAD_WAIT_SECONDS = 120       # teto para o arquivo chegar (sai assim que chega)
DATA_LOAD_TIMEOUT_SECONDS = 600   # teto para a carga da base (sai assim que carrega)

# --------------------------------------------------------------------------- #
# Comportamento
# --------------------------------------------------------------------------- #
KEEP_BROWSER_OPEN = False   # fecha o Edge ao terminar
EXCEL_VISIBLE = False       # Excel em segundo plano: bem mais rapido
KEEP_EXCEL_OPEN = False     # fecha o Excel depois de salvar a "Base Nova.xlsx"
DELETE_IQY_AFTER = False    # True apaga o "query.iqy" ao final
