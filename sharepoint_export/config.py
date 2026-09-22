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

# O Playwright guarda o download com nome aleatorio e extensao ".tmp" ate
# alguem consumi-lo. Com True, o script tambem adota esses arquivos: e o
# plano B que garante a captura mesmo se o evento de download falhar.
ACCEPT_TMP_DOWNLOADS = True

# --------------------------------------------------------------------------- #
# Tempos do navegador (segundos)
# --------------------------------------------------------------------------- #
# Esta etapa e a mais fragil (SharePoint lento, Edge travando), entao trabalha
# com folga. Todos os valores sao TETOS: o script segue assim que der certo.
PAGE_LOAD_TIMEOUT_MS = 180_000     # teto para a pagina do SharePoint responder
PAGE_SETTLE_SECONDS = 8            # respiro apos carregar, antes de mexer na pagina
EXPORT_SEARCH_SECONDS = 150        # teto para achar o botao "Export to Excel"
DOWNLOAD_WAIT_SECONDS = 240        # teto para o arquivo chegar apos o clique
POST_CLICK_WAIT_SECONDS = 6        # respiro logo apos o clique, antes de cobrar
POPUP_GRACE_SECONDS = 30           # tempo que um popup pode viver antes de ser fechado
RETRY_CLICK_AFTER_SECONDS = 75     # sem download nesse tempo? clica de novo
EXPORT_CLICK_ATTEMPTS = 3          # quantas vezes re-clicar no "Export to Excel"
BROWSER_ATTEMPTS = 3               # se o Edge travar/crashar, reabre do zero
BROWSER_RETRY_PAUSE_SECONDS = 6    # pausa entre uma tentativa e outra
SLOW_MO_MS = 120                   # freia um pouco os cliques (mais estavel)

# --------------------------------------------------------------------------- #
# Tempos do Excel (segundos)
# --------------------------------------------------------------------------- #
DATA_LOAD_TIMEOUT_SECONDS = 600    # teto para a carga da base (sai assim que carrega)
EXCEL_OPEN_TIMEOUT_SECONDS = 120   # teto para o Excel abrir/preparar a consulta

# Monta a consulta web a partir da URL de dentro do .iqy, em vez de pedir ao
# Excel para interpretar o arquivo. E o caminho mais confiavel: a carga comeca
# imediatamente e de forma sincrona. Com False, usa Workbooks.Open(.iqy).
BUILD_QUERY_FROM_IQY = True

# --------------------------------------------------------------------------- #
# Comportamento
# --------------------------------------------------------------------------- #
KEEP_BROWSER_OPEN = False   # fecha o Edge ao terminar

# Excel visivel. Mantenha True: com a janela oculta, qualquer pedido de
# autenticacao do SharePoint fica invisivel e a carga trava sem aviso.
# O redesenho continua desligado, entao a perda de velocidade e pequena.
EXCEL_VISIBLE = True
KEEP_EXCEL_OPEN = False     # fecha o Excel depois de salvar a "Base Nova.xlsx"
DELETE_IQY_AFTER = False    # True apaga o "query.iqy" ao final
