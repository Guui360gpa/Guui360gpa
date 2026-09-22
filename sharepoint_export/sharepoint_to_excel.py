# -*- coding: utf-8 -*-
"""
Automatiza o fluxo completo:

  1. Abre o Microsoft Edge (perfil pessoal, maximizado) via Playwright.
  2. Acessa o SharePoint e clica em "Export to Excel".
  3. Intercepta o download, renomeia para "query.iqy" e move para a Area de
     Trabalho REAL (inclusive quando ela esta redirecionada para o OneDrive).
  4. Abre o arquivo no Excel (em segundo plano), libera a conexao externa e
     aguarda a carga da base.
  5. Aplica AutoFit de largura de coluna e de altura de linha.
  6. Salva como "Base Nova.xlsx" na Area de Trabalho e FECHA o Excel.

Uso:
    python sharepoint_to_excel.py

Requisitos: Windows + Microsoft Edge + Microsoft Excel instalados.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
import time
import traceback
from datetime import datetime
from typing import Optional

import config

# --------------------------------------------------------------------------- #
# Log
# --------------------------------------------------------------------------- #

_START = time.time()


def log(message: str) -> None:
    print(f"[{datetime.now():%H:%M:%S} | +{time.time() - _START:6.1f}s] {message}",
          flush=True)


class AutomationError(RuntimeError):
    """Erro de negocio previsto pela automacao (mensagem amigavel)."""


# --------------------------------------------------------------------------- #
# Pastas reais do Windows (Desktop/Downloads podem estar no OneDrive)
# --------------------------------------------------------------------------- #

FOLDERID_DESKTOP = "{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}"
FOLDERID_DOWNLOADS = "{374DE290-123F-4565-9164-39C4925E467B}"


def _known_folder(folder_id: str) -> Optional[str]:
    """Consulta o caminho real da pasta especial via API do Windows."""
    if os.name != "nt":
        return None
    try:
        import ctypes
        from ctypes import wintypes

        class GUID(ctypes.Structure):
            _fields_ = [
                ("Data1", ctypes.c_uint32),
                ("Data2", ctypes.c_uint16),
                ("Data3", ctypes.c_uint16),
                ("Data4", ctypes.c_ubyte * 8),
            ]

        guid = GUID()
        if ctypes.windll.ole32.CLSIDFromString(
            ctypes.create_unicode_buffer(folder_id), ctypes.byref(guid)
        ) != 0:
            return None

        pointer = ctypes.c_wchar_p()
        result = ctypes.windll.shell32.SHGetKnownFolderPath(
            ctypes.byref(guid), 0, wintypes.HANDLE(0), ctypes.byref(pointer)
        )
        if result != 0 or not pointer.value:
            return None
        path = pointer.value
        ctypes.windll.ole32.CoTaskMemFree(pointer)
        return path
    except Exception:
        return None


def _shell_folder_from_registry(value_name: str) -> Optional[str]:
    """Plano B: le a pasta especial no registro do usuario."""
    if os.name != "nt":
        return None
    try:
        import winreg

        key_path = (r"Software\Microsoft\Windows\CurrentVersion"
                    r"\Explorer\User Shell Folders")
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as key:
            raw, _ = winreg.QueryValueEx(key, value_name)
        return os.path.expandvars(raw)
    except Exception:
        return None


def _resolve_folder(configured: Optional[str], folder_id: str,
                    registry_name: str, fallback_name: str, label: str) -> str:
    """
    Descobre a pasta real (Desktop/Downloads), nesta ordem:
    config.py -> API do Windows -> registro -> OneDrive -> %USERPROFILE%.
    """
    if configured:
        return os.path.abspath(os.path.expandvars(configured))

    for candidate in (
        _known_folder(folder_id),
        _shell_folder_from_registry(registry_name),
        os.path.join(os.environ.get("OneDriveCommercial", "") or
                     os.environ.get("OneDrive", "") or "\x00", fallback_name),
        os.path.join(os.path.expanduser("~"), fallback_name),
    ):
        if candidate and "\x00" not in candidate and os.path.isdir(candidate):
            return os.path.abspath(candidate)

    # Ultimo recurso: cria dentro do perfil do usuario.
    fallback = os.path.join(os.path.expanduser("~"), fallback_name)
    os.makedirs(fallback, exist_ok=True)
    log(f"AVISO: nao consegui detectar a pasta '{label}'. Usando: {fallback}")
    return fallback


def get_desktop_dir() -> str:
    return _resolve_folder(getattr(config, "DESKTOP_DIR", None),
                           FOLDERID_DESKTOP, "Desktop", "Desktop", "Desktop")


def get_downloads_dir() -> str:
    return _resolve_folder(getattr(config, "DOWNLOADS_DIR", None),
                           FOLDERID_DOWNLOADS, "{374DE290-123F-4565-9164-39C4925E467B}",
                           "Downloads", "Downloads")


DESKTOP_DIR = ""
DOWNLOADS_DIR = ""


# --------------------------------------------------------------------------- #
# Pre-requisitos
# --------------------------------------------------------------------------- #


def ensure_windows() -> None:
    if os.name != "nt":
        raise AutomationError(
            "Esta automacao foi feita para Windows (usa Excel via COM). "
            f"Sistema detectado: {os.name}."
        )


def edge_is_running() -> bool:
    """Retorna True se houver processo msedge.exe ativo."""
    try:
        output = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq msedge.exe", "/NH"],
            capture_output=True, text=True, timeout=30,
        ).stdout.lower()
    except Exception:
        return False
    return "msedge.exe" in output


PROFILE_SKIP_NAMES = {
    "cache", "code cache", "gpucache", "dawncache", "shadercache",
    "grshadercache", "media cache", "service worker", "crashpad",
    "sessions", "gcm store", "optimization guide model store",
    "component_crx_cache", "extensions_crx_cache", "safe browsing",
}


def _copy_profile_tolerant(src: str, dst: str):
    """
    Copia o perfil do Edge ignorando o que estiver travado por outro processo.

    O Edge mantem 'Cookies', 'Sessions' e afins abertos enquanto roda; um
    shutil.copytree comum aborta tudo com WinError 32. Aqui cada arquivo e
    copiado isoladamente e as falhas sao apenas contadas.
    """
    copied = skipped = 0
    for root, dirs, files in os.walk(src):
        dirs[:] = [d for d in dirs if d.lower() not in PROFILE_SKIP_NAMES]

        relative = os.path.relpath(root, src)
        destination_root = dst if relative == "." else os.path.join(dst, relative)
        try:
            os.makedirs(destination_root, exist_ok=True)
        except OSError:
            skipped += len(files)
            continue

        for name in files:
            if name.lower() in PROFILE_SKIP_NAMES or name.lower().endswith(".log"):
                continue
            try:
                shutil.copy2(os.path.join(root, name),
                             os.path.join(destination_root, name))
                copied += 1
            except (OSError, shutil.Error):
                skipped += 1  # arquivo em uso pelo Edge: segue o jogo
    return copied, skipped


def _clear_profile_locks(destination: str) -> None:
    """Remove travas remanescentes da copia do perfil."""
    for lock in ("SingletonLock", "SingletonCookie", "SingletonSocket"):
        lock_path = os.path.join(destination, lock)
        if os.path.exists(lock_path):
            try:
                os.remove(lock_path)
            except OSError:
                pass


def prepare_profile_dir() -> str:
    """
    Devolve o diretorio de perfil que sera usado pelo Playwright.

    A copia e criada uma unica vez (marcador '.profile_seeded'); dali em diante
    ela tem a propria sessao logada e nao precisa mais ser sincronizada.
    """
    source = config.EDGE_USER_DATA_DIR
    if not source or not os.path.isdir(source):
        raise AutomationError(
            "Perfil do Microsoft Edge nao encontrado em:\n"
            f"  {source}\n"
            "Ajuste EDGE_USER_DATA_DIR em config.py."
        )

    if not config.USE_PROFILE_COPY:
        if edge_is_running():
            raise AutomationError(
                "O Microsoft Edge esta aberto e o script foi configurado para usar "
                "o perfil original.\nFeche TODAS as janelas do Edge e rode de novo, "
                "ou deixe USE_PROFILE_COPY = True em config.py."
            )
        return source

    destination = config.PROFILE_COPY_DIR
    profile = config.EDGE_PROFILE_DIRECTORY
    src_profile = os.path.join(source, profile)
    dst_profile = os.path.join(destination, profile)
    marker = os.path.join(destination, ".profile_seeded")

    if not os.path.isdir(src_profile):
        raise AutomationError(
            f"Perfil '{profile}' nao existe em {source}. "
            "Ajuste EDGE_PROFILE_DIRECTORY em config.py."
        )

    if (os.path.isfile(marker) and os.path.isdir(dst_profile)
            and not getattr(config, "PROFILE_RESYNC_EACH_RUN", False)):
        _clear_profile_locks(destination)
        return destination

    os.makedirs(destination, exist_ok=True)
    for name in ("Local State", "Last Version"):
        src_file = os.path.join(source, name)
        if os.path.isfile(src_file):
            try:
                shutil.copy2(src_file, os.path.join(destination, name))
            except OSError:
                pass

    log(f"Preparando copia do perfil do Edge ('{profile}')... so acontece na 1a vez.")
    copied, skipped = _copy_profile_tolerant(src_profile, dst_profile)
    log(f"Perfil copiado: {copied} arquivos ({skipped} ignorados por estarem em uso).")
    if skipped:
        log("Se o SharePoint pedir login nesta janela, faca o login uma vez: "
            "a sessao fica salva na copia e as proximas execucoes rodam sozinhas.")

    _clear_profile_locks(destination)
    try:
        with open(marker, "w", encoding="utf-8") as handle:
            handle.write(datetime.now().isoformat())
    except OSError:
        pass
    return destination


# --------------------------------------------------------------------------- #
# Etapas 1 a 3 - Playwright / SharePoint
# --------------------------------------------------------------------------- #


def _snapshot_downloads() -> set:
    """Fotografa a pasta Downloads para detectar arquivos novos depois."""
    try:
        return set(os.listdir(DOWNLOADS_DIR))
    except OSError:
        return set()


def _file_is_stable(path: str) -> bool:
    """Confirma que o arquivo parou de crescer (download concluido)."""
    try:
        size = os.path.getsize(path)
        for _ in range(2):
            time.sleep(0.4)
            current = os.path.getsize(path)
            if current != size:
                return False
            size = current
        return True
    except OSError:
        return False


def _looks_like_iqy(path: str) -> bool:
    """Um .iqy e texto e comeca com 'WEB'. Serve para confirmar o .tmp."""
    try:
        with open(path, "rb") as handle:
            head = handle.read(16).lstrip()
        return head[:3].upper() == b"WEB"
    except OSError:
        return False


def _new_finished_download(before: set) -> Optional[str]:
    """
    Devolve o arquivo novo e ja finalizado na pasta Downloads, ou None.

    Inclui os '<guid>.tmp' que o Playwright cria: quando o evento de download
    nao e consumido, o arquivo fica so com esse nome temporario - e ele e a
    base exportada. Consulta unica (sem bloquear): quem chama define o ritmo.
    """
    try:
        current = set(os.listdir(DOWNLOADS_DIR))
    except OSError:
        return None

    accept_tmp = bool(getattr(config, "ACCEPT_TMP_DOWNLOADS", True))
    named, temporary = [], []

    for name in current - before:
        lowered = name.lower()
        if lowered.endswith((".crdownload", ".partial")):
            continue  # ainda baixando
        path = os.path.join(DOWNLOADS_DIR, name)
        if lowered.endswith(config.ACCEPTED_DOWNLOAD_SUFFIXES):
            named.append(path)
        elif accept_tmp and lowered.endswith(".tmp"):
            temporary.append(path)

    # Prefere um arquivo com extensao de verdade; so depois recorre ao .tmp.
    for group, is_temp in ((named, False), (temporary, True)):
        if not group:
            continue
        newest = max(group, key=os.path.getmtime)
        try:
            if os.path.getsize(newest) == 0:
                continue
        except OSError:
            continue
        if not _file_is_stable(newest):
            continue
        if is_temp and not _looks_like_iqy(newest):
            log(f"Arquivo temporario '{os.path.basename(newest)}' nao parece um "
                ".iqy; vou usa-lo assim mesmo.")
        return newest

    return None


def download_export_file() -> str:
    """
    Executa o fluxo no SharePoint e devolve o caminho do 'query.iqy' na Area
    de Trabalho. Se o Edge travar ou fechar sozinho, reabre e tenta de novo.
    """
    attempts = max(1, int(getattr(config, "BROWSER_ATTEMPTS", 3)))
    last_error: Optional[Exception] = None

    for attempt in range(1, attempts + 1):
        if attempt > 1:
            pause = getattr(config, "BROWSER_RETRY_PAUSE_SECONDS", 6)
            log(f"Reiniciando o navegador (tentativa {attempt} de {attempts}) "
                f"apos {pause}s...")
            time.sleep(pause)
        try:
            return _download_once()
        except AutomationError:
            raise  # erro de negocio: repetir nao ajuda
        except Exception as exc:
            # Tipicamente o Edge caiu no meio do caminho.
            last_error = exc
            log(f"A sessao do navegador falhou: {exc}")

    raise AutomationError(
        "O navegador nao completou a exportacao apos "
        f"{attempts} tentativas. Ultimo erro: {last_error}"
    )


def _download_once() -> str:
    """Uma tentativa completa: abre o Edge, clica, espera o arquivo."""
    from playwright.sync_api import sync_playwright

    user_data_dir = prepare_profile_dir()
    os.makedirs(DESKTOP_DIR, exist_ok=True)

    target_path = os.path.join(DESKTOP_DIR, config.IQY_FILENAME)
    captured: dict = {}
    popups_seen: dict = {}          # pagina -> instante em que apareceu
    before = _snapshot_downloads()

    with sync_playwright() as playwright:
        log("Abrindo o Microsoft Edge (maximizado)...")
        context = playwright.chromium.launch_persistent_context(
            user_data_dir=user_data_dir,
            channel=config.BROWSER_CHANNEL,
            headless=False,
            accept_downloads=True,
            downloads_path=DOWNLOADS_DIR,
            no_viewport=True,  # necessario para a janela ficar realmente maximizada
            slow_mo=int(getattr(config, "SLOW_MO_MS", 0)),
            args=[
                "--start-maximized",
                f"--profile-directory={config.EDGE_PROFILE_DIRECTORY}",
                "--no-first-run",
                "--no-default-browser-check",
                "--disable-features=msEdgeIdentityRefresh",
                # Evita que o Edge "hiberne" a aba e derrube o download.
                "--disable-background-timer-throttling",
                "--disable-renderer-backgrounding",
                "--disable-backgrounding-occluded-windows",
            ],
        )

        def handle_download(download) -> None:
            """Intercepta o download em qualquer aba/popup do contexto."""
            if "path" in captured:
                return
            try:
                log(f"Download interceptado: {download.suggested_filename}")
                download.save_as(target_path)
                captured["path"] = target_path
                log(f"Salvo como: {target_path}")
            except Exception as exc:
                log(f"Nao foi possivel salvar pelo Playwright ({exc}). "
                    "Vou pegar o arquivo na pasta Downloads.")

        def handle_page(new_page) -> None:
            """
            Registra o popup que o SharePoint abre e escuta o download NELE.

            O evento 'download' e emitido pela pagina, nao pelo contexto - por
            isso ele precisa ser ligado em cada aba/popup que aparecer.
            """
            popups_seen[new_page] = time.time()
            log("O SharePoint abriu uma nova janela (e dela que vem o download).")
            try:
                new_page.on("download", handle_download)
            except Exception:
                pass

        context.on("page", handle_page)

        try:
            page = context.pages[0] if context.pages else context.new_page()
            page.set_default_timeout(config.PAGE_LOAD_TIMEOUT_MS)
            page.on("download", handle_download)  # a aba principal tambem baixa

            log(f"Acessando: {config.SHAREPOINT_URL}")
            # 'domcontentloaded' basta para seguir; o SharePoint mantem conexoes
            # abertas e nunca atinge 'networkidle'.
            page.goto(config.SHAREPOINT_URL, wait_until="domcontentloaded",
                      timeout=config.PAGE_LOAD_TIMEOUT_MS)

            # Deixa a pagina assentar: o SharePoint monta a barra de comandos
            # depois do DOM, e clicar cedo demais nao dispara a exportacao.
            try:
                page.wait_for_load_state("load", timeout=30_000)
            except Exception:
                pass
            settle = float(getattr(config, "PAGE_SETTLE_SECONDS", 8))
            if settle > 0:
                log(f"Aguardando a pagina assentar ({settle:.0f}s)...")
                page.wait_for_timeout(int(settle * 1000))

            max_clicks = max(1, int(getattr(config, "EXPORT_CLICK_ATTEMPTS", 3)))
            retry_after = float(getattr(config, "RETRY_CLICK_AFTER_SECONDS", 75))
            post_click = float(getattr(config, "POST_CLICK_WAIT_SECONDS", 6))
            overall_deadline = time.time() + config.DOWNLOAD_WAIT_SECONDS

            for click_number in range(1, max_clicks + 1):
                label = (f" (tentativa {click_number} de {max_clicks})"
                         if click_number > 1 else "")
                log(f"Procurando '{config.EXPORT_MENU_ITEM}'{label}...")
                if not _click_export(page):
                    raise AutomationError(
                        f"Nao encontrei o item '{config.EXPORT_MENU_ITEM}' na pagina. "
                        "Confirme se a pagina carregou logada e se o nome do botao mudou."
                    )

                # Respiro obrigatorio: o download so comeca depois que o
                # SharePoint monta a janela auxiliar e responde a requisicao.
                if post_click > 0:
                    log(f"Dando {post_click:.0f}s para a exportacao comecar...")
                    time.sleep(post_click)

                log("Aguardando o arquivo exportado...")
                window_deadline = min(time.time() + retry_after, overall_deadline)
                if click_number == max_clicks:
                    window_deadline = overall_deadline

                if _wait_for_file(context, captured, before, target_path,
                                  window_deadline, popups_seen):
                    break

                if time.time() >= overall_deadline:
                    break
                log("Nada baixou ainda. Vou clicar em 'Export to Excel' novamente.")

        finally:
            if not config.KEEP_BROWSER_OPEN:
                try:
                    context.close()
                except Exception:
                    pass

    if "path" not in captured:
        raise AutomationError(
            "O download nao foi concluido dentro do tempo limite "
            f"({config.DOWNLOAD_WAIT_SECONDS}s). O SharePoint pode estar lento; "
            "aumente DOWNLOAD_WAIT_SECONDS em config.py e tente de novo."
        )
    return captured["path"]


def _wait_for_file(context, captured: dict, before: set, target_path: str,
                   deadline: float, popups_seen: dict) -> bool:
    """
    Espera o arquivo ate 'deadline'. Devolve True se conseguiu.

    Vigia as duas frentes: o evento de download do Playwright e a pasta
    Downloads (o Edge as vezes baixa sem disparar o evento).

    IMPORTANTE: a janela auxiliar que o SharePoint abre costuma ficar em
    'about:blank' justamente enquanto prepara o arquivo. Fecha-la cedo demais
    CANCELA o download - por isso so encerramos popups bem antigos.
    """
    grace = float(getattr(config, "POPUP_GRACE_SECONDS", 30))

    while time.time() < deadline:
        if "path" in captured:
            return True

        found = _new_finished_download(before)
        if found:
            log(f"Arquivo novo em Downloads: {os.path.basename(found)}")
            captured["path"] = _move_to_desktop(found, target_path)
            return True

        # Faxina conservadora: so fecha popup vazio que ja passou do prazo.
        for extra in list(context.pages)[1:]:
            try:
                if extra.is_closed():
                    continue
                appeared = popups_seen.setdefault(extra, time.time())
                if (extra.url in ("about:blank", "")
                        and time.time() - appeared > grace):
                    extra.close()
            except Exception:
                pass

        time.sleep(0.25)

    return "path" in captured


def _click_export(page) -> bool:
    """
    Clica em 'Export to Excel' assim que ele estiver realmente pronto.

    Varre todas as formas em que o SharePoint pode renderizar o comando, em
    ciclos curtos, ate o teto de EXPORT_SEARCH_SECONDS.
    """
    name = config.EXPORT_MENU_ITEM
    started = time.time()
    deadline = started + config.EXPORT_SEARCH_SECONDS
    openers_tried = False

    def candidates():
        yield page.get_by_role("menuitem", name=name), "menuitem"
        yield page.get_by_role("button", name=name), "button"
        yield page.get_by_role("link", name=name), "link"
        yield page.get_by_text(name, exact=False), "texto"
        for frame in page.frames:
            if frame is page.main_frame:
                continue
            yield frame.get_by_role("menuitem", name=name), "menuitem (iframe)"
            yield frame.get_by_role("button", name=name), "button (iframe)"
            yield frame.get_by_text(name, exact=False), "texto (iframe)"

    def try_click(locator, description: str) -> bool:
        try:
            if locator.count() == 0:
                return False
            target = locator.first
            # Espera o elemento ficar visivel e habilitado antes de clicar:
            # clicar num item que ainda esta montando nao dispara nada.
            target.wait_for(state="visible", timeout=5_000)
            if not target.is_enabled():
                return False
            target.scroll_into_view_if_needed(timeout=5_000)
            target.click(timeout=15_000)
            log(f"Clique efetuado em: {description}")
            page.wait_for_timeout(1_500)  # deixa o SharePoint reagir ao clique
            return True
        except Exception:
            return False

    while time.time() < deadline:
        for locator, description in candidates():
            if try_click(locator, description):
                return True

        # Depois de alguns ciclos sem achar, abre os menus que podem escondê-lo.
        if not openers_tried and time.time() - started > 8:
            openers_tried = True
            for opener in ("Export", "More options", "More", "..."):
                try:
                    button = page.get_by_role("button", name=opener, exact=False)
                    if button.count() and button.first.is_visible():
                        button.first.click(timeout=8_000)
                        page.wait_for_timeout(1_000)
                except Exception:
                    continue

        page.wait_for_timeout(400)

    return False


def _move_to_desktop(source: str, target_path: str) -> str:
    """
    Leva o arquivo baixado para a Area de Trabalho com o nome definitivo.

    Se o navegador ainda segurar o arquivo (WinError 32), copia em vez de
    mover e apaga a origem depois, quando der.
    """
    os.makedirs(os.path.dirname(target_path), exist_ok=True)
    if os.path.exists(target_path):
        try:
            os.remove(target_path)
        except OSError:
            pass

    last_error: Optional[Exception] = None
    for attempt in range(4):
        try:
            shutil.move(source, target_path)
            log(f"Arquivo movido para: {target_path}")
            return target_path
        except (OSError, shutil.Error) as exc:
            last_error = exc
            time.sleep(0.5 * (attempt + 1))

    try:
        shutil.copy2(source, target_path)
        log(f"Arquivo copiado para: {target_path} (a origem estava em uso).")
        try:
            os.remove(source)
        except OSError:
            pass
        return target_path
    except (OSError, shutil.Error) as exc:
        raise AutomationError(
            f"Nao consegui levar o arquivo para a Area de Trabalho: {exc}\n"
            f"Origem: {source}\nUltimo erro ao mover: {last_error}"
        )


# --------------------------------------------------------------------------- #
# Etapas 4 a 6 - Excel
# --------------------------------------------------------------------------- #

XL_OPENXML_WORKBOOK = 51          # xlOpenXMLWorkbook (.xlsx)
XL_CALCULATION_MANUAL = -4135
XL_CALCULATION_AUTOMATIC = -4105
MSO_AUTOMATION_SECURITY_LOW = 1   # msoAutomationSecurityLow

# Constantes de consulta web (QueryTable)
XL_ENTIRE_PAGE = 1
XL_ALL_TABLES = 2
XL_WEB_FORMATTING_ALL = 1
XL_WEB_FORMATTING_RTF = 2
XL_WEB_FORMATTING_NONE = 3


def _com_retry(action, description: str, attempts: int = 6, pause: float = 2.0):
    """
    Executa uma chamada COM tolerando o Excel 'ocupado'.

    Quando o Excel esta processando algo, ele rejeita chamadas com
    'Call was rejected by callee' (RPC_E_CALL_REJECTED). Insistir resolve.
    """
    last_error: Optional[Exception] = None
    for attempt in range(1, attempts + 1):
        try:
            return action()
        except Exception as exc:
            last_error = exc
            message = str(exc)
            busy = ("rejected by callee" in message
                    or "0x80010001" in message
                    or "Call was rejected" in message
                    or "busy" in message.lower())
            if not busy or attempt == attempts:
                raise
            log(f"  Excel ocupado em '{description}'; tentando de novo "
                f"({attempt}/{attempts})...")
            time.sleep(pause)
    raise last_error  # pragma: no cover


def _parse_iqy(path: str) -> dict:
    """
    Le o .iqy exportado pelo SharePoint.

    Formato:
        WEB
        1
        <url da consulta>
        <linha em branco>
        Selection=EntirePage
        Formatting=None
        ...
    """
    try:
        with open(path, "r", encoding="utf-8-sig", errors="replace") as handle:
            lines = [line.strip() for line in handle.read().splitlines()]
    except OSError as exc:
        raise AutomationError(f"Nao consegui ler '{path}': {exc}")

    url = ""
    options: dict = {}
    for line in lines[2:]:           # as duas primeiras sao 'WEB' e '1'
        if not line:
            continue
        if not url and ("://" in line or line.lower().startswith("//")):
            url = line
            continue
        if "=" in line:
            key, _, value = line.partition("=")
            options[key.strip().lower()] = value.strip()

    if not url:
        raise AutomationError(
            f"Nao encontrei a URL da consulta dentro de '{os.path.basename(path)}'.\n"
            "O arquivo baixado pode nao ser um .iqy valido."
        )
    return {"url": url, "options": options}


def process_in_excel(iqy_path: str) -> str:
    """Carrega a base no Excel, aplica AutoFit, salva como .xlsx e fecha."""
    import pythoncom
    import win32com.client as win32

    output_path = os.path.join(DESKTOP_DIR, config.OUTPUT_XLSX_NAME)

    pythoncom.CoInitialize()
    excel = None
    workbook = None
    try:
        log("Abrindo o Microsoft Excel...")
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = bool(getattr(config, "EXCEL_VISIBLE", True))
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        excel.ScreenUpdating = False   # sem redesenho: bem mais rapido
        try:
            excel.EnableEvents = False
        except Exception:
            pass
        try:
            excel.AutomationSecurity = MSO_AUTOMATION_SECURITY_LOW
        except Exception:
            pass

        workbook = _open_base(excel, iqy_path)

        try:
            excel.Calculation = XL_CALCULATION_MANUAL
        except Exception:
            pass

        _autofit_used_range(workbook)

        try:
            excel.Calculation = XL_CALCULATION_AUTOMATIC
        except Exception:
            pass

        if os.path.exists(output_path):
            log("Removendo versao anterior de 'Base Nova.xlsx'...")
            try:
                os.remove(output_path)
            except OSError as exc:
                raise AutomationError(
                    f"Nao consegui substituir '{output_path}': {exc}\n"
                    "O arquivo provavelmente esta aberto no Excel. Feche-o e rode de novo."
                )

        log(f"Salvando como: {output_path}")
        _com_retry(lambda: workbook.SaveAs(output_path,
                                           FileFormat=XL_OPENXML_WORKBOOK),
                   "SaveAs")
        log("Salvo com sucesso.")

        if not os.path.exists(output_path):
            raise AutomationError(
                f"O Excel nao gravou o arquivo em: {output_path}\n"
                "Verifique permissoes/sincronizacao do OneDrive nessa pasta."
            )

        return output_path
    finally:
        _shutdown_excel(excel, workbook)
        pythoncom.CoUninitialize()


def _open_base(excel, iqy_path: str):
    """
    Abre a base e carrega os dados.

    Caminho principal: le a URL dentro do .iqy e monta a consulta web
    diretamente, o que dispara a carga na hora e de forma sincrona.
    Caminho alternativo: entrega o .iqy para o Excel interpretar.
    """
    if bool(getattr(config, "BUILD_QUERY_FROM_IQY", True)):
        try:
            return _build_workbook_from_iqy(excel, iqy_path)
        except AutomationError:
            raise
        except Exception as exc:
            log(f"Nao consegui montar a consulta manualmente ({exc}).")
            log("Vou pedir para o proprio Excel abrir o 'query.iqy'.")

    return _open_iqy_directly(excel, iqy_path)


def _build_workbook_from_iqy(excel, iqy_path: str):
    """Cria a pasta de trabalho e a consulta web a partir da URL do .iqy."""
    parsed = _parse_iqy(iqy_path)
    url = parsed["url"]
    options = parsed["options"]
    log(f"Consulta lida do '{os.path.basename(iqy_path)}': {url[:90]}...")

    workbook = _com_retry(lambda: excel.Workbooks.Add(), "Workbooks.Add")
    sheet = workbook.Worksheets(1)

    query_table = _com_retry(
        lambda: sheet.QueryTables.Add(Connection="URL;" + url,
                                      Destination=sheet.Range("A1")),
        "QueryTables.Add",
    )

    selection = options.get("selection", "").lower()
    formatting = options.get("formatting", "").lower()

    for attribute, value in (
        ("BackgroundQuery", False),           # sincrono: sabemos quando acabou
        ("WebSelectionType",
         XL_ALL_TABLES if "alltables" in selection else XL_ENTIRE_PAGE),
        ("WebFormatting",
         XL_WEB_FORMATTING_ALL if formatting == "all"
         else XL_WEB_FORMATTING_RTF if formatting == "rtf"
         else XL_WEB_FORMATTING_NONE),
        ("WebDisableDateRecognition", True),
        ("WebDisableRedirections", False),
        ("AdjustColumnWidth", False),         # o AutoFit vem depois
        ("SaveData", True),
        ("RefreshStyle", 1),                  # xlInsertDeleteCells
    ):
        try:
            setattr(query_table, attribute, value)
        except Exception:
            pass  # nem toda versao do Excel expoe todas as propriedades

    _refresh_and_wait(excel, workbook, query_table)
    return workbook


def _open_iqy_directly(excel, iqy_path: str):
    """Plano B: deixa o Excel interpretar o proprio arquivo .iqy."""
    log(f"Abrindo a base: {os.path.basename(iqy_path)}")
    workbook = _com_retry(
        lambda: excel.Workbooks.Open(os.path.abspath(iqy_path)),
        "Workbooks.Open",
        attempts=int(getattr(config, "EXCEL_OPEN_TIMEOUT_SECONDS", 120) / 10) or 6,
        pause=10.0,
    )
    try:
        _com_retry(lambda: workbook.EnableConnections(), "EnableConnections",
                   attempts=3, pause=1.0)
        log("Conexoes de dados externos habilitadas.")
    except Exception:
        pass  # algumas versoes ja abrem habilitado e lancam erro aqui

    _refresh_and_wait(excel, workbook, None)
    return workbook


def _refresh_and_wait(excel, workbook, query_table) -> None:
    """
    Dispara a carga e so volta quando os dados chegarem.

    A atualizacao roda em primeiro plano (BackgroundQuery=False), entao a
    propria chamada segura o script pelo tempo que a base levar - sem espera
    fixa: assim que a query termina, seguimos.
    """
    started = time.time()
    log("Carregando a base agora (contagem comecou)...")

    refreshed = False
    if query_table is not None:
        try:
            _com_retry(lambda: query_table.Refresh(BackgroundQuery=False),
                       "QueryTable.Refresh", attempts=3, pause=3.0)
            refreshed = True
        except Exception as exc:
            log(f"  aviso na atualizacao da consulta: {exc}")
    else:
        for sheet in workbook.Worksheets:
            try:
                tables = sheet.QueryTables
                for index in range(1, tables.Count + 1):
                    table = tables.Item(index)
                    try:
                        table.BackgroundQuery = False
                    except Exception:
                        pass
                    try:
                        _com_retry(lambda t=table: t.Refresh(BackgroundQuery=False),
                                   "QueryTable.Refresh", attempts=3, pause=3.0)
                        refreshed = True
                    except Exception as exc:
                        log(f"  aviso na atualizacao da consulta: {exc}")
            except Exception:
                continue

    if not refreshed:
        try:
            _com_retry(lambda: workbook.RefreshAll(), "RefreshAll",
                       attempts=3, pause=3.0)
        except Exception:
            pass

    # Se sobrou alguma consulta rodando em segundo plano, aguarda (com teto).
    deadline = time.time() + config.DATA_LOAD_TIMEOUT_SECONDS
    while _is_refreshing(workbook) and time.time() < deadline:
        time.sleep(1)

    try:
        excel.CalculateUntilAsyncQueriesDone()
    except Exception:
        pass

    rows = _used_rows(workbook)
    if rows <= 1:
        raise AutomationError(
            "A base nao carregou: a planilha continua vazia.\n"
            "Abra o 'query.iqy' manualmente uma vez, autorize o acesso ao "
            "SharePoint e rode o script de novo."
        )
    log(f"Base carregada: {rows} linhas em {time.time() - started:.0f}s.")


def _shutdown_excel(excel, workbook) -> None:
    """Fecha a pasta de trabalho e encerra o Excel (a nao ser que configurado)."""
    keep_open = bool(getattr(config, "KEEP_EXCEL_OPEN", False))
    try:
        if excel is not None:
            excel.ScreenUpdating = True
            excel.DisplayAlerts = True
            try:
                excel.EnableEvents = True
            except Exception:
                pass
    except Exception:
        pass

    if keep_open:
        try:
            if excel is not None:
                excel.Visible = True
        except Exception:
            pass
        return

    try:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
            log("Pasta de trabalho fechada.")
    except Exception:
        pass
    try:
        if excel is not None:
            excel.Quit()
            log("Excel encerrado.")
    except Exception:
        pass


def _is_refreshing(workbook) -> bool:
    try:
        for sheet in workbook.Worksheets:
            query_tables = sheet.QueryTables
            for index in range(1, query_tables.Count + 1):
                if query_tables.Item(index).Refreshing:
                    return True
    except Exception:
        return False
    return False


def _used_rows(workbook) -> int:
    try:
        return int(workbook.Worksheets(1).UsedRange.Rows.Count)
    except Exception:
        return 0


def _autofit_used_range(workbook) -> None:
    """
    Ctrl+A > Format > AutoFit Column Width e depois AutoFit Row Height.

    Aplicado sobre o UsedRange (a area que realmente tem dados). Rodar sobre
    Cells inteiro faria o Excel avaliar 16.384 colunas x 1.048.576 linhas -
    era isso que fazia a etapa levar varios minutos. O resultado visual e o
    mesmo, porque fora do UsedRange nao ha conteudo para ajustar.
    """
    for sheet in workbook.Worksheets:
        try:
            used = sheet.UsedRange
            started = time.time()
            used.EntireColumn.AutoFit()
            used.EntireRow.AutoFit()
            log(f"AutoFit aplicado em '{sheet.Name}' "
                f"({used.Rows.Count} x {used.Columns.Count}) "
                f"em {time.time() - started:.0f}s.")
        except Exception as exc:
            log(f"  aviso: AutoFit falhou em '{sheet.Name}' ({exc})")

    try:
        workbook.Worksheets(1).Activate()
    except Exception:
        pass


# --------------------------------------------------------------------------- #
# Orquestracao
# --------------------------------------------------------------------------- #


def main() -> int:
    global DESKTOP_DIR, DOWNLOADS_DIR

    log("=" * 70)
    log("SharePoint -> Excel | inicio")
    try:
        ensure_windows()

        DESKTOP_DIR = get_desktop_dir()
        DOWNLOADS_DIR = get_downloads_dir()
        log(f"Area de Trabalho: {DESKTOP_DIR}")
        log(f"Downloads.......: {DOWNLOADS_DIR}")

        iqy_path = download_export_file()
        if not os.path.isfile(iqy_path):
            raise AutomationError(f"Arquivo esperado nao existe: {iqy_path}")

        output = process_in_excel(iqy_path)

        if getattr(config, "DELETE_IQY_AFTER", False):
            try:
                os.remove(iqy_path)
                log("Arquivo 'query.iqy' removido.")
            except OSError:
                pass

        log("=" * 70)
        log(f"CONCLUIDO em {time.time() - _START:.0f}s! Arquivo pronto em:")
        log(f"  {output}")
        log("=" * 70)
        return 0
    except AutomationError as exc:
        log("")
        log("FALHOU: " + str(exc))
        return 1
    except Exception as exc:
        log("")
        log(f"ERRO INESPERADO: {exc}")
        traceback.print_exc()
        return 2


if __name__ == "__main__":
    sys.exit(main())
