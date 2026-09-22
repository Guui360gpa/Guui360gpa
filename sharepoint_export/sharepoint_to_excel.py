# -*- coding: utf-8 -*-
"""
Automatiza o fluxo completo:

  1. Abre o Microsoft Edge (perfil pessoal, maximizado) via Playwright.
  2. Acessa o SharePoint e clica em "Export to Excel".
  3. Intercepta o download (nome aleatorio do Playwright), renomeia para
     "query.iqy" e move para a Area de Trabalho.
  4. Abre o arquivo no Excel, libera a conexao externa e aguarda a carga
     da base (por padrao ate 3 minutos).
  5. Aplica AutoFit de largura de coluna e de altura de linha em toda a planilha.
  6. Salva como "Base Nova.xlsx" na Area de Trabalho.

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


def log(message: str) -> None:
    print(f"[{datetime.now():%H:%M:%S}] {message}", flush=True)


class AutomationError(RuntimeError):
    """Erro de negocio previsto pela automacao (mensagem amigavel)."""


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
            capture_output=True,
            text=True,
            timeout=30,
        ).stdout.lower()
    except Exception:
        return False
    return "msedge.exe" in output


# Pastas/arquivos que nunca precisam ser copiados (cache pesado e estado de
# sessao que o Edge mantem travado enquanto esta aberto).
PROFILE_SKIP_NAMES = {
    "cache", "code cache", "gpucache", "dawncache", "shadercache",
    "grshadercache", "media cache", "service worker", "crashpad",
    "sessions", "gcm store", "optimization guide model store",
    "component_crx_cache", "extensions_crx_cache", "safe browsing",
}


def _copy_profile_tolerant(src: str, dst: str) -> tuple[int, int]:
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
                # Arquivo em uso pelo Edge: segue o jogo.
                skipped += 1
    return copied, skipped


def prepare_profile_dir() -> str:
    """
    Devolve o diretorio de perfil que sera usado pelo Playwright.

    O Edge nao permite que dois processos usem o mesmo 'User Data' ao mesmo
    tempo. Por isso, por padrao, trabalhamos sobre uma copia do perfil. A copia
    e criada uma unica vez (marcador '.profile_seeded'); dali em diante ela tem
    a propria sessao logada e nao precisa mais ser sincronizada.
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

    already_seeded = os.path.isfile(marker) and os.path.isdir(dst_profile)
    if already_seeded and not getattr(config, "PROFILE_RESYNC_EACH_RUN", False):
        log("Usando a copia de perfil ja existente.")
        _clear_profile_locks(destination)
        return destination

    os.makedirs(destination, exist_ok=True)

    # Arquivos de estado global do Edge (necessarios para o perfil abrir).
    for name in ("Local State", "Last Version"):
        src_file = os.path.join(source, name)
        if os.path.isfile(src_file):
            try:
                shutil.copy2(src_file, os.path.join(destination, name))
            except OSError:
                pass

    log(f"Preparando copia do perfil do Edge ('{profile}')... pode demorar na 1a vez.")
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


def _clear_profile_locks(destination: str) -> None:
    """Remove travas remanescentes da copia do perfil."""
    for lock in ("SingletonLock", "SingletonCookie", "SingletonSocket"):
        lock_path = os.path.join(destination, lock)
        if os.path.exists(lock_path):
            try:
                os.remove(lock_path)
            except OSError:
                pass


# --------------------------------------------------------------------------- #
# Etapa 1 e 2 - Playwright / SharePoint
# --------------------------------------------------------------------------- #


def _snapshot_downloads() -> set:
    """Fotografa a pasta Downloads para detectar arquivos novos depois."""
    try:
        return set(os.listdir(config.DOWNLOADS_DIR))
    except OSError:
        return set()


def _find_new_download(before: set, deadline: float) -> Optional[str]:
    """
    Vigia a pasta Downloads e devolve o caminho do primeiro arquivo novo,
    ja finalizado (sem .crdownload / .tmp / .partial).
    """
    while time.time() < deadline:
        try:
            current = set(os.listdir(config.DOWNLOADS_DIR))
        except OSError:
            current = set()

        candidates = []
        for name in current - before:
            lowered = name.lower()
            if lowered.endswith((".crdownload", ".tmp", ".partial")):
                continue
            if not lowered.endswith(config.ACCEPTED_DOWNLOAD_SUFFIXES):
                continue
            candidates.append(os.path.join(config.DOWNLOADS_DIR, name))

        if candidates:
            newest = max(candidates, key=os.path.getmtime)
            if _file_is_stable(newest):
                return newest

        time.sleep(0.5)
    return None


def _file_is_stable(path: str, checks: int = 3, interval: float = 0.4) -> bool:
    """Confirma que o arquivo parou de crescer (download concluido)."""
    try:
        last = os.path.getsize(path)
    except OSError:
        return False
    for _ in range(checks):
        time.sleep(interval)
        try:
            size = os.path.getsize(path)
        except OSError:
            return False
        if size != last:
            last = size
            continue
    return True


def download_export_file() -> str:
    """
    Executa o fluxo no SharePoint e devolve o caminho do arquivo baixado
    ja renomeado para 'query.iqy' na Area de Trabalho.
    """
    from playwright.sync_api import sync_playwright

    user_data_dir = prepare_profile_dir()
    os.makedirs(config.DESKTOP_DIR, exist_ok=True)

    target_path = os.path.join(config.DESKTOP_DIR, config.IQY_FILENAME)
    captured: dict[str, str] = {}
    before = _snapshot_downloads()

    with sync_playwright() as playwright:
        log("Abrindo o Microsoft Edge (maximizado)...")
        context = playwright.chromium.launch_persistent_context(
            user_data_dir=user_data_dir,
            channel=config.BROWSER_CHANNEL,
            headless=False,
            accept_downloads=True,
            downloads_path=config.DOWNLOADS_DIR,
            no_viewport=True,  # necessario para a janela ficar realmente maximizada
            args=[
                "--start-maximized",
                f"--profile-directory={config.EDGE_PROFILE_DIRECTORY}",
                "--no-first-run",
                "--no-default-browser-check",
                "--disable-features=msEdgeIdentityRefresh",
            ],
        )

        def handle_download(download) -> None:
            """Intercepta o download em qualquer aba/popup do contexto."""
            if "path" in captured:
                return
            try:
                suggested = download.suggested_filename
                log(f"Download interceptado: {suggested}")
                download.save_as(target_path)
                captured["path"] = target_path
                log(f"Salvo como: {target_path}")
            except Exception as exc:  # pragma: no cover - depende do ambiente
                log(f"Nao foi possivel salvar direto pelo Playwright ({exc}). "
                    "Vou vigiar a pasta Downloads.")

        context.on("download", handle_download)

        page = context.pages[0] if context.pages else context.new_page()
        page.set_default_timeout(config.PAGE_LOAD_TIMEOUT_MS)

        log(f"Acessando: {config.SHAREPOINT_URL}")
        page.goto(config.SHAREPOINT_URL, wait_until="domcontentloaded",
                  timeout=config.PAGE_LOAD_TIMEOUT_MS)
        try:
            page.wait_for_load_state("networkidle", timeout=45_000)
        except Exception:
            pass  # o SharePoint mantem conexoes abertas; seguir em frente

        log(f"Procurando o menu '{config.EXPORT_MENU_ITEM}'...")
        clicked = _click_export(page)
        if not clicked:
            raise AutomationError(
                f"Nao encontrei o item '{config.EXPORT_MENU_ITEM}' na pagina. "
                "Confirme se a pagina carregou logada e se o nome do botao mudou."
            )

        log("Aguardando o arquivo exportado...")
        deadline = time.time() + config.DOWNLOAD_WAIT_SECONDS
        while time.time() < deadline and "path" not in captured:
            # Fecha popups vazios que o SharePoint abre so para disparar o download.
            for extra in list(context.pages)[1:]:
                try:
                    if extra.url in ("about:blank", "") and not extra.is_closed():
                        extra.close()
                except Exception:
                    pass
            time.sleep(0.5)

        if "path" not in captured:
            # Plano B: o Edge pode ter baixado direto para a pasta Downloads.
            found = _find_new_download(before, time.time() + 30)
            if found:
                log(f"Arquivo novo detectado em Downloads: {os.path.basename(found)}")
                captured["path"] = _move_to_desktop(found, target_path)

        if not config.KEEP_BROWSER_OPEN:
            try:
                context.close()
            except Exception:
                pass

    if "path" not in captured:
        raise AutomationError(
            "O download nao foi concluido dentro do tempo limite "
            f"({config.DOWNLOAD_WAIT_SECONDS}s)."
        )
    return captured["path"]


def _click_export(page) -> bool:
    """
    Clica em 'Export to Excel'. Tenta varias estrategias porque o SharePoint
    renderiza o comando ora como menuitem, ora como botao, as vezes dentro
    do menu 'Export' ou do overflow ('...').
    """
    name = config.EXPORT_MENU_ITEM

    def try_click(locator, description: str) -> bool:
        try:
            if locator.count() == 0:
                return False
            target = locator.first
            target.wait_for(state="visible", timeout=10_000)
            target.scroll_into_view_if_needed(timeout=5_000)
            target.click(timeout=15_000)
            log(f"Clique efetuado em: {description}")
            return True
        except Exception:
            return False

    # 1) Caminho direto (igual ao gravado pelo codegen).
    if try_click(page.get_by_role("menuitem", name=name), "menuitem 'Export to Excel'"):
        return True

    # 2) Abrir menus intermediarios e tentar de novo.
    for opener in ("Export", "More options", "More", "..."):
        try:
            button = page.get_by_role("button", name=opener, exact=False)
            if button.count():
                button.first.click(timeout=8_000)
                page.wait_for_timeout(1_200)
                if try_click(page.get_by_role("menuitem", name=name),
                             f"menuitem apos abrir '{opener}'"):
                    return True
        except Exception:
            continue

    # 3) Botao/link com o mesmo rotulo.
    for locator, description in (
        (page.get_by_role("button", name=name), "button 'Export to Excel'"),
        (page.get_by_role("link", name=name), "link 'Export to Excel'"),
        (page.get_by_text(name, exact=False), "texto 'Export to Excel'"),
    ):
        if try_click(locator, description):
            return True

    # 4) Dentro de iframes (web parts classicas).
    for frame in page.frames:
        if frame is page.main_frame:
            continue
        for locator, description in (
            (frame.get_by_role("menuitem", name=name), "menuitem (iframe)"),
            (frame.get_by_role("button", name=name), "button (iframe)"),
            (frame.get_by_text(name, exact=False), "texto (iframe)"),
        ):
            if try_click(locator, description):
                return True

    return False


def _move_to_desktop(source: str, target_path: str) -> str:
    """Move o arquivo baixado para a Area de Trabalho com o nome definitivo."""
    os.makedirs(config.DESKTOP_DIR, exist_ok=True)
    if os.path.exists(target_path):
        os.remove(target_path)
    shutil.move(source, target_path)
    log(f"Arquivo movido para: {target_path}")
    return target_path


# --------------------------------------------------------------------------- #
# Etapas 4, 5 e 6 - Excel
# --------------------------------------------------------------------------- #

XL_OPENXML_WORKBOOK = 51          # xlOpenXMLWorkbook (.xlsx)
MSO_AUTOMATION_SECURITY_LOW = 1   # msoAutomationSecurityLow


def process_in_excel(iqy_path: str) -> str:
    """Abre o .iqy, aguarda a carga, aplica AutoFit e salva como .xlsx."""
    import pythoncom
    import win32com.client as win32

    output_path = os.path.join(config.DESKTOP_DIR, config.OUTPUT_XLSX_NAME)

    pythoncom.CoInitialize()
    excel = None
    workbook = None
    try:
        log("Abrindo o Microsoft Excel...")
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = True
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        try:
            excel.AutomationSecurity = MSO_AUTOMATION_SECURITY_LOW
        except Exception:
            pass
        try:
            excel.WindowState = -4137  # xlMaximized
        except Exception:
            pass

        log(f"Abrindo a base: {os.path.basename(iqy_path)}")
        workbook = excel.Workbooks.Open(os.path.abspath(iqy_path))

        _enable_and_refresh_connections(excel, workbook)
        _wait_for_data(excel, workbook)
        _autofit_all(workbook)

        if os.path.exists(output_path):
            log("Removendo versao anterior de 'Base Nova.xlsx'...")
            os.remove(output_path)

        log(f"Salvando como: {output_path}")
        workbook.SaveAs(output_path, FileFormat=XL_OPENXML_WORKBOOK)

        if not config.KEEP_EXCEL_OPEN:
            workbook.Close(SaveChanges=False)
            excel.Quit()
        else:
            excel.DisplayAlerts = True
            excel.Visible = True

        return output_path
    except Exception:
        # Em caso de falha, nao deixa instancias fantasma do Excel abertas.
        try:
            if workbook is not None and not config.KEEP_EXCEL_OPEN:
                workbook.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if excel is not None and not config.KEEP_EXCEL_OPEN:
                excel.Quit()
        except Exception:
            pass
        raise
    finally:
        pythoncom.CoUninitialize()


def _enable_and_refresh_connections(excel, workbook) -> None:
    """
    Libera as conexoes de dados externos (o equivalente a clicar em
    'Habilitar conteudo' / 'Enable external data') e dispara a atualizacao
    em primeiro plano, para conseguirmos aguardar de forma confiavel.
    """
    try:
        workbook.EnableConnections()
        log("Conexoes de dados externos habilitadas.")
    except Exception:
        pass  # algumas versoes ja abrem habilitado e lancam erro aqui

    refreshed = False
    for sheet in workbook.Worksheets:
        try:
            query_tables = sheet.QueryTables
            for index in range(1, query_tables.Count + 1):
                query_table = query_tables.Item(index)
                try:
                    query_table.BackgroundQuery = False
                except Exception:
                    pass
                try:
                    query_table.Refresh(BackgroundQuery=False)
                    refreshed = True
                except Exception:
                    pass
        except Exception:
            continue

    if not refreshed:
        try:
            workbook.RefreshAll()
            log("RefreshAll disparado.")
        except Exception:
            pass


def _wait_for_data(excel, workbook) -> None:
    """
    Aguarda a base carregar. Comeca pela espera fixa configurada (3 min por
    padrao) e, se a query ainda estiver atualizando, concede uma folga extra.
    """
    total = config.DATA_LOAD_WAIT_SECONDS
    log(f"Aguardando a carga da base (ate {total // 60} min {total % 60}s)...")

    deadline = time.time() + total
    while time.time() < deadline:
        if not _is_refreshing(workbook) and _has_data(workbook):
            # Pequena estabilizacao antes de formatar.
            time.sleep(3)
            if not _is_refreshing(workbook):
                log("Base carregada.")
                break
        time.sleep(2)
    else:
        log("Tempo padrao esgotado; verificando se a query ainda esta rodando...")

    grace_deadline = time.time() + config.DATA_LOAD_EXTRA_GRACE_SECONDS
    while _is_refreshing(workbook) and time.time() < grace_deadline:
        time.sleep(3)

    try:
        excel.CalculateUntilAsyncQueriesDone()
    except Exception:
        pass

    if not _has_data(workbook):
        raise AutomationError(
            "A base nao carregou: a planilha continua vazia. "
            "Abra o 'query.iqy' manualmente uma vez para validar o acesso "
            "ao SharePoint pelo Excel e rode o script de novo."
        )


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


def _has_data(workbook) -> bool:
    """Considera carregada quando ha mais de uma linha preenchida."""
    try:
        sheet = workbook.Worksheets(1)
        used = sheet.UsedRange
        return used.Rows.Count > 1 and used.Columns.Count >= 1 and \
            str(sheet.Cells(1, 1).Value or "").strip() != ""
    except Exception:
        return False


def _autofit_all(workbook) -> None:
    """Ctrl+A > Format > AutoFit Column Width e, depois, AutoFit Row Height."""
    for sheet in workbook.Worksheets:
        try:
            sheet.Activate()
        except Exception:
            pass
        try:
            log(f"AutoFit de largura das colunas em '{sheet.Name}'...")
            sheet.Cells.EntireColumn.AutoFit()
        except Exception as exc:
            log(f"  aviso: nao foi possivel ajustar colunas ({exc})")
        try:
            log(f"AutoFit de altura das linhas em '{sheet.Name}'...")
            sheet.Cells.EntireRow.AutoFit()
        except Exception as exc:
            log(f"  aviso: nao foi possivel ajustar linhas ({exc})")
    try:
        workbook.Worksheets(1).Activate()
        workbook.Worksheets(1).Range("A1").Select()
    except Exception:
        pass


# --------------------------------------------------------------------------- #
# Orquestracao
# --------------------------------------------------------------------------- #


def main() -> int:
    log("=" * 68)
    log("SharePoint -> Excel | inicio")
    log("=" * 68)
    try:
        ensure_windows()

        iqy_path = download_export_file()

        if not os.path.isfile(iqy_path):
            raise AutomationError(f"Arquivo esperado nao existe: {iqy_path}")

        output = process_in_excel(iqy_path)

        log("=" * 68)
        log(f"CONCLUIDO! Arquivo pronto em: {output}")
        log("=" * 68)
        return 0
    except AutomationError as exc:
        log("")
        log("FALHOU: " + str(exc))
        return 1
    except Exception as exc:  # pragma: no cover
        log("")
        log(f"ERRO INESPERADO: {exc}")
        traceback.print_exc()
        return 2


if __name__ == "__main__":
    sys.exit(main())
