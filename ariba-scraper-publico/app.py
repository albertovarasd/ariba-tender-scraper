import asyncio
import os
import re
import sys
import time
from pathlib import Path
from urllib.parse import urlparse
from dotenv import load_dotenv
from playwright.async_api import async_playwright
from file_filter import should_ignore
from dom_selectors import SELECTORS
from google_io import GoogleIo
from datetime import datetime
import zipfile
import shutil

# Fuerza UTF-8 en la salida: redirigida a un log (ejecutar_diario.bat,
# Programador de Tareas) Windows usa cp1252 y los emojis la rompen.
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

load_dotenv()

# Soporte opcional para .rar: 'rarfile' puede no estar instalado en el entorno.
try:
    import rarfile
    RARFILE_AVAILABLE = True
    # Usa el UnRAR.exe embebido si está, en vez del PATH. Debe ser la versión de
    # consola: la GUI abre una ventana y deja colgada la corrida desatendida.
    _bundled_unrar = Path(__file__).parent / "UnRAR.exe"
    if _bundled_unrar.exists():
        rarfile.UNRAR_TOOL = str(_bundled_unrar)
except ImportError:
    RARFILE_AVAILABLE = False
    print("⚠️ Paquete 'rarfile' no instalado: los .rar se subirán sin descomprimir. "
          "Ejecuta 'pip install rarfile' para habilitar la extracción.")

# Variables de ambiente
ARIBA_LOGIN_URL = os.getenv("ARIBA_LOGIN_URL")
ARIBA_EMAIL = os.getenv("ARIBA_EMAIL")
ARIBA_PASSWORD = os.getenv("ARIBA_PASSWORD")
GOOGLE_SERVICE_ACCOUNT_FILE = os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE")
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID")
GOOGLE_DRIVE_FOLDER_ID = os.getenv("GOOGLE_DRIVE_FOLDER_ID")
DOWNLOAD_TEMP = Path(os.getenv("DOWNLOAD_TEMP", "./downloads"))
DOWNLOAD_TEMP.mkdir(parents=True, exist_ok=True)
HEADLESS = os.getenv("HEADLESS", "true").lower() == "true"

# Nombre del comprador tal como aparece en el menú de Ariba. COMPRADOR_PATRON
# permite una expresión regular propia (ej. "Empresa Demo|Demo S\.A\.").
COMPRADOR_NOMBRE = (os.getenv("COMPRADOR_NOMBRE") or "").strip()
COMPRADOR_PATRON = os.getenv("COMPRADOR_PATRON") or re.escape(COMPRADOR_NOMBRE)

ALLOWED_EXTENSIONS = {'.pdf', '.dwg', '.xls', '.xlsx', '.ppt', '.pptx'}

def is_allowed(filename):
    return any(filename.lower().endswith(ext) for ext in ALLOWED_EXTENSIONS)

def process_archive(filepath, folder_id, g):
    """
    Descomprime un .zip o .rar y sube a Drive solo los archivos permitidos
    (ver ALLOWED_EXTENSIONS). Si la extracción falla o no aplica (p. ej. .rar
    sin 'rarfile'/unrar disponible), se sube el archivo comprimido tal cual
    para no perder el adjunto silenciosamente.
    """
    temp_dir = filepath.parent / f"extracted_{filepath.stem}"
    temp_dir.mkdir(exist_ok=True)
    extracted_files = []

    try:
        if filepath.suffix.lower() == '.zip':
            with zipfile.ZipFile(filepath, 'r') as archive:
                for member in archive.namelist():
                    if is_allowed(member):
                        archive.extract(member, temp_dir)
                        extracted_files.append(temp_dir / member)

        elif filepath.suffix.lower() == '.rar':
            if not RARFILE_AVAILABLE:
                raise RuntimeError("'rarfile' no está disponible en este entorno")
            with rarfile.RarFile(filepath, 'r') as archive:
                for member in archive.namelist():
                    if is_allowed(member):
                        archive.extract(member, temp_dir)
                        extracted_files.append(temp_dir / member)

    except Exception as e:
        print(f"⚠️ No se pudo descomprimir '{filepath.name}' ({e}). "
              f"Subiendo el archivo comprimido sin descomprimir.")
        g.upload_file(filepath, folder_id)
        shutil.rmtree(temp_dir, ignore_errors=True)
        return

    if not extracted_files:
        # Nada permitido adentro (o venía vacío): se sube el original igual.
        print(f"⚠️ No se encontraron archivos permitidos dentro de '{filepath.name}'. "
              f"Subiendo el archivo comprimido original como respaldo.")
        g.upload_file(filepath, folder_id)
    else:
        for file in extracted_files:
            g.upload_file(file, folder_id)
            print(f"Archivo extraído y subido a Drive: {file}")

    shutil.rmtree(temp_dir, ignore_errors=True)

# -----------------------
# Login / Web scraping
# -----------------------
def validar_configuracion():
    """Falla de inmediato, con un mensaje claro, si falta alguna variable del .env."""
    requeridas = {
        "ARIBA_LOGIN_URL": ARIBA_LOGIN_URL,
        "ARIBA_EMAIL": ARIBA_EMAIL,
        "ARIBA_PASSWORD": ARIBA_PASSWORD,
        "GOOGLE_SERVICE_ACCOUNT_FILE": GOOGLE_SERVICE_ACCOUNT_FILE,
        "GOOGLE_SHEET_ID": GOOGLE_SHEET_ID,
        "GOOGLE_DRIVE_FOLDER_ID": GOOGLE_DRIVE_FOLDER_ID,
        "COMPRADOR_NOMBRE": COMPRADOR_NOMBRE,
    }
    faltantes = [nombre for nombre, valor in requeridas.items() if not valor]
    if faltantes:
        raise SystemExit(
            "Faltan variables en el archivo .env: " + ", ".join(faltantes) +
            "\nCopia .env.example a .env y completa los valores."
        )


async def scrap():
    validar_configuracion()
    nuevas_filas = 0
    start_time = time.time()
    # Cargar configuración
    g = GoogleIo(GOOGLE_SERVICE_ACCOUNT_FILE, sheet_id=GOOGLE_SHEET_ID, root_folder_id=GOOGLE_DRIVE_FOLDER_ID)    

    execution_date = datetime.now().strftime("%d/%m/%y")

    async with async_playwright() as p:
        # Viewport fijo: en headless Ariba renderizaba menos filas de su tabla de
        # scroll virtual que en modo visible, y se perdían licitaciones.
        browser = await p.chromium.launch(
            headless=HEADLESS,
            args=["--disable-blink-features=AutomationControlled", "--window-size=1920,1080"]
        )
        context = await browser.new_context(viewport={"width": 1920, "height": 1080})

        # Algunos adjuntos (ej. .tif) el navegador los muestra en vez de descargarlos
        # si falta Content-Disposition: attachment, y nunca llega el evento de descarga.
        async def _forzar_descarga_imagenes(route, request):
            try:
                response = await route.fetch()
                headers = {k.lower(): v for k, v in response.headers.items()}
                content_type = headers.get("content-type", "")
                if content_type.startswith("image/") and "content-disposition" not in headers:
                    headers["content-disposition"] = "attachment"
                    await route.fulfill(response=response, headers=headers)
                else:
                    await route.fulfill(response=response)
            except Exception:
                # Si falla la intercepción, dejar pasar la solicitud sin romper la navegación.
                await route.continue_()

        await context.route("**/s3.ariba.com/**", _forzar_descarga_imagenes)
        await context.route("**/*.tif", _forzar_descarga_imagenes)
        await context.route("**/*.tiff", _forzar_descarga_imagenes)

        page = await context.new_page()

        print("Abriendo página de login...")
        await page.goto(ARIBA_LOGIN_URL, timeout=180000)

        print("Ingresando email...")
        await page.wait_for_selector(SELECTORS["username"], timeout=120000)
        await page.fill(SELECTORS["username"], ARIBA_EMAIL)
        await page.wait_for_selector(SELECTORS["username_button"], timeout=60000)
        await page.locator(SELECTORS["username_button"]).click()

        print("Ingresando password...")
        await page.wait_for_selector(SELECTORS["password"], timeout=180000)
        await page.fill(SELECTORS["password"], ARIBA_PASSWORD)
        await asyncio.sleep(3)
        await page.locator(SELECTORS["password_button"]).click()

        print("Esperando redirección post-login...")
        await page.wait_for_load_state("domcontentloaded", timeout=60000)
        await asyncio.sleep(5)
        print(f"URL actual post-login: {page.url}")
        print(f"Título de página: {await page.title()}")

        # Detección condicional de cookies y popup
        print("Verificando si aparece banner de aceptación de cookies...")
        try:
            await page.wait_for_selector(SELECTORS["cookie_accept"], timeout=10000)
            print("Banner de cookies detectado. Presionando 'Aceptar todo'...")
            await page.locator(SELECTORS["cookie_accept"]).click()
            await asyncio.sleep(2)
        except Exception:
            print("No apareció ningún banner de cookies (o ya fue aceptado).")

        print("Verificando si aparece popup de publicidad...")
        try:
            await page.wait_for_selector(SELECTORS["ad_popup_close"], timeout=10000)
            print("Popup detectado. Cerrando...")
            await page.locator(SELECTORS["ad_popup_close"]).click()
            await asyncio.sleep(1)
        except Exception:
            print("No apareció ningún popup. Continuando con el proceso...")

        print("Esperando para abrir menu superior...")
        try:
            await page.wait_for_selector(SELECTORS["upper_menu"], timeout=60000)
        except Exception as err:
            await page.screenshot(path="login_debug.png", full_page=True)
            print(f"❌ Error al buscar el menú superior. URL: {page.url}")
            print(f"📸 Se guardó una captura de pantalla del estado actual en 'login_debug.png'")
            raise err

        # Re-verificar si el banner de cookies interfiere antes de hacer click
        try:
            cookie_btn = page.locator(SELECTORS["cookie_accept"])
            if await cookie_btn.is_visible():
                await cookie_btn.click()
                await asyncio.sleep(1)
        except Exception:
            pass

        await asyncio.sleep(3)
        await page.locator(SELECTORS["upper_menu"]).click()


        await asyncio.sleep(5)
        print("Esperando para seleccionar elemento de propuestas...")
        await page.wait_for_selector(SELECTORS["list_item"], timeout=180000)
        await asyncio.sleep(5)

        # El clic encadena redirecciones de OAuth que pueden tardar: se reintenta con
        # más margen en vez de abortar por un timeout aislado.
        clic_exitoso = False
        for intento_menu in range(1, 4):
            try:
                await page.locator(SELECTORS["list_item"]).click(timeout=60000)
                clic_exitoso = True
                break
            except Exception as e:
                print(f"⚠️ Timeout/erro al hacer clic en el menú de propuestas (intento {intento_menu}): {e}")
                if intento_menu < 3:
                    await asyncio.sleep(5)

        if not clic_exitoso:
            raise RuntimeError(
                "No se pudo hacer clic en 'Ariba Proposals & Questionnaires' tras 3 "
                "intentos. Probablemente el servidor/red estaba muy lento en ese "
                "momento. Revisa la conexión e intenta correr el script de nuevo."
            )

        await asyncio.sleep(10)


        print("Esperando para presionar menu secundario...")
        comprador_seleccionado = False
        for intento_comprador in range(1, 4):
            try:
                dropdown_locator = page.locator('div.w-dashboard-tabs a.w-tabitem-a, a[bh="PML"]').first
                await dropdown_locator.wait_for(state="visible", timeout=30000)
                box = await dropdown_locator.bounding_box()
                if box:
                    await page.mouse.move(box["x"] + box["width"]/2, box["y"] + box["height"]/2)
                    print(f"Presionando menu secundario... (intento {intento_comprador})")
                    await page.mouse.down()
                    await asyncio.sleep(1)
                    await page.mouse.up()
                    await asyncio.sleep(3)

                # Imprimir opciones disponibles en el menú de compradores para depuración
                await asyncio.sleep(2)
                menu_items = await page.locator('a.w-pmi-item, div.awmenu a').all()
                found_buyers = []
                for item in menu_items:
                    t = await item.text_content()
                    if t:
                        found_buyers.append(t.strip())
                print(f"🔍 Opciones encontradas en el menú de compradores ({len(found_buyers)}): {found_buyers}")

                # Buscar la opción del comprador configurado (COMPRADOR_NOMBRE) dentro del menú emergente
                comprador_item = page.locator('a.w-pmi-item, div.awmenu a', has_text=re.compile(COMPRADOR_PATRON, re.IGNORECASE)).first
                if await comprador_item.count() > 0:
                    print(f"Seleccionando comprador: {COMPRADOR_NOMBRE}...")
                    await comprador_item.wait_for(state="visible", timeout=10000)
                    await comprador_item.click()
                    await asyncio.sleep(8)
                    comprador_seleccionado = True
                    break
                else:
                    print(f"⚠️ No se desplegó o no se encontró '{COMPRADOR_NOMBRE}' en el menú (intento {intento_comprador}).")

            except Exception as e:
                print(f"⚠️ Error durante la selección del comprador (intento {intento_comprador}): {e}")

            if intento_comprador < 3:
                print("   Reintentando abrir el menú de compradores...")
                await asyncio.sleep(3)

        if not comprador_seleccionado:
            # Sin confirmar el comprador, el resto del scraping leería datos del comprador
            # equivocado sin avisar: mejor cortar acá.
            raise RuntimeError(
                f"No se pudo seleccionar '{COMPRADOR_NOMBRE}' en el menú de "
                "compradores después de 3 intentos. Abortando para no scrapear datos "
                "del comprador equivocado. Revisa 'table_debug.png' / vuelve a intentar."
            )

        # networkidle puede no cumplirse nunca por el tráfico de fondo de Ariba; si no
        # se logra en 30 s se sigue, los pasos siguientes tienen sus propias esperas.
        try:
            await page.wait_for_load_state("networkidle", timeout=30000)
        except Exception as e:
            print(f"⚠️ No se alcanzó 'networkidle' en 30s (se continúa igual, probablemente la página ya está lista): {e}")


        print("Esperando iframe...")
        iframe_el = await page.wait_for_selector(SELECTORS["iframe"], timeout=120000)
        frame = await iframe_el.content_frame()
        if frame is None:
            raise RuntimeError("No se pudo obtener el content_frame del iframe")

        print("Cargando siguiente pagina...")
        await frame.wait_for_selector("body", timeout=120000)

        print("Antes de intentar click dentro del iframe...")
        clicked = False
        try:
            target_link = frame.locator(SELECTORS["active_proposals_link"]).first
            if await target_link.count() > 0 and await target_link.is_visible():
                await target_link.click(force=True)
                clicked = True
                print("✅ Clic exitoso en el enlace de propuestas activas.")
        except Exception as e:
            print(f"⚠️ No se encontró por selector exacto ({e}). Probando por texto...")

        if not clicked:
            for text_pattern in ["Propuestas", "Active", "Cuestionarios"]:
                loc = frame.locator("a", has_text=re.compile(text_pattern, re.IGNORECASE)).first
                if await loc.count() > 0 and await loc.is_visible():
                    try:
                        await loc.click(force=True)
                        clicked = True
                        print(f"✅ Clic realizado mediante texto '{text_pattern}'.")
                        break
                    except Exception:
                        continue

        if not clicked:
            print("⚠️ No se requirió hacer clic o no se encontró enlace secundario; continuando a la tabla...")

        await asyncio.sleep(5)

        # Captura de pantalla para inspeccionar el estado actual de la tabla dentro del iframe
        await page.screenshot(path="table_debug.png", full_page=True)
        print("📸 Captura de pantalla guardada en 'table_debug.png'")

        # --- Expandir solo el grupo de licitaciones abiertas ---
        # Ariba los muestra colapsados. El clic va en el <span class="w-togglebox-icon-off">,
        # no en la celda. La interfaz puede estar en español o inglés: se prueban ambas.
        total_abiertas_esperadas = None
        _VARIANTES_ABIERTO = [
            "Estado: Abierto",   # español completo
            "Status: Abierto",   # interfaz en inglés, valor en español (visto en el servidor)
            "Status: Open",      # inglés completo
            "Estado: Open",      # mezcla inversa, por si acaso
        ]
        try:
            fila_abierto = None
            variante_encontrada = None
            for _variante in _VARIANTES_ABIERTO:
                _candidata = frame.locator("tr", has_text=_variante).first
                if await _candidata.count() > 0:
                    fila_abierto = _candidata
                    variante_encontrada = _variante
                    break

            if fila_abierto is not None:
                # El total que muestra Ariba ("Estado: Abierto (109)"), para comparar al final.
                texto_fila_abierto = await fila_abierto.text_content()
                match_total = re.search(r"(?:Estado|Status):\s*(?:Abierto|Open)\D*(\d+)",
                                        texto_fila_abierto or "", re.IGNORECASE)
                if match_total:
                    total_abiertas_esperadas = int(match_total.group(1))
                    print(f"🔢 Ariba reporta {total_abiertas_esperadas} licitaciones en '{variante_encontrada}'.")
                else:
                    print(f"⚠️ No se pudo leer el número total de '{variante_encontrada}' (formato inesperado).")

                toggle_icon = fila_abierto.locator(
                    "span.w-togglebox-icon-off, span.w-togglebox-icon-on"
                ).first
                if await toggle_icon.count() > 0:
                    await toggle_icon.click()
                    await asyncio.sleep(3)
                    print(f"🔽 Grupo '{variante_encontrada}' expandido.")
                else:
                    print(f"⚠️ No se encontró el ícono de expansión (w-togglebox-icon-off/on) en la fila '{variante_encontrada}'.")
            else:
                print(f"⚠️ No se encontró la fila de licitaciones abiertas para expandir. "
                      f"Variantes probadas: {_VARIANTES_ABIERTO}. "
                      f"Revisa 'table_debug.png' para ver cómo se ve la página.")
        except Exception as e:
            print(f"⚠️ No se pudo expandir el grupo de licitaciones abiertas: {e}")

        # Nueva captura ya con el grupo expandido, útil para depurar si algo falla
        await page.screenshot(path="table_debug_expandido.png", full_page=True)
        print("📸 Captura de pantalla (post-expansión) guardada en 'table_debug_expandido.png'")

        # Diagnóstico de estructura dentro del iframe
        all_tr = await frame.locator("tr").all()
        print(f"🔍 Total de filas 'tr' encontradas dentro del iframe: {len(all_tr)}")

        # Seleccionar todas las filas
        print("Iniciando extracción de filas con scroll virtual...")

        processed_ids = set()
        processed_urls = set()  # dedup adicional por URL (más confiable que el ID
                                 # cuando una fila del scroll virtual aún no terminó
                                 # de cargar y el ID sale "raro")
        entries = []
        col_ids = set(g.sheet.col_values(4))  # IDs existentes en Google Sheet
        col_urls = set(g.sheet.col_values(9))  # Links ya guardados (columna I) — permite
                                                # saltar de inmediato una fila ya conocida
                                                # SIN gastar los 4 reintentos de carga.
        descartadas_link_fantasma = 0
        descartadas_forma = 0
        descartadas_duplicado = 0
        filas_no_cargaron = 0

        # --- Varias pasadas completas de scroll ---
        # Con muchas filas cambiando a la vez, un reintento por fila no alcanza. Se
        # repite el recorrido hasta MAX_PASADAS, cortando si dos pasadas seguidas no
        # suman nada o si ya se alcanzó el total esperado.
        MAX_PASADAS = 5
        for numero_pasada in range(1, MAX_PASADAS + 1):
            print(f"\n🔁 Pasada de scroll #{numero_pasada} de {MAX_PASADAS}...")
            nuevos_antes_de_pasada = len(entries)

            # Volver al inicio antes de cada pasada (la primera ya arranca arriba).
            if numero_pasada > 1:
                await frame.evaluate("""
                    (containerSelector) => {
                        const container = document.querySelector(containerSelector);
                        if (container) container.scrollTop = 0;
                    }
                """, SELECTORS["table_container"])
                await asyncio.sleep(3)

            while True:
                # La tabla real no siempre usa tableRow1/tableRow2 (Ariba reutiliza esas clases),
                # así que se filtra por contenido: links fantasma y textos de widget fuera.
                rows = await frame.locator("tr.tableRow1, tr.tableRow2, tr[class*='tableRow'], tr:has(a)").all()

                current_frame_url = frame.url

                for row in rows:
                    try:
                        link = await row.evaluate("""row => {
                            const a = row.querySelector('a');
                            return a ? a.href : null;
                        }""")
                        if not link:
                            continue

                        # Links fantasma: los <a href="#"> de acordeones y filtros. Se compara solo
                        # dominio + ruta, ignorando parámetros de sesión (awssk=...) que cambian siempre.
                        _link_sin_query = urlparse(link)
                        _frame_sin_query = urlparse(current_frame_url)
                        if (_link_sin_query.netloc == _frame_sin_query.netloc
                                and _link_sin_query.path == _frame_sin_query.path):
                            descartadas_link_fantasma += 1
                            continue

                        # El href está disponible antes que el texto de la fila: si ya está en el Sheet
                        # o ya se vio en esta corrida, se salta sin gastar los 4 reintentos de carga.
                        if link in col_urls or link in processed_urls:
                            descartadas_duplicado += 1
                            continue

                        # ":scope > td" limita a los <td> hijos directos: cada celda de Ariba trae tablas
                        # anidadas y recorrer todo el árbol corría los índices ("Fecha de Cierre" salía
                        # vacía). En scroll virtual la fila puede existir antes de que el AJAX la
                        # rellene, así que se relee hasta 4 veces.
                        entry_id = ""
                        name_val = ""
                        end_time_val = ""
                        cell_texts = []
                        fila_completa = False

                        for intento in range(4):
                            tds = await row.locator(":scope > td").all()
                            if not tds:
                                # Sin ninguna celda <td> no es una fila de datos: no vale la pena reintentar.
                                break

                            cell_texts = []
                            for td in tds:
                                txt = (await td.text_content()).strip()
                                if txt:
                                    cell_texts.append(txt)

                            if not cell_texts:
                                # Igual que arriba: celdas vacías de entrada son basura de interfaz.
                                break

                            name_el = row.locator("a span, a").first
                            name_val = (await name_el.text_content()).strip() if await name_el.count() > 0 else cell_texts[0]
                            entry_id = cell_texts[1] if len(cell_texts) > 1 else cell_texts[0]
                            end_time_val = cell_texts[2] if len(cell_texts) > 2 else ""

                            def _normalizar(txt):
                                return re.sub(r"\s+", " ", txt).strip()

                            entry_id = _normalizar(entry_id)
                            name_val = _normalizar(name_val)

                            if entry_id != name_val and end_time_val:
                                fila_completa = True
                                break

                            # Contenido parcial: la fila está cargando de verdad, vale la pena reintentar.
                            await asyncio.sleep(0.75)

                        if not cell_texts:
                            continue

                        if not fila_completa:
                            # Tras 4 intentos (~3 s) se descarta. Se muestra el link para revisarla a mano.
                            print(f"   ↳ ⏳ fila no terminó de cargar tras varios intentos, se omite: "
                                  f"nombre={name_val[:60]!r} link={link}")
                            continue

                        # Palabras que solo aparecen en widgets de interfaz o en el encabezado de la
                        # tabla, nunca en un título o ID real.
                        _MARCAS_BASURA = ("function(", "ariba.Dom", "RJS(", "ariba.Chooser",
                                           "ariba.AWWidgets", "removeRelocatedCopy",
                                           "Hora de finalización", "No hay artículos")

                        def _parece_basura(txt, max_len=300):
                            if not txt or len(txt) > max_len:
                                return True
                            if txt in ("Título", "ID", "Estado", "Tipo de evento", "Mercancía", "Regiones"):
                                return True
                            return any(marca in txt for marca in _MARCAS_BASURA)

                        if _parece_basura(entry_id) or _parece_basura(name_val):
                            descartadas_forma += 1
                            if descartadas_forma <= 5:
                                print(f"   ↳ (muestra descarte #{descartadas_forma}) id={entry_id[:80]!r} name={name_val[:80]!r}")
                            continue

                        # Duplicados entre scrolls: por ID y por URL (cubre una fila redetectada con otro ID).
                        if entry_id in processed_ids or entry_id in col_ids or link in processed_urls:
                            descartadas_duplicado += 1
                            continue

                        entry = {
                            "id": entry_id,
                            "url": link,
                            "name": name_val,
                            "end_time": end_time_val
                        }


                        processed_ids.add(entry_id)
                        processed_urls.add(link)
                        entries.append(entry)
                        print(f"✅ Nuevo registro detectado: {entry_id}")

                    except Exception as e:
                        print(f"⚠️ Error procesando fila: {e}")

                scrolled = await frame.evaluate("""
                    (containerSelector) => {
                        const container = document.querySelector(containerSelector);
                        if (!container) return true;  // si no se encuentra, consideramos que terminó

                        const before = container.scrollTop;
                        container.scrollBy(0, container.clientHeight * 0.9);
                        const after = container.scrollTop;

                        // Si el scrollTop no cambia, llegamos al final
                        return before === after;
                    }
                """, SELECTORS["table_container"])

                if scrolled:
                    print("✅ Fin del scroll detectado (contenedor al final).")
                    break

                await asyncio.sleep(5)  # Esperar carga de nuevas filas

            # Nuevas de esta corrida + ya conocidas, que también cuentan como capturadas.
            total_procesadas_hasta_ahora = len(entries) + descartadas_duplicado

            # Caso 1: ya está todo lo que Ariba dice que hay.
            if total_abiertas_esperadas is not None and total_procesadas_hasta_ahora >= total_abiertas_esperadas:
                print(f"   (pasada #{numero_pasada}: ya se cubrieron las {total_abiertas_esperadas} "
                      f"licitaciones esperadas, no hace falta seguir)")
                break

            # Caso 2: pasada sin nada nuevo. Se permite una más (el scroll virtual puede
            # comportarse distinto), pero dos seguidas en vano cortan el ciclo.
            if len(entries) == nuevos_antes_de_pasada:
                print(f"   (pasada #{numero_pasada} no sumó licitaciones nuevas)")
                if numero_pasada > 1:
                    break

        print(f"Total de registros nuevos: {len(entries)}")
        print(f"🔎 Diagnóstico de filas descartadas → link fantasma: {descartadas_link_fantasma}, "
              f"forma inválida (basura de interfaz): {descartadas_forma}, duplicadas: {descartadas_duplicado}")

        # --- Verificación de completitud ---
        # Nuevas + duplicadas = licitaciones reales identificadas, contra el total que
        # Ariba mostraba en "Estado: Abierto (XXX)".
        total_procesadas = len(entries) + descartadas_duplicado
        if total_abiertas_esperadas is not None:
            faltantes = total_abiertas_esperadas - total_procesadas
            if faltantes <= 0:
                print(f"✅ Verificación de completitud: se esperaban {total_abiertas_esperadas} licitaciones "
                      f"abiertas y se contabilizaron {total_procesadas} (nuevas + ya conocidas). Todo cuadra.")
            else:
                print(f"⚠️ Verificación de completitud: Ariba reporta {total_abiertas_esperadas} licitaciones "
                      f"abiertas, pero esta corrida solo contabilizó {total_procesadas} (nuevas + ya conocidas), "
                      f"incluso después de {numero_pasada} pasada(s) de scroll. Podrían faltar hasta {faltantes} "
                      f"licitaciones sin capturar — vuelve a correr el script para intentar capturarlas.")
        else:
            print("⚠️ No se pudo verificar completitud (no se leyó el total esperado de 'Estado: Abierto').")

        for entry in entries:
            try:
                folder_url = ""
                entry_description = ""
                product_quantity = ""
                archivos_descargados_nombres = []  # Para verificar después si el plano mencionado en el texto sí se descargó
                print("Visitando link: "+entry["url"])
                await page.goto(entry["url"], timeout=60000)
                await asyncio.sleep(10);
                try:
                    # La etiqueta puede venir en español o inglés según el idioma de Ariba.
                    _ETIQUETAS_TEXTO_LARGO = ["Texto Largo:", "Long Text:", "Long Description:", "Texto largo:"]
                    td_label = None
                    etiqueta_usada = None
                    for _etiqueta in _ETIQUETAS_TEXTO_LARGO:
                        _cand = page.locator("td", has_text=_etiqueta)
                        if await _cand.count() > 0:
                            td_label = _cand
                            etiqueta_usada = _etiqueta
                            break

                    if td_label is not None:
                        td_value = td_label.locator("xpath=following-sibling::td[1]")
                        entry_description = await td_value.text_content()
                        if entry_description:
                            entry_description = entry_description.strip()
                            print(f"Texto Largo encontrado ('{etiqueta_usada}'): {entry_description}")
                        else:
                            print(f"'{etiqueta_usada}' encontrado pero sin valor")
                    else:
                        print("No se encontró etiqueta de Texto Largo en esta página (ni en español ni en inglés)")
                except Exception as e:
                    print(f"Error al buscar el Texto Largo: {e}")

                # --- Respaldo: las plantillas tipo "ECO" no usan "Texto Largo:" ---
                # Ahí la descripción vive en una celda con colspan, de ancho y estilo variables:
                # se toma la de texto más largo, en todos los frames (puede ir dentro de un iframe).
                if not entry_description:
                    try:
                        mejor_texto = ""
                        candidatos_totales = 0
                        for fr in page.frames:
                            try:
                                candidatos = fr.locator('td[colspan]')
                                count_fr = await candidatos.count()
                            except Exception:
                                continue
                            candidatos_totales += count_fr
                            for idx in range(count_fr):
                                try:
                                    texto_alt = await candidatos.nth(idx).text_content()
                                except Exception:
                                    continue
                                if texto_alt:
                                    texto_alt = texto_alt.strip()
                                    if len(texto_alt) > len(mejor_texto):
                                        mejor_texto = texto_alt
                        print(f"🔍 Respaldo Texto Largo: {candidatos_totales} celda(s) con colspan revisadas en "
                              f"{len(page.frames)} frame(s); mejor candidato ({len(mejor_texto)} caracteres): "
                              f"{mejor_texto[:80]!r}")
                        # Descartar navegación de Ariba: las licitaciones de servicios no tienen ficha de
                        # producto y guardaban "< Volver al panel de instrucciones..." como descripción.
                        _BASURA_NAVEGACION = (
                            # Español
                            "Volver al panel",
                            "Sincronizar archivos con el escritorio",
                            "Descargar tutoriales",
                            "Mensajes de evento",
                            "Equipo de respuesta",
                            # Inglés (por si la interfaz está en ese idioma)
                            "Go back to",
                            "Sync files with desktop",
                            "Download Tutorials",
                            "Event Messages",
                            "Response Team",
                        )
                        _es_navegacion = any(marca in mejor_texto for marca in _BASURA_NAVEGACION)

                        if len(mejor_texto) > 60 and not _es_navegacion:
                            entry_description = mejor_texto
                        elif _es_navegacion:
                            print("⚠️ Esta licitación no tiene ficha de producto (probablemente es de "
                                  "servicios/proyecto). Se deja el texto largo vacío en vez de guardar "
                                  "texto de navegación.")
                            print(f"Texto Largo (formato alternativo) encontrado: {entry_description}")
                    except Exception as e:
                        print(f"Error al buscar Texto Largo (formato alternativo): {e}")
                try:
                    td_quantity = page.locator("td", has_text="Less...")
                    td_quantity_es = page.locator("td", has_text="Menos...")
                    td_q = td_quantity if (await td_quantity.count()) > 0 else td_quantity_es
                    if td_q:
                        td_qvalue = td_q.locator("xpath=following-sibling::td[1]")
                        product_quantity = await td_qvalue.text_content()
                        if product_quantity:
                            product_quantity = product_quantity.strip()
                            print(f"Cantidad del producto: {product_quantity}")
                        else:
                            print("'Cantidad:' encontrada pero sin valor")
                    else:
                        print("No se encontró 'Less...' en esta página")
                except Exception as e:
                    print(f"Error al buscar 'Less...': {e}")
                links = await page.locator(SELECTORS["download_links"]).all()
                print(f"Links encontrados: {len(links)}")
                for element in links:
                  try:
                    await element.click()
                    await page.wait_for_selector(SELECTORS["link_popup"], timeout=30000)
                    popup = page.locator(SELECTORS["link_popup"])
                    menu_items = popup.locator("a[role=menuitem]")
                    await asyncio.sleep(1)
                    count = await menu_items.count()
                    if count >= 2:
                        last_text = await menu_items.nth(count - 1).text_content()
                        if last_text and ("Download all attachments" in last_text or "Descargar todos los adjuntos" in last_text):
                            print(f"Encontrado popup con archivos descargables: {count-1}")
                            for i in range(count - 1):
                                try:
                                    await element.click()
                                    await page.wait_for_selector(SELECTORS["link_popup"], timeout=10000)
                                    _popup = page.locator(SELECTORS["link_popup"])
                                    await asyncio.sleep(1)
                                    _menu_items = _popup.locator("a[role=menuitem]")
                                    raw_text = await _menu_items.nth(i).text_content()
                                    if not raw_text:
                                        continue
                                    filename = raw_text.strip()

                                    if should_ignore(filename):
                                        print(f"⏩ Archivo ignorado por filtro: {filename}")
                                        continue
                                    async with page.expect_download(timeout=30000) as dl_info:
                                        await _menu_items.nth(i).click()
                                    download = await dl_info.value

                                    file_name = download.suggested_filename or f"{entry['id']}_{i}.dat"

                                    # El texto del menú no siempre coincide con el nombre real del archivo ya
                                    # descargado, así que se filtra también por ese nombre antes de subirlo.
                                    if should_ignore(file_name):
                                        print(f"⏩ Archivo ignorado por filtro (por nombre real tras descargar): {file_name}")
                                        continue

                                    saved_path = DOWNLOAD_TEMP / file_name

                                    await download.save_as(str(saved_path))
                                    print(f"Archivo descargado: {saved_path}")
                                    archivos_descargados_nombres.append(file_name)
                                    folder_id, folder_link = g.ensure_folder(entry["id"])
                                    folder_url = f"https://drive.google.com/drive/folders/{folder_id}"

                                    if saved_path.suffix.lower() in {'.zip', '.rar'}:
                                        process_archive(saved_path, folder_id, g)
                                    else:
                                        file_id, file_link = g.upload_file(saved_path, folder_id)
                                        print(f"Archivo subido a Drive: {file_link}")
                                    await asyncio.sleep(8)

                                except Exception as e:
                                    print(f"Error al descargar/subir archivo #{i}: {e}")
                        else:
                            print("No se encontró popup con archivos descargables")
                    else:
                        print(f"Se esperaban al menos 2 elementos en popup, encontrados {count}")
                  except Exception as e:
                    # Un fallo en este adjunto no debe cortar el resto de la licitación.
                    print(f"⚠️ Error procesando un link de descarga (se continúa con el siguiente): {e}")

                # --- ¿Falta el plano? ---
                # Busca "NUMERO PLANO:" / "PLANO:" / "N° PLANO:" en el texto largo y avisa si
                # ningún archivo descargado trae ese código en el nombre.
                match_plano = re.search(
                    r'(?:N[ÚU]MERO\s+PLANO|N[°º]?\s*PLANO|PLANO)\s*:?\s*([A-Za-z0-9][A-Za-z0-9\-_./]{2,40})',
                    entry_description, re.IGNORECASE
                )
                if match_plano:
                    codigo_plano = match_plano.group(1).strip().rstrip('.,;')
                    codigo_normalizado = re.sub(r'[^A-Za-z0-9]', '', codigo_plano).lower()
                    encontrado = any(
                        codigo_normalizado in re.sub(r'[^A-Za-z0-9]', '', nombre).lower()
                        for nombre in archivos_descargados_nombres
                    )
                    if not encontrado:
                        print(f"⚠️ PLANO MENCIONADO PERO NO DESCARGADO: el texto menciona el plano "
                              f"'{codigo_plano}', pero ningún archivo descargado ({archivos_descargados_nombres}) "
                              f"coincide con ese código. Revisar manualmente: {entry['url']}")

                g.insert_rows([[execution_date, "NO", entry["name"], entry["id"], entry["end_time"], entry_description, product_quantity, "Pendiente revisión", entry["url"], folder_url]])
                nuevas_filas += 1
            except Exception as e:
                # Una licitación caída no debe perder el resto del trabajo.
                print(f"❌ Error procesando la licitación {entry.get('id','?')} (se continúa con la siguiente): {e}")

                # Si se cayó el navegador entero, las siguientes fallarían en cascada: se reabre.
                if "browser has been closed" in str(e) or "Target page, context or browser" in str(e):
                    print("🔄 El navegador se cerró inesperadamente. Reabriéndolo para continuar...")
                    try:
                        try:
                            await browser.close()
                        except Exception:
                            pass

                        browser = await p.chromium.launch(
                            headless=HEADLESS,
                            args=["--disable-blink-features=AutomationControlled", "--window-size=1920,1080"]
                        )
                        context = await browser.new_context(viewport={"width": 1920, "height": 1080})
                        await context.route("**/s3.ariba.com/**", _forzar_descarga_imagenes)
                        await context.route("**/*.tif", _forzar_descarga_imagenes)
                        await context.route("**/*.tiff", _forzar_descarga_imagenes)
                        page = await context.new_page()

                        # El navegador nuevo no conserva la sesión.
                        print("🔑 Volviendo a iniciar sesión...")
                        await page.goto(ARIBA_LOGIN_URL, timeout=180000)
                        await page.wait_for_selector(SELECTORS["username"], timeout=120000)
                        await page.fill(SELECTORS["username"], ARIBA_EMAIL)
                        await page.wait_for_selector(SELECTORS["username_button"], timeout=60000)
                        await page.locator(SELECTORS["username_button"]).click()
                        await page.wait_for_selector(SELECTORS["password"], timeout=180000)
                        await page.fill(SELECTORS["password"], ARIBA_PASSWORD)
                        await asyncio.sleep(3)
                        await page.locator(SELECTORS["password_button"]).click()
                        await asyncio.sleep(15)
                        print("✅ Sesión restablecida, continuando con las licitaciones restantes.")
                    except Exception as e2:
                        print(f"❌ No se pudo reabrir el navegador: {e2}. Se detiene el procesamiento.")
                        break

        end_time = time.time()
        elapsed = end_time - start_time
        minutes, seconds = divmod(int(elapsed), 60)
        print(f"\n✅ Proceso terminado. Nuevas filas creadas: {nuevas_filas}. Tiempo total: {elapsed:.1f} segundos.")


# -----------------------
# CLI
# -----------------------
def print_usage():
    print("""
Uso:
  python app.py bootstrap      # Ejecuta scraping, guarda en Google Sheet y sube docs a Drive

Configuración en .env obligatoria:
  ARIBA_LOGIN_URL, ARIBA_EMAIL, ARIBA_PASSWORD,
  GOOGLE_SERVICE_ACCOUNT_FILE, GOOGLE_SHEET_ID, GOOGLE_DRIVE_FOLDER_ID
""")

if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print_usage()
        raise SystemExit(1)
    cmd = sys.argv[1].lower()
    if cmd == "bootstrap":
        asyncio.run(scrap())
    else:
        print_usage()