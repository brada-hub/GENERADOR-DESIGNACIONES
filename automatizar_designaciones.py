import os
import sys
import time
import openpyxl
from playwright.sync_api import sync_playwright

# Asegurar codificación utf-8 en consola de Windows
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

# Directorios y rutas
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE = os.path.join(BASE_DIR, "DESIGNACIONES.xlsx")
OUTPUT_DIR = os.path.join(BASE_DIR, "Designaciones_2026")

URL_LOGIN = "http://181.188.185.214:8980/login.aspx"
URL_SISTEMA = "http://181.188.185.214:8980/administrador/Consultas%20de%20personal/DesignacionDocente_V2.aspx"

# Cargar variables de entorno desde archivo .env si existe
ENV_PATH = os.path.join(BASE_DIR, ".env")
if os.path.exists(ENV_PATH):
    with open(ENV_PATH, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, val = line.split("=", 1)
                os.environ.setdefault(key.strip(), val.strip().strip('"').strip("'"))

# Credenciales para Auto-Login (inicia sesión y se re-autentica automáticamente si expira)
USER = os.getenv("SISTEMA_USER", "")
PASS = os.getenv("SISTEMA_PASS", "")

# -------------------------------------------------------------------------
# REGLA DE FACULTAD UMBRELLA: FACEFA
# En la base de datos de la universidad, FACEFA sí jala todas las materias
# consolidadas de administración, contaduría, comercial, economía y complementarias.
# Para las Ingenierías (Sistemas, Electrónica, etc.), el sistema de la universidad
# requiere seleccionarlas por su carrera específica para que salgan sus materias.
# -------------------------------------------------------------------------
FACEFA_CARRERAS = {
    'ADMINISTRACION DE EMPRESAS', 'ADMINISTRACION DE EMPRESAS(R)',
    'CONTADURIA PUBLICA', 'CONTADURIA PUBLICA(R)',
    'INGENIERIA COMERCIAL', 'INGENIERIA COMERCIAL(R)',
    'ECONOMIA', 'ECONOMIA(R)',
    'COMPLEMENTARIA ADMINISTRACION DE EMPRESAS', 'COMPLEMENTARIA ADMINISTRACION DE EMPRESAS(N)',
    'COMPLEMENTARIA CONTADURIA PUBLICA', 'COMPLEMENTARIA CONTADURIA PUBLICA(N)',
    'COMPLEMENTARIA INGENIERIA COMERCIAL', 'COMPLEMENTARIA INGENIERIA COMERCIAL(N)'
}

def mapear_carrera_a_facultad(nombre_carrera):
    """
    - Ramas Económicas/Financieras/Administrativas -> 'FACEFA' (funciona unificado en la web).
    - Ingenierías y demás carreras -> se mantienen individuales para que el sistema liste sus materias.
    """
    nom = nombre_carrera.strip()
    nom_upper = nom.upper()
    
    if nom in FACEFA_CARRERAS or nom_upper in [c.upper() for c in FACEFA_CARRERAS]:
        return "FACEFA"
    return nom


def cargar_y_ordenar_designaciones(archivo_excel):
    """
    Lee todas las pestañas de las 8 sedes, aplica el mapeo de facultades
    (INGENIERIAS y FACEFA), consolida para emitir 1 sola hoja por docente
    dentro de su facultad o carrera, y ordena alfabéticamente por:
      Sede -> Carrera/Facultad -> Ape1 -> Ape2 -> Nombres
    """
    if not os.path.exists(archivo_excel):
        print(f"[!] Error: No se encontró el archivo '{archivo_excel}'.")
        sys.exit(1)

    print(f"[*] Leyendo el archivo '{archivo_excel}'...")
    wb = openpyxl.load_workbook(archivo_excel, read_only=True)
    
    registros = []
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue
        headers = [str(h).strip() if h else '' for h in rows[0]]
        for r in rows[1:]:
            if not any(r):
                continue
            item = dict(zip(headers, r))
            registros.append(item)
        print(f"    - Hoja cargada: {sheet}")

    print(f"[+] Total de materias/filas leídas en Excel: {len(registros)}")

    unicos = {}
    for r in registros:
        sede = str(r.get('Sede') or '').strip()
        carrera_original = str(r.get('nomCarrera') or '').strip()
        carrera_facultad = mapear_carrera_a_facultad(carrera_original)
        ci = str(r.get('Ci') or '').strip()

        if not ci or ci.lower() == 'none':
            continue

        # CLAVE UNICA: (Sede, Carrera o Facultad, CI)
        # Esto hace que un docente con materias en Sistemas y Electrónica genere 1 sola designación en INGENIERIAS.
        key = (sede, carrera_facultad, ci)
        if key not in unicos:
            unicos[key] = {
                'Sede': sede,
                'nomCarrera': carrera_facultad,
                'Ci': ci,
                'Ape1': str(r.get('Ape1') or '').strip(),
                'Ape2': str(r.get('Ape2') or '').strip(),
                'Nombres': str(r.get('Nombres') or '').strip()
            }

    lista = list(unicos.values())

    # ORDENAR: Sede -> Carrera/Facultad -> Apellido 1 -> Apellido 2 -> Nombres
    lista.sort(key=lambda x: (
        x['Sede'].upper(),
        x['nomCarrera'].upper(),
        x['Ape1'].upper(),
        x['Ape2'].upper(),
        x['Nombres'].upper()
    ))

    print(f"[OK] Total de hojas de designación a generar (con FACEFA unificado): {len(lista)}")
    return lista


def sanitizar_nombre(texto):
    caracteres = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
    for c in caracteres:
        texto = texto.replace(c, '_')
    return texto.strip()


def asegurar_sesion(page):
    """
    Verifica si la sesión está activa. Si redirigió a login.aspx
    o no están los controles del sistema, inicia sesión automáticamente.
    """
    if "login.aspx" in page.url or page.locator("#ctl00_ContentPlaceHolder1_DDListCarre").count() == 0:
        print("[*] Sesión expirada o no iniciada. Ejecutando Auto-Login...")
        page.goto(URL_LOGIN)
        page.wait_for_timeout(1000)
        
        page.locator("#Text2").fill(USER)
        page.locator("#Password2").fill(PASS)
        page.locator("#Checkbox2").check()
        page.locator("#Button3").click()
        
        # Esperar que procese el login
        page.wait_for_timeout(3000)
        
        # Ir a la página de designaciones
        page.goto(URL_SISTEMA)
        page.wait_for_selector("#ctl00_ContentPlaceHolder1_DDListCarre", timeout=15000)
        print("[OK] Sesión iniciada con éxito.")


def ejecutar_automatizacion(limite=None):
    designaciones = cargar_y_ordenar_designaciones(EXCEL_FILE)
    if limite:
        print(f"[*] MODO DE PRUEBA: Procesando únicamente los primeros {limite} docentes...")
        designaciones = designaciones[:limite]

    print("\n[*] Iniciando navegador Firefox...")
    with sync_playwright() as p:
        browser = p.firefox.launch(headless=True)
        ctx = browser.new_context(accept_downloads=True)

        page = ctx.new_page()
        # Aceptar automáticamente cualquier ventana emergente/alerta JS
        page.on("dialog", lambda dialog: dialog.accept())

        # Iniciar sesión al comienzo
        asegurar_sesion(page)

        total = len(designaciones)
        exitosos = 0
        fallidos = 0

        for i, doc in enumerate(designaciones, 1):
            sede = doc['Sede']
            carrera = doc['nomCarrera']
            ci = doc['Ci']
            ape1 = doc['Ape1']
            ape2 = doc['Ape2']
            nombres = doc['Nombres']

            nombre_legible = f"{ape1} {ape2} {nombres}".strip()
            archivo_nombre = sanitizar_nombre(f"{ape1}_{ape2}_{nombres}_CI_{ci}") + ".pdf"

            carpeta_destino = os.path.join(OUTPUT_DIR, sanitizar_nombre(sede), sanitizar_nombre(carrera))
            os.makedirs(carpeta_destino, exist_ok=True)
            ruta_pdf_final = os.path.join(carpeta_destino, archivo_nombre)

            print(f"\n[{i}/{total}] [{sede}] [{carrera}]")
            print(f"    -> Docente: {nombre_legible} (CI: {ci})")

            # Salta si ya existe en disco (permite pausar y reanudar)
            if os.path.exists(ruta_pdf_final) and os.path.getsize(ruta_pdf_final) > 1000:
                print(f"    [i] Ya existe en disco, omitiendo: {archivo_nombre}")
                exitosos += 1
                continue

            try:
                # Comprobar que no haya expirado la sesión antes de procesar
                asegurar_sesion(page)

                # 1. Seleccionar Carrera o Facultad (INGENIERIAS / FACEFA / carrera)
                page.select_option("#ctl00_ContentPlaceHolder1_DDListCarre", label=carrera)
                page.wait_for_timeout(1500)

                # 2. Seleccionar Sede
                page.select_option("#ctl00_ContentPlaceHolder1_DDListSede", label=sede)
                page.wait_for_timeout(1500)

                # 3. Llenar CI y presionar Tab
                campo_ci = page.locator("#ctl00_ContentPlaceHolder1_TBCi")
                campo_ci.fill(ci)
                campo_ci.press("Tab")
                page.wait_for_timeout(2500)

                # 4. Clic en 'Actualizar Para Imprimir'
                page.locator("#ctl00_ContentPlaceHolder1_BtnActualizaImpre").click()
                page.wait_for_timeout(3500)

                # 5. Clic en el icono de la impresora (Crystal Reports Toolbar)
                btn_print = page.locator("input[type='image'][title='Imprimir']").first
                btn_print.wait_for(timeout=8000)

                with page.expect_popup() as popup_info:
                    btn_print.click()

                popup = popup_info.value
                popup.wait_for_load_state()

                # 6. Disparar submit en el popup y capturar los bytes del PDF
                with popup.expect_response(
                    lambda r: "pdf" in r.headers.get("content-type", "").lower() or r.status == 200,
                    timeout=15000
                ) as resp_info:
                    popup.evaluate("checkValuesAndSubmit()")

                resp = resp_info.value
                pdf_bytes = resp.body()

                with open(ruta_pdf_final, "wb") as f_pdf:
                    f_pdf.write(pdf_bytes)

                if os.path.exists(ruta_pdf_final) and os.path.getsize(ruta_pdf_final) > 1000:
                    print(f"    [OK] PDF guardado con exito: {archivo_nombre} ({len(pdf_bytes)} bytes)")
                    exitosos += 1
                else:
                    print(f"    [!] Advertencia: Archivo vacio para CI {ci}")
                    fallidos += 1

                try:
                    popup.close()
                except:
                    pass

                time.sleep(1)

            except Exception as e:
                motivo = "No apareció botón imprimir (posiblemente sin materias asignadas)" if "Timeout" in str(e) else str(e)
                print(f"    [AVISO] Docente CI {ci} ({nombre_legible}): {motivo}")
                fallidos += 1
                
                with open(os.path.join(BASE_DIR, "docentes_sin_reporte.txt"), "a", encoding="utf-8") as f_err:
                    f_err.write(f"[{sede}] [{carrera}] CI: {ci} - {nombre_legible} | {motivo}\n")

                try:
                    page.goto(URL_SISTEMA)
                    page.wait_for_timeout(1500)
                except:
                    pass
                continue

        print("\n" + "="*50)
        print("  PROCESO COMPLETADO")
        print(f"  Total a procesar:  {total}")
        print(f"  Exitosos:          {exitosos}")
        print(f"  Fallidos:          {fallidos}")
        print(f"  Carpeta de salida: {OUTPUT_DIR}")
        print("="*50)
        browser.close()


if __name__ == "__main__":
    limite = int(sys.argv[1]) if len(sys.argv) > 1 and sys.argv[1].isdigit() else None
    ejecutar_automatizacion(limite=limite)
