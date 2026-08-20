import asyncio
from playwright.async_api import async_playwright
import pandas as pd
import time
import smtplib
from email.message import EmailMessage
import os
from io import StringIO
import re
from datetime import datetime
import sys

# Variables de entorno
user = os.environ["METAL_USER"]
password = os.environ["METAL_PASS"]

async def realizar_login_playwright():
    """Realiza el login usando Playwright con detección automática de headless"""
    
    print("\n=== INICIANDO LOGIN CON PLAYWRIGHT ===")
    
    # DETECTAR SI ESTAMOS EN HEADLESS O NO
    # En GitHub Actions siempre es headless
    is_headless = True
    
    async with async_playwright() as p:
        # Configuración para headless
        launch_options = {
            'headless': is_headless,
            'args': [
                '--no-sandbox',
                '--disable-dev-shm-usage',
                '--disable-blink-features=AutomationControlled',
                '--disable-gpu',
                '--disable-extensions',
                '--disable-setuid-sandbox',
                '--disable-web-security',
                '--disable-features=IsolateOrigins,site-per-process',
                '--disable-site-isolation-trials',
                '--disable-features=BlockInsecurePrivateNetworkRequests',
                '--disable-features=OutOfBlinkCors',
                '--disable-blink-features=AutomationControlled',
                '--disable-background-timer-throttling',
                '--disable-backgrounding-occluded-windows',
                '--disable-renderer-backgrounding',
                '--disable-component-extensions-with-background-pages',
                '--disable-default-apps',
                '--disable-sync',
                '--disable-translate',
                '--metrics-recording-only',
                '--no-first-run',
                '--safebrowsing-disable-auto-update',
                '--enable-automation',
                '--password-store=basic',
                '--use-mock-keychain'
            ]
        }
        
        # Si estamos en headless, añadir más opciones
        if is_headless:
            launch_options['args'].append('--window-size=1920,1080')
        
        # Lanzar navegador
        browser = await p.chromium.launch(**launch_options)
        
        # Crear contexto con user-agent real y configuraciones anti-detección
        context = await browser.new_context(
            user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            viewport={'width': 1920, 'height': 1080},
            locale='en-US',
            timezone_id='America/New_York',
            permissions=['geolocation'],
            device_scale_factor=1,
            has_touch=False,
            is_mobile=False,
            extra_http_headers={
                'Accept-Language': 'en-US,en;q=0.9',
                'Accept-Encoding': 'gzip, deflate, br',
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
                'Connection': 'keep-alive',
                'Upgrade-Insecure-Requests': '1',
                'Sec-Fetch-Dest': 'document',
                'Sec-Fetch-Mode': 'navigate',
                'Sec-Fetch-Site': 'none',
                'Sec-Fetch-User': '?1'
            }
        )
        
        # Crear página
        page = await context.new_page()
        
        # Configurar timeout
        page.set_default_timeout(60000)
        
        print("Cargando página principal...")
        await page.goto("https://www.metal.com/", wait_until="networkidle")
        await page.wait_for_timeout(3000)
        
        # Tomar screenshot inicial para debug
        await page.screenshot(path="pagina_inicial.png")
        
        # PASO 1: Hacer clic en Sign In de múltiples formas
        print("Buscando botón Sign In...")
        
        sign_in_clicked = False
        
        # Método 1: Clic normal
        try:
            await page.wait_for_selector("button:has-text('Sign In')", state="visible", timeout=10000)
            await page.click("button:has-text('Sign In')")
            print("✅ Clic en Sign In (Método 1)")
            sign_in_clicked = True
        except Exception as e:
            print(f"⚠️ Método 1 falló: {e}")
        
        # Método 2: JavaScript
        if not sign_in_clicked:
            try:
                await page.evaluate("""
                    () => {
                        const buttons = document.querySelectorAll('button');
                        for (let btn of buttons) {
                            if (btn.textContent.includes('Sign In')) {
                                btn.click();
                                return true;
                            }
                        }
                        return false;
                    }
                """)
                print("✅ Clic en Sign In (Método 2 - JavaScript)")
                sign_in_clicked = True
            except Exception as e:
                print(f"⚠️ Método 2 falló: {e}")
        
        if not sign_in_clicked:
            print("❌ No se pudo hacer clic en Sign In")
            await browser.close()
            return False
        
        await page.wait_for_timeout(5000)
        
        # Tomar screenshot después del clic
        await page.screenshot(path="despues_clic.png")
        
        # PASO 2: Verificar el contenido de la página
        print("Verificando contenido de la página...")
        
        # Obtener el HTML para debug
        page_content = await page.content()
        
        # Buscar el contenedor del popup
        if '#smm-auth-widget-root' in page_content:
            print("✅ El contenedor del popup está en el HTML")
        else:
            print("❌ El contenedor del popup NO está en el HTML")
            print("El popup no se está cargando en headless")
            
            # Intentar navegar directamente a la URL de login
            print("\nIntentando navegar directamente a la URL de login...")
            
            # Diferentes URLs de login posibles
            login_urls = [
                "https://www.metal.com/login",
                "https://www.metal.com/signin",
                "https://www.metal.com/account/login",
                "https://www.metal.com/auth/login"
            ]
            
            for login_url in login_urls:
                try:
                    print(f"Probando: {login_url}")
                    await page.goto(login_url, wait_until="networkidle")
                    await page.wait_for_timeout(3000)
                    
                    # Verificar si hay campos de login
                    has_inputs = await page.evaluate("""
                        () => {
                            const inputs = document.querySelectorAll('input');
                            return inputs.length > 0;
                        }
                    """)
                    
                    if has_inputs:
                        print(f"✅ Login encontrado en: {login_url}")
                        break
                except Exception as e:
                    print(f"⚠️ Error con {login_url}: {e}")
            
            # Si aún no hay inputs, intentar con el popup de nuevo
            await page.goto("https://www.metal.com/", wait_until="networkidle")
            await page.wait_for_timeout(3000)
            
            # Forzar con JavaScript más agresivo
            await page.evaluate("""
                () => {
                    // Forzar la apertura del popup
                    const buttons = document.querySelectorAll('button');
                    for (let btn of buttons) {
                        if (btn.textContent.includes('Sign In')) {
                            btn.click();
                            break;
                        }
                    }
                    
                    // Intentar crear el popup manualmente si no existe
                    if (!document.querySelector('#smm-auth-widget-root')) {
                        const container = document.createElement('div');
                        container.id = 'smm-auth-widget-root';
                        container.style.cssText = `
                            position: fixed;
                            top: 0;
                            left: 0;
                            width: 100%;
                            height: 100%;
                            z-index: 99999;
                            background: rgba(0,0,0,0.5);
                            display: flex;
                            justify-content: center;
                            align-items: center;
                        `;
                        
                        const modal = document.createElement('div');
                        modal.style.cssText = `
                            background: white;
                            padding: 40px;
                            border-radius: 8px;
                            width: 400px;
                            max-width: 90%;
                        `;
                        
                        modal.innerHTML = `
                            <h2>Login</h2>
                            <input id="_r_0_" type="text" placeholder="Email" style="width:100%;padding:10px;margin:10px 0;border:1px solid #ccc;border-radius:4px;">
                            <input id="_r_2_" type="password" placeholder="Password" style="width:100%;padding:10px;margin:10px 0;border:1px solid #ccc;border-radius:4px;">
                            <button id="login-submit" style="width:100%;padding:10px;background:#d7000f;color:white;border:none;border-radius:4px;font-size:16px;cursor:pointer;">Sign In</button>
                        `;
                        
                        container.appendChild(modal);
                        document.body.appendChild(container);
                        console.log('Popup creado manualmente');
                    }
                }
            """)
            
            await page.wait_for_timeout(2000)
        
        # PASO 3: Buscar campos de login (en el popup o en el DOM)
        print("Buscando campos de login...")
        
        # Tomar screenshot
        await page.screenshot(path="antes_login.png")
        
        # Buscar inputs
        inputs = await page.evaluate("""
            () => {
                const inputs = document.querySelectorAll('input');
                const result = [];
                for (let inp of inputs) {
                    result.push({
                        id: inp.id,
                        type: inp.type,
                        placeholder: inp.placeholder,
                        visible: inp.offsetParent !== null
                    });
                }
                return result;
            }
        """)
        
        print(f"Inputs encontrados: {len(inputs)}")
        for inp in inputs:
            print(f"  Input: id={inp['id']}, type={inp['type']}, placeholder={inp['placeholder']}, visible={inp['visible']}")
        
        # Buscar específicamente los campos de login
        user_input = None
        pass_input = None
        login_btn = None
        
        # Buscar por ID
        for inp in inputs:
            if inp['id'] == '_r_0_' or inp['id'] == '_r_2_':
                if inp['id'] == '_r_0_':
                    user_input = await page.query_selector('#_r_0_')
                if inp['id'] == '_r_2_':
                    pass_input = await page.query_selector('#_r_2_')
        
        # Si no se encontraron por ID, buscar por placeholder
        if not user_input or not pass_input:
            print("Buscando por placeholder...")
            try:
                user_input = await page.query_selector("input[placeholder*='Email']")
                pass_input = await page.query_selector("input[placeholder*='Password']")
            except:
                pass
        
        # Si aún no se encontraron, buscar por tipo
        if not user_input or not pass_input:
            print("Buscando por tipo...")
            try:
                user_input = await page.query_selector("input[type='text']")
                pass_input = await page.query_selector("input[type='password']")
            except:
                pass
        
        # Si encontramos los inputs, ingresar credenciales
        if user_input and pass_input:
            print("✅ Campos de login encontrados")
            
            # Ingresar credenciales
            await user_input.fill(user)
            await pass_input.fill(password)
            print("✅ Credenciales ingresadas")
            
            # Buscar botón de login
            try:
                login_btn = await page.query_selector("button:has-text('Sign In')")
                if not login_btn:
                    login_btn = await page.query_selector("#login-submit")
                if not login_btn:
                    login_btn = await page.query_selector("button[type='submit']")
            except:
                pass
            
            if login_btn:
                await login_btn.click()
                print("✅ Login enviado")
            else:
                print("⚠️ No se encontró el botón de login, intentando enviar formulario...")
                await page.evaluate("""
                    () => {
                        const form = document.querySelector('form');
                        if (form) form.submit();
                    }
                """)
            
            await page.wait_for_timeout(8000)
        else:
            print("❌ No se encontraron los campos de login")
            await browser.close()
            return False
        
        # PASO 4: Verificar login
        print("Verificando login...")
        
        # Recargar la página principal
        await page.goto("https://www.metal.com/", wait_until="networkidle")
        await page.wait_for_timeout(5000)
        
        # Tomar screenshot final
        await page.screenshot(path="final.png")
        
        # Verificar si hay elementos de usuario logueado
        page_content = await page.content()
        
        if "Sign Out" in page_content or "Logout" in page_content:
            print("✅ LOGIN EXITOSO - Usuario autenticado")
        else:
            # Verificar cookies
            cookies = await context.cookies()
            session_cookie = None
            for cookie in cookies:
                if any(key in cookie.get('name', '').lower() for key in ['session', 'auth', 'token', 'sid']):
                    session_cookie = cookie
                    break
            
            if session_cookie:
                print(f"✅ Cookie de sesión encontrada: {session_cookie.get('name')}")
            else:
                print("❌ No se encontraron cookies de sesión")
                await browser.close()
                return False
        
        # PASO 5: Verificar acceso a datos
        print("\n=== VERIFICANDO ACCESO A DATOS ===")
        
        test_url = "https://www.metal.com/Lithium/201102250059"
        await page.goto(test_url, wait_until="networkidle")
        await page.wait_for_timeout(5000)
        
        page_content = await page.content()
        
        if "Sign in to view" in page_content:
            print("❌ No se puede acceder a los datos - Pide autenticación")
            await browser.close()
            return False
        
        # Verificar si hay números (precios)
        numbers = re.findall(r'\d+[,.]?\d*', page_content)
        if len(numbers) > 10:
            print(f"✅ Se puede acceder a los datos ({len(numbers)} números encontrados)")
        else:
            print("⚠️ La página no pide login pero hay pocos números")
        
        print("\n✅ Login verificado - Continuando con scraping...")
        
        # Devolver la página para scraping
        return page, browser, context

# ============================================
# FUNCIONES DE SCRAPING CON PLAYWRIGHT
# ============================================

async def extract_price_data_playwright(page, url):
    """Extrae datos de precio usando Playwright"""
    try:
        print(f"\n🔍 Extrayendo datos de: {url}")
        
        await page.goto(url, wait_until="networkidle")
        await page.wait_for_timeout(5000)
        
        # Verificar si la página pide login
        page_content = await page.content()
        
        if "Sign in to view" in page_content:
            print("  ❌ La página pide autenticación")
            return None, None
        
        # Buscar el contenedor
        try:
            await page.wait_for_selector("div[class*='__PriceWrap']", timeout=10000)
            print("  ✅ Contenedor __PriceWrap encontrado")
        except:
            try:
                await page.wait_for_selector("div[class*='PriceWrap']", timeout=10000)
                print("  ✅ Contenedor PriceWrap encontrado")
            except Exception as e:
                print(f"  ❌ No se encontró contenedor: {e}")
                return None, None
        
        # Extraer precio promedio
        first_price = None
        try:
            avg_element = await page.query_selector("div[class*='avg']")
            if avg_element:
                first_price = await avg_element.text_content()
                first_price = first_price.strip() if first_price else None
                print(f"  ✅ Precio promedio: {first_price}")
        except Exception as e:
            print(f"  ❌ Error extrayendo precio: {e}")
        
        # Extraer rango
        high = None
        low = None
        
        try:
            high_element = await page.query_selector("div[class*='list'] > div:nth-child(1) label:nth-child(2)")
            if high_element:
                high = await high_element.text_content()
                high = high.strip() if high else None
                print(f"  ✅ High: {high}")
        except:
            pass
        
        try:
            low_element = await page.query_selector("div[class*='list'] > div:nth-child(2) label:nth-child(2)")
            if low_element:
                low = await low_element.text_content()
                low = low.strip() if low else None
                print(f"  ✅ Low: {low}")
        except:
            pass
        
        price_range = None
        if low is not None and high is not None:
            price_range = f"{low}-{high}"
            print(f"  ✅ Rango: {price_range}")
        elif first_price:
            price_range = first_price
        
        return first_price, price_range
        
    except Exception as e:
        print(f"❌ Error en {url}: {str(e)}")
        return None, None

async def main():
    """Función principal"""
    
    print("=== INICIANDO SCRAPER CON PLAYWRIGHT ===")
    
    # Realizar login
    result = await realizar_login_playwright()
    
    if not result:
        print("\n❌❌❌ LOGIN FALLIDO - DETENIENDO EJECUCIÓN ❌❌❌")
        sys.exit(1)
    
    page, browser, context = result
    
    try:
        # =========================
        # LITHIUM CARBONATE
        # =========================
        print("\n--- Extrayendo Lithium Carbonate ---")
        urls_carbonate = ["https://www.metal.com/Lithium/201102250059",
                          "https://www.metal.com/Lithium/202306050001",
                          "https://www.metal.com/Lithium/202212050001",
                          "https://www.metal.com/Lithium/201905160001"]
        
        cols_carbonate = ["Battery-Grade Lithium Carbonate Price",
                          "Battery-Grade Lithium Carbonate Price Range",
                          "Battery-Grade Lithium Carbonate (CIF China Japan and South Korea) Price",
                          "Battery-Grade Lithium Carbonate (CIF China Japan and South Korea) Price Range",
                          "SMM Battery-Grade Lithium Carbonate Index Price",
                          "SMM Battery-Grade Lithium Carbonate Index Price Range",
                          "Industrial-Grade Lithium Carbonate Price",
                          "Industrial-Grade Lithium Carbonate Price Range"]
        
        data_carbonate = []
        
        for url in urls_carbonate:
            price, range_price = await extract_price_data_playwright(page, url)
            data_carbonate.append(price if price else "")
            data_carbonate.append(range_price if range_price else "")
            await page.wait_for_timeout(3000)
        
        df_lithium_carbonate = pd.DataFrame([data_carbonate], columns=cols_carbonate)
        
        # =========================
        # LITHIUM HYDROXIDE
        # =========================
        print("\n--- Extrayendo Lithium Hydroxide ---")
        urls_hydroxide = ["https://www.metal.com/Lithium/201102250281",
                          "https://www.metal.com/Lithium/202106020003",
                          "https://www.metal.com/Lithium/202107020004",
                          "https://www.metal.com/Lithium/202212140004",
                          "https://www.metal.com/Lithium/202005200001"]
        
        cols_hydroxide = ["Battery-Grade Lithium Hydroxide (Coarse Particles) Price",
                          "Battery-Grade Lithium Hydroxide (Coarse Particles) Price Range",
                          "Battery-Grade Lithium Hydroxide (Micro Powder) Price",
                          "Battery-Grade Lithium Hydroxide (Micro Powder) Price Range",
                          "Battery-Grade Lithium Hydroxide (CIF China Japan and South Korea) Price",
                          "Battery-Grade Lithium Hydroxide (CIF China Japan and South Korea) Price Range",
                          "SMM Battery-Grade Lithium Hydroxide Index Price",
                          "SMM Battery-Grade Lithium Hydroxide Index Price Range",
                          "Industrial-Grade Lithium Hydroxide Price",
                          "Industrial-Grade Lithium Hydroxide Price Range"]
        
        data_hydroxide = []
        
        for url in urls_hydroxide:
            price, range_price = await extract_price_data_playwright(page, url)
            data_hydroxide.append(price if price else "")
            data_hydroxide.append(range_price if range_price else "")
            await page.wait_for_timeout(3000)
        
        df_lithium_hydroxide = pd.DataFrame([data_hydroxide], columns=cols_hydroxide)
        
        # =========================
        # LITHIUM METAL
        # =========================
        print("\n--- Extrayendo Lithium Metal ---")
        urls_metal = ["https://www.metal.com/Lithium/202304250001",
                      "https://www.metal.com/Lithium/202304250002"]
        
        cols_metal = ["Industrial-Grade Lithium Metal (Weekly) Price",
                      "Industrial-Grade Lithium Metal (Weekly) Price Range",
                      "Battery-Grade Lithium Metal (Weekly) Price",
                      "Battery-Grade Lithium Metal (Weekly) Price Range"]
        
        data_metal = []
        
        for url in urls_metal:
            price, range_price = await extract_price_data_playwright(page, url)
            data_metal.append(price if price else "")
            data_metal.append(range_price if range_price else "")
            await page.wait_for_timeout(3000)
        
        df_lithium_metal = pd.DataFrame([data_metal], columns=cols_metal)
        
        # =========================
        # OTHER CHEMICALS
        # =========================
        print("\n--- Extrayendo Other Chemicals ---")
        urls_other = ["https://www.metal.com/Lithium/202110220001",
                      "https://www.metal.com/Lithium/202307040006"]
        
        cols_other = ["LiPF6 (Domestic) Price",
                      "LiPF6 (Domestic) Price Range",
                      "Battery-Grade Lithium Fluoride Price",
                      "Battery-Grade Lithium Fluoride Price Range"]
        
        data_other = []
        
        for url in urls_other:
            price, range_price = await extract_price_data_playwright(page, url)
            data_other.append(price if price else "")
            data_other.append(range_price if range_price else "")
            await page.wait_for_timeout(3000)
        
        df_other = pd.DataFrame([data_other], columns=cols_other)
        
        # =========================
        # RARE EARTH OXIDES
        # =========================
        print("\n--- Extrayendo Rare Earth Oxides ---")
        await page.goto("https://www.metal.com/Rare-Earth-Oxides", wait_until="networkidle")
        await page.wait_for_timeout(5000)
        
        # Esperar a que la tabla cargue
        await page.wait_for_selector(".ant-table-content table", timeout=10000)
        
        # Obtener HTML de la tabla
        table_html = await page.evaluate("""
            () => {
                const table = document.querySelector('.ant-table-content table');
                return table ? table.outerHTML : null;
            }
        """)
        
        if table_html:
            df_rare_earth = pd.read_html(StringIO(table_html))[0]
            df_rare_earth['Name'] = df_rare_earth['Name'].str.replace(r'SMM.*$', '', regex=True).str.strip()
            df_rare_earth = df_rare_earth.rename(columns={
                "Name": "Price_description",
                "Average": "Avg."
            })
        else:
            df_rare_earth = pd.DataFrame()
        
        # ============================================
        # VERIFICAR QUE SE EXTRAJERON DATOS
        # ============================================
        print("\n=== VERIFICANDO DATOS EXTRAÍDOS ===")
        
        def df_tiene_datos(df):
            if df.empty:
                return False
            for col in df.columns:
                if df[col].notna().any() and (df[col] != "").any() and (df[col] != "N/A").any():
                    return True
            return False
        
        tiene_datos = False
        
        if df_tiene_datos(df_lithium_carbonate):
            tiene_datos = True
            print("✅ Lithium Carbonate: Tiene datos")
        else:
            print("❌ Lithium Carbonate: Sin datos")
        
        if df_tiene_datos(df_lithium_hydroxide):
            tiene_datos = True
            print("✅ Lithium Hydroxide: Tiene datos")
        else:
            print("❌ Lithium Hydroxide: Sin datos")
        
        if df_tiene_datos(df_lithium_metal):
            tiene_datos = True
            print("✅ Lithium Metal: Tiene datos")
        else:
            print("❌ Lithium Metal: Sin datos")
        
        if df_tiene_datos(df_other):
            tiene_datos = True
            print("✅ Other: Tiene datos")
        else:
            print("❌ Other: Sin datos")
        
        if not tiene_datos:
            print("\n❌❌❌ NO SE EXTRAJERON DATOS - DETENIENDO EJECUCIÓN ❌❌❌")
            await browser.close()
            sys.exit(1)
        
        print("✅ Datos extraídos correctamente")
        
        # ============================================
        # RESULTADOS Y GUARDADO
        # ============================================
        print("\n=== RESUMEN DE DATOS ===")
        print(f"Lithium Carbonate: {len(df_lithium_carbonate)} registros")
        print(f"Lithium Hydroxide: {len(df_lithium_hydroxide)} registros")
        print(f"Lithium Metal: {len(df_lithium_metal)} registros")
        print(f"Other Chemicals: {len(df_other)} registros")
        print(f"Rare Earth Oxides: {len(df_rare_earth)} registros")
        print("========================")
        
        file_name = "Reporte_Diario.xlsx"
        
        engine = "xlsxwriter"
        try:
            __import__("xlsxwriter")
        except ImportError:
            engine = "openpyxl"
        
        with pd.ExcelWriter(file_name, engine=engine) as writer:
            
            df_lithium_carbonate.to_excel(writer, sheet_name="Lithium carbonate", index=False)
            df_lithium_hydroxide.to_excel(writer, sheet_name="Lithium hydroxide", index=False)
            df_lithium_metal.to_excel(writer, sheet_name="Lithium metal", index=False)
            df_other.to_excel(writer, sheet_name="Other", index=False)
            df_rare_earth.to_excel(writer, sheet_name="REO", index=False)
            
            workbook = writer.book
            
            dfs = [
                ("Lithium carbonate", df_lithium_carbonate, "LC_Data"),
                ("Lithium hydroxide", df_lithium_hydroxide, "LH_Data"),
                ("Lithium metal", df_lithium_metal, "LM_Data"),
                ("Other", df_other, "Other_Data"),
                ("REO", df_rare_earth, "REO_Data"),
            ]
            
            for sheet_name, df, table_name in dfs:
                worksheet = writer.sheets[sheet_name]
                (rows, cols) = df.shape
                column_settings = [{"header": col} for col in df.columns]
                worksheet.add_table(
                    0, 0, rows, cols-1,
                    {
                        "columns": column_settings,
                        "name": table_name
                    }
                )
        
        # =========================
        # ENVIAR EMAIL
        # =========================
        print("\n--- Enviando email...")
        sender = os.environ["EMAIL_USER"]
        password_email = os.environ["EMAIL_PASS"]
        receiver = "market.intelligence@JGI.be"
        
        msg = EmailMessage()
        
        msg["Subject"] = f"Price Tracking Data - {datetime.now().strftime('%d/%m/%Y')}"
        msg["From"] = sender
        msg["To"] = receiver
        
        msg.set_content("Daily report.")
        
        with open(file_name, "rb") as f:
            file_data = f.read()
            file_name = f.name
        
        msg.add_attachment(
            file_data,
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=file_name
        )
        
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(sender, password_email)
            smtp.send_message(msg)
        
        await browser.close()
        print("\n✅ Proceso completado exitosamente - Email enviado con datos")
        
    except Exception as e:
        print(f"❌ Error en el proceso: {e}")
        await browser.close()
        sys.exit(1)

# ============================================
# EJECUTAR
# ============================================
if __name__ == "__main__":
    asyncio.run(main())
