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
    """Realiza el login usando el botón real smm-auth-submit"""
    
    print("\n=== INICIANDO LOGIN CON PLAYWRIGHT ===")
    
    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=True,
            args=[
                '--no-sandbox',
                '--disable-dev-shm-usage',
                '--disable-blink-features=AutomationControlled',
                '--disable-gpu',
                '--disable-extensions',
                '--disable-setuid-sandbox',
                '--window-size=1920,1080'
            ]
        )
        
        context = await browser.new_context(
            user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            viewport={'width': 1920, 'height': 1080},
            locale='en-US',
            timezone_id='America/New_York'
        )
        
        page = await context.new_page()
        page.set_default_timeout(60000)
        
        print("Cargando página principal...")
        await page.goto("https://www.metal.com/", wait_until="networkidle")
        await page.wait_for_timeout(3000)
        
        # PASO 1: Hacer clic en Sign In
        print("Haciendo clic en Sign In...")
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
        print("✅ Clic en Sign In")
        await page.wait_for_timeout(3000)
        
        # PASO 2: Esperar a que el popup esté visible y los campos existan
        print("Esperando popup y campos...")
        
        # Esperar a que el contenedor del popup esté visible
        try:
            await page.wait_for_selector('#smm-auth-widget-root', state='visible', timeout=10000)
            print("✅ Popup visible")
        except:
            print("⚠️ Popup no visible, forzando visibilidad...")
            await page.evaluate("""
                () => {
                    const popup = document.querySelector('#smm-auth-widget-root');
                    if (popup) {
                        popup.style.display = 'block';
                        popup.style.visibility = 'visible';
                        popup.style.opacity = '1';
                    }
                }
            """)
            await page.wait_for_timeout(1000)
        
        # Verificar si los campos existen en el DOM normal
        campos_en_dom = await page.evaluate("""
            () => {
                const user = document.querySelector('#_r_0_');
                const pass = document.querySelector('#_r_2_');
                return {
                    user_exists: !!user,
                    pass_exists: !!pass
                };
            }
        """)
        
        print(f"Campos en DOM normal: {campos_en_dom}")
        
        # Si no existen en DOM normal, buscar en Shadow DOM
        if not campos_en_dom.get('user_exists') or not campos_en_dom.get('pass_exists'):
            print("Campos no encontrados en DOM normal, buscando en Shadow DOM...")
            shadow_campos = await page.evaluate("""
                () => {
                    const host = document.querySelector('#smm-auth-widget-root');
                    if (!host) return { user_exists: false, pass_exists: false };
                    if (!host.shadowRoot) return { user_exists: false, pass_exists: false };
                    const user = host.shadowRoot.querySelector('#_r_0_');
                    const pass = host.shadowRoot.querySelector('#_r_2_');
                    return {
                        user_exists: !!user,
                        pass_exists: !!pass
                    };
                }
            """)
            print(f"Campos en Shadow DOM: {shadow_campos}")
            
            # Si están en Shadow DOM, usar esa ruta para llenar
            if shadow_campos.get('user_exists') and shadow_campos.get('pass_exists'):
                print("Llenando campos desde Shadow DOM...")
                llenado_shadow = await page.evaluate(f"""
                    () => {{
                        const host = document.querySelector('#smm-auth-widget-root');
                        if (!host || !host.shadowRoot) return {{ success: false }};
                        const shadow = host.shadowRoot;
                        
                        function llenarInput(element, value) {{
                            if (!element) return false;
                            element.focus();
                            element.dispatchEvent(new Event('focus', {{ bubbles: true }}));
                            element.select();
                            element.value = value;
                            ['input', 'change', 'blur'].forEach(ev => {{
                                element.dispatchEvent(new Event(ev, {{ bubbles: true }}));
                            }});
                            const native = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
                            native.call(element, value);
                            element.dispatchEvent(new Event('input', {{ bubbles: true }}));
                            return true;
                        }}
                        
                        const userInput = shadow.querySelector('#_r_0_');
                        const passInput = shadow.querySelector('#_r_2_');
                        if (!userInput || !passInput) return {{ success: false }};
                        
                        const userOk = llenarInput(userInput, '{user}');
                        const passOk = llenarInput(passInput, '{password}');
                        
                        // Habilitar y hacer clic en el botón
                        const loginBtn = shadow.querySelector('button.smm-auth-submit');
                        if (loginBtn) {{
                            loginBtn.disabled = false;
                            loginBtn.removeAttribute('disabled');
                            loginBtn.click();
                        }}
                        
                        return {{
                            success: userOk && passOk,
                            user_value: userInput.value,
                            pass_value: passInput.value ? '****' : 'vacio'
                        }};
                    }}
                """)
                print(f"Resultado llenado Shadow DOM: {llenado_shadow}")
                if llenado_shadow.get('success'):
                    print("✅ Login enviado desde Shadow DOM")
                    await page.wait_for_timeout(5000)
                    # Verificar login
                    await page.goto("https://www.metal.com/", wait_until="networkidle")
                    await page.wait_for_timeout(5000)
                    if "Sign Out" in await page.content():
                        print("✅ LOGIN EXITOSO (Shadow DOM)")
                        return page, browser, context
                    else:
                        print("⚠️ Login no confirmado después de Shadow DOM")
                        await browser.close()
                        return False
                else:
                    print("❌ Error llenando campos desde Shadow DOM")
                    await browser.close()
                    return False
        
        # Si los campos están en DOM normal, continuar con el método anterior
        # (El código de llenado normal ya está arriba, pero lo pondré completo)
        print("Llenando campos desde DOM normal...")
        
        # Esperar a que los campos estén disponibles
        await page.wait_for_selector('#_r_0_', state='attached', timeout=10000)
        await page.wait_for_selector('#_r_2_', state='attached', timeout=10000)
        
        await page.screenshot(path="screenshot_antes_llenar.png")
        print("📸 Screenshot: screenshot_antes_llenar.png")
        
        llenar_campos = f"""
            (function() {{
                function llenarInput(element, value) {{
                    if (!element) return false;
                    element.focus();
                    element.dispatchEvent(new Event('focus', {{ bubbles: true }}));
                    element.select();
                    element.value = value;
                    ['input', 'change', 'blur'].forEach(ev => {{
                        element.dispatchEvent(new Event(ev, {{ bubbles: true }}));
                    }});
                    const native = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
                    native.call(element, value);
                    element.dispatchEvent(new Event('input', {{ bubbles: true }}));
                    return true;
                }}
                
                const userInput = document.querySelector('#_r_0_');
                const passInput = document.querySelector('#_r_2_');
                if (!userInput || !passInput) return {{ success: false }};
                
                const userOk = llenarInput(userInput, '{user}');
                const passOk = llenarInput(passInput, '{password}');
                
                return {{
                    success: userOk && passOk,
                    user_value: userInput.value,
                    pass_value: passInput.value ? '****' : 'vacio'
                }};
            }})();
        """
        
        resultado_llenado = await page.evaluate(llenar_campos)
        print(f"Resultado llenado: {resultado_llenado}")
        
        if not resultado_llenado.get('success'):
            print("❌ Error al llenar los campos")
            await browser.close()
            return False
        
        await page.screenshot(path="screenshot_campos_llenados.png")
        print("📸 Screenshot: screenshot_campos_llenados.png")
        
        # PASO 3: Enviar login (botón real)
        print("Enviando login...")
        click_result = await page.evaluate("""
            () => {
                const loginBtn = document.querySelector('button.smm-auth-submit');
                if (!loginBtn) return 'button_not_found';
                loginBtn.disabled = false;
                loginBtn.removeAttribute('disabled');
                loginBtn.removeAttribute('aria-busy');
                loginBtn.click();
                const event = new MouseEvent('click', { view: window, bubbles: true, cancelable: true });
                loginBtn.dispatchEvent(event);
                return 'click_sent';
            }
        """)
        print(f"Resultado envío: {click_result}")
        
        await page.wait_for_timeout(10000)
        await page.screenshot(path="screenshot_despues_envio.png")
        print("📸 Screenshot: screenshot_despues_envio.png")
        
        # PASO 4: Verificar login
        print("Verificando login...")
        await page.goto("https://www.metal.com/", wait_until="networkidle")
        await page.wait_for_timeout(5000)
        await page.screenshot(path="screenshot_post_login.png")
        print("📸 Screenshot: screenshot_post_login.png")
        
        if "Sign Out" in await page.content():
            print("✅ LOGIN EXITOSO - Usuario autenticado")
        else:
            cookies = await context.cookies()
            session_cookie = next((c for c in cookies if any(k in c['name'].lower() for k in ['session','auth','token','sid'])), None)
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
        
        numbers = re.findall(r'\d+[,.]?\d*', page_content)
        if len(numbers) > 10:
            print(f"✅ Se puede acceder a los datos ({len(numbers)} números encontrados)")
        else:
            print("⚠️ La página no pide login pero hay pocos números")
        
        print("\n✅ Login verificado - Continuando con scraping...")
        return page, browser, context

# ============================================
# FUNCIONES DE SCRAPING (sin cambios)
# ============================================

async def extract_price_data_playwright(page, url):
    """Extrae datos de precio usando Playwright"""
    try:
        print(f"\n🔍 Extrayendo datos de: {url}")
        
        await page.goto(url, wait_until="networkidle")
        await page.wait_for_timeout(5000)
        
        page_content = await page.content()
        
        if "Sign in to view" in page_content:
            print("  ❌ La página pide autenticación")
            return None, None
        
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
        
        first_price = None
        try:
            avg_element = await page.query_selector("div[class*='avg']")
            if avg_element:
                first_price = await avg_element.text_content()
                first_price = first_price.strip() if first_price else None
                print(f"  ✅ Precio promedio: {first_price}")
        except Exception as e:
            print(f"  ❌ Error extrayendo precio: {e}")
        
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
        
        await page.wait_for_selector(".ant-table-content table", timeout=10000)
        
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
        # VERIFICAR DATOS
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

if __name__ == "__main__":
    asyncio.run(main())
