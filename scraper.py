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
    """Realiza el login usando Playwright - Versión con JavaScript para llenar campos"""
    
    print("\n=== INICIANDO LOGIN CON PLAYWRIGHT ===")
    
    async with async_playwright() as p:
        # Lanzar navegador en modo headless
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
        
        # Crear contexto
        context = await browser.new_context(
            user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            viewport={'width': 1920, 'height': 1080},
            locale='en-US',
            timezone_id='America/New_York'
        )
        
        # Crear página
        page = await context.new_page()
        page.set_default_timeout(60000)
        
        print("Cargando página principal...")
        await page.goto("https://www.metal.com/", wait_until="networkidle")
        await page.wait_for_timeout(3000)
        
        # PASO 1: Hacer clic en Sign In con JavaScript
        print("Buscando botón Sign In...")
        
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
        print("✅ Clic en Sign In (JavaScript)")
        await page.wait_for_timeout(3000)
        
        # PASO 2: Crear el popup manualmente si no existe (con los inputs correctos)
        print("Verificando popup...")
        
        await page.evaluate("""
            () => {
                // Verificar si el popup existe
                let popup = document.querySelector('#smm-auth-widget-root');
                if (!popup) {
                    console.log('Creando popup manualmente...');
                    
                    // Crear el contenedor del popup
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
                    
                    // Crear el modal
                    const modal = document.createElement('div');
                    modal.style.cssText = `
                        background: white;
                        padding: 40px;
                        border-radius: 8px;
                        width: 420px;
                        max-width: 90%;
                        box-shadow: 0 4px 20px rgba(0,0,0,0.3);
                        position: relative;
                    `;
                    
                    modal.innerHTML = `
                        <button class="iconClose" style="position:absolute;top:10px;right:15px;background:none;border:none;font-size:24px;cursor:pointer;color:#999;">×</button>
                        <h2 style="margin-top:0;text-align:center;color:#333;font-size:20px;">Welcome to SMM</h2>
                        <div style="margin-bottom:15px;">
                            <label style="display:block;margin-bottom:5px;color:#555;font-size:14px;">Email address or phone number</label>
                            <input id="_r_0_" type="text" placeholder="Email or phone" style="width:100%;padding:10px;border:1px solid #d9d9d9;border-radius:4px;font-size:14px;box-sizing:border-box;">
                        </div>
                        <div style="margin-bottom:20px;">
                            <label style="display:block;margin-bottom:5px;color:#555;font-size:14px;">Password</label>
                            <input id="_r_2_" type="password" placeholder="Password" style="width:100%;padding:10px;border:1px solid #d9d9d9;border-radius:4px;font-size:14px;box-sizing:border-box;">
                        </div>
                        <div style="margin-bottom:15px;">
                            <label style="color:#555;font-size:14px;">
                                <input type="checkbox" style="margin-right:8px;"> Remember me
                            </label>
                        </div>
                        <button id="login-submit" class="smm-auth-submit" style="width:100%;padding:12px;background:#d7000f;color:white;border:none;border-radius:4px;font-size:16px;font-weight:600;cursor:pointer;">Sign in</button>
                        <div style="margin-top:15px;text-align:center;color:#555;font-size:14px;">
                            Don't have an account? <a href="#" style="color:#d7000f;text-decoration:none;">Sign up</a>
                        </div>
                    `;
                    
                    container.appendChild(modal);
                    document.body.appendChild(container);
                    
                    // Agregar evento para cerrar el popup
                    const closeBtn = modal.querySelector('.iconClose');
                    if (closeBtn) {
                        closeBtn.addEventListener('click', () => {
                            container.remove();
                        });
                    }
                    
                    console.log('Popup creado manualmente');
                }
            }
        """)
        
        await page.wait_for_timeout(2000)
        
        # PASO 3: LLENAR LOS CAMPOS CON JAVASCRIPT (NO con fill())
        print("Llenando campos con JavaScript...")
        
        # Usar JavaScript para establecer los valores y disparar eventos
        await page.evaluate(f"""
            () => {{
                function setReactValue(element, value) {{
                    // Establecer el valor
                    element.value = value;
                    
                    // Disparar eventos en el orden correcto para React
                    element.dispatchEvent(new Event('focus', {{ bubbles: true }}));
                    element.dispatchEvent(new Event('input', {{ bubbles: true }}));
                    element.dispatchEvent(new Event('change', {{ bubbles: true }}));
                    element.dispatchEvent(new Event('blur', {{ bubbles: true }}));
                    
                    // También intentar con eventos adicionales
                    const event = new Event('input', {{ bubbles: true }});
                    const tracker = element._valueTracker;
                    if (tracker) {{
                        tracker.setValue(value);
                    }}
                    
                    console.log('Valor establecido:', value);
                }}
                
                // Buscar los inputs
                const userInput = document.querySelector('#_r_0_');
                const passInput = document.querySelector('#_r_2_');
                
                if (userInput && passInput) {{
                    // Limpiar primero
                    userInput.value = '';
                    passInput.value = '';
                    
                    // Establecer valores
                    setReactValue(userInput, '{user}');
                    setReactValue(passInput, '{password}');
                    
                    console.log('Campos llenados exitosamente');
                    
                    // Verificar que se llenaron
                    console.log('Usuario:', userInput.value);
                    console.log('Contraseña:', passInput.value ? '****' : 'vacío');
                    
                    // Habilitar el botón
                    const loginBtn = document.querySelector('#login-submit');
                    if (loginBtn) {{
                        loginBtn.disabled = false;
                        loginBtn.removeAttribute('disabled');
                        loginBtn.style.opacity = '1';
                        loginBtn.style.cursor = 'pointer';
                        console.log('Botón habilitado');
                    }}
                    
                    return true;
                }}
                return false;
            }}
        """)
        
        print("✅ Campos llenados con JavaScript")
        await page.wait_for_timeout(2000)
        
        # Tomar screenshot para verificar que los campos estén llenos
        await page.screenshot(path="screenshot_verificar_campos_llenados.png")
        print("📸 Screenshot: screenshot_verificar_campos_llenados.png")
        
        # PASO 4: Verificar que los campos se llenaron correctamente
        campos_verificados = await page.evaluate("""
            () => {
                const userInput = document.querySelector('#_r_0_');
                const passInput = document.querySelector('#_r_2_');
                return {
                    user_value: userInput ? userInput.value : 'no_encontrado',
                    pass_value: passInput ? (passInput.value ? '****' : 'vacio') : 'no_encontrado'
                };
            }
        """)
        
        print(f"Verificación de campos: {campos_verificados}")
        
        if campos_verificados.get('user_value') != user:
            print("⚠️ El campo de usuario no se llenó correctamente. Intentando nuevamente...")
            
            # Intentar nuevamente con un enfoque diferente
            await page.evaluate(f"""
                () => {{
                    const userInput = document.querySelector('#_r_0_');
                    if (userInput) {{
                        userInput.focus();
                        userInput.value = '{user}';
                        userInput.dispatchEvent(new Event('input', {{ bubbles: true }}));
                        userInput.dispatchEvent(new Event('change', {{ bubbles: true }}));
                        console.log('Usuario re-llenado:', userInput.value);
                    }}
                }}
            """)
            await page.wait_for_timeout(1000)
        
        # PASO 5: Enviar el login
        print("Enviando login...")
        
        login_result = await page.evaluate("""
            () => {
                // Buscar el botón de login
                let loginBtn = document.querySelector('#login-submit');
                if (!loginBtn) {
                    loginBtn = document.querySelector('button.smm-auth-submit');
                }
                
                if (loginBtn) {
                    // Habilitar el botón
                    loginBtn.disabled = false;
                    loginBtn.removeAttribute('disabled');
                    
                    // Hacer clic
                    loginBtn.click();
                    console.log('Login enviado');
                    return 'login_sent';
                }
                return 'button_not_found';
            }
        """)
        
        print(f"Resultado del login: {login_result}")
        await page.wait_for_timeout(5000)
        
        # Tomar screenshot después del login
        await page.screenshot(path="screenshot_despues_login.png")
        print("📸 Screenshot: screenshot_despues_login.png")
        
        # PASO 6: Esperar procesamiento y verificar
        print("Esperando procesamiento del login...")
        await page.wait_for_timeout(10000)
        
        # PASO 7: Verificar login
        print("Verificando login...")
        
        # Recargar la página principal
        await page.goto("https://www.metal.com/", wait_until="networkidle")
        await page.wait_for_timeout(5000)
        await page.screenshot(path="screenshot_post_login.png")
        print("📸 Screenshot: screenshot_post_login.png")
        
        # Verificar si hay elementos de usuario logueado
        page_content = await page.content()
        
        if "Sign Out" in page_content or "Logout" in page_content:
            print("✅ LOGIN EXITOSO - Usuario autenticado")
        else:
            print("⚠️ No se encontró 'Sign Out' en la página principal")
            
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
        
        # PASO 8: Verificar acceso a datos
        print("\n=== VERIFICANDO ACCESO A DATOS ===")
        
        test_url = "https://www.metal.com/Lithium/201102250059"
        await page.goto(test_url, wait_until="networkidle")
        await page.wait_for_timeout(5000)
        
        page_content = await page.content()
        
        if "Sign in to view" in page_content:
            print("❌ No se puede acceder a los datos - Pide autenticación")
            
            # Mostrar cookies para debug
            print("\n=== COOKIES ===")
            cookies = await context.cookies()
            for cookie in cookies:
                print(f"  {cookie.get('name')}: {cookie.get('value')[:50]}...")
            
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
