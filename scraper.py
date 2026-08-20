from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
import pandas as pd
import time
import smtplib
from email.message import EmailMessage
import os
from io import StringIO
import undetected_chromedriver as uc
from selenium.webdriver.chrome.options import Options
import subprocess
import random
from datetime import datetime

def debug_login_page(driver):
    """Guarda información de la página de login para debugging"""
    print("=== DEBUG LOGIN PAGE ===")
    # Guardar HTML
    with open('login_page.html', 'w', encoding='utf-8') as f:
        f.write(driver.page_source)
    
    # Mostrar todos los inputs
    inputs = driver.find_elements(By.XPATH, "//input")
    print(f"Total inputs encontrados: {len(inputs)}")
    for i, inp in enumerate(inputs):
        try:
            print(f"Input {i}: type={inp.get_attribute('type')}, placeholder={inp.get_attribute('placeholder')}, id={inp.get_attribute('id')}, class={inp.get_attribute('class')}")
        except:
            pass
    
    # Mostrar todos los botones
    botones = driver.find_elements(By.XPATH, "//button")
    print(f"Total botones encontrados: {len(botones)}")
    for i, btn in enumerate(botones):
        try:
            print(f"Botón {i}: text={btn.text}, type={btn.get_attribute('type')}, class={btn.get_attribute('class')}")
        except:
            pass
    print("=== FIN DEBUG ===")

###----------------------------------------------------------------------> INICIO <----------------------------------------------------------------------###
user = os.environ["METAL_USER"]
password = os.environ["METAL_PASS"]

# Configuración optimizada para GitHub Actions
options = Options()

options.add_argument("--headless=new")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--disable-blink-features=AutomationControlled")

options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

chrome_version = subprocess.check_output(["google-chrome", "--version"]).decode()
chrome_version = int(chrome_version.split(" ")[2].split(".")[0])

driver = uc.Chrome(
    options=options,
    headless=True,
    version_main=chrome_version
)

service = Service()

driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")

time.sleep(random.uniform(1.5, 3.5))
driver.get('https://www.metal.com/')
time.sleep(random.uniform(8, 14))

# Sign in
# ============================================
# LOGIN - BUSCANDO EN IFRAMES CON SHADOW DOM
# ============================================

try:
    # Primero, hacer clic en Sign In
    print("Abriendo popup de login...")
    
    # Buscar el botón de Sign In de varias formas
    boton_signin = None
    try:
        boton_signin = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Sign In')]"))
        )
    except:
        try:
            boton_signin = driver.find_element(By.CSS_SELECTOR, "button.signInButton")
        except:
            pass
    
    if boton_signin:
        boton_signin.click()
        print("✅ Clic en Sign In ejecutado")
    else:
        # Usar JavaScript como respaldo
        driver.execute_script("""
            const btn = document.querySelector('button.signInButton');
            if (btn) btn.click();
        """)
        print("✅ Clic en Sign In con JavaScript")
    
    # Esperar a que el popup se abra
    time.sleep(5)
    print("Esperando carga del popup...")
    
    # ============================================
    # BUSCAR EN TODOS LOS IFRAMES
    # ============================================
    print("Buscando en iframes...")
    
    input_user = None
    input_pass = None
    boton_login = None
    iframe_encontrado = None
    
    # Volver al contexto principal
    driver.switch_to.default_content()
    
    # Encontrar todos los iframes
    iframes = driver.find_elements(By.TAG_NAME, "iframe")
    print(f"Iframes encontrados: {len(iframes)}")
    
    # Mostrar información de cada iframe
    for i, iframe in enumerate(iframes):
        try:
            src = iframe.get_attribute('src')
            print(f"Iframe {i}: src={src[:100] if src else 'No src'}")
        except:
            pass
    
    # Buscar en cada iframe
    for i in range(len(iframes)):
        try:
            print(f"\n--- Probando iframe {i} ---")
            driver.switch_to.default_content()
            
            # Cambiar al iframe
            iframe = driver.find_elements(By.TAG_NAME, "iframe")[i]
            driver.switch_to.frame(iframe)
            print(f"Cambiado al iframe {i}")
            
            # Buscar inputs en este iframe
            inputs = driver.find_elements(By.XPATH, "//input")
            print(f"Inputs en iframe {i}: {len(inputs)}")
            
            # Mostrar detalles de los inputs
            for inp in inputs:
                try:
                    inp_id = inp.get_attribute('id')
                    inp_type = inp.get_attribute('type')
                    inp_placeholder = inp.get_attribute('placeholder')
                    print(f"  Input: id={inp_id}, type={inp_type}, placeholder={inp_placeholder}")
                    
                    if inp_id == '_r_0_':
                        input_user = inp
                        iframe_encontrado = i
                        print(f"✅ Campo usuario encontrado en iframe {i}")
                    elif inp_id == '_r_2_':
                        input_pass = inp
                        print(f"✅ Campo contraseña encontrado en iframe {i}")
                except:
                    pass
            
            # Buscar botón
            if not boton_login:
                try:
                    botones = driver.find_elements(By.XPATH, "//button")
                    for btn in botones:
                        try:
                            btn_text = btn.text
                            btn_class = btn.get_attribute('class')
                            if 'smm-auth-submit' in btn_class or 'Sign in' in btn_text:
                                boton_login = btn
                                print(f"✅ Botón login encontrado en iframe {i}")
                                break
                        except:
                            pass
                except:
                    pass
            
            # Si encontramos todos los elementos, salir
            if input_user and input_pass and boton_login:
                print(f"✅ Todos los elementos encontrados en iframe {i}")
                break
                
        except Exception as e:
            print(f"Error en iframe {i}: {e}")
        finally:
            # Volver al contexto principal
            if not (input_user and input_pass and boton_login):
                try:
                    driver.switch_to.default_content()
                except:
                    pass
    
    # ============================================
    # SI NO SE ENCONTRÓ EN IFRAMES, BUSCAR CON JAVASCRIPT
    # ============================================
    if not input_user:
        print("\nNo se encontraron elementos en iframes. Intentando con JavaScript...")
        
        # Volver al contexto principal
        driver.switch_to.default_content()
        
        # Buscar con JavaScript en todos los iframes
        result = driver.execute_script("""
            function findInIframes() {
                const iframes = document.querySelectorAll('iframe');
                for (let iframe of iframes) {
                    try {
                        const doc = iframe.contentDocument || iframe.contentWindow.document;
                        if (doc) {
                            // Buscar inputs
                            const inputs = doc.querySelectorAll('input');
                            for (let inp of inputs) {
                                if (inp.id === '_r_0_') {
                                    return { type: 'user', element: inp, iframe: iframe };
                                }
                                if (inp.id === '_r_2_') {
                                    return { type: 'pass', element: inp, iframe: iframe };
                                }
                            }
                        }
                    } catch(e) {
                        console.log('Error accediendo a iframe:', e);
                    }
                }
                return null;
            }
            
            // Buscar en el DOM principal también
            const mainInputs = document.querySelectorAll('input');
            for (let inp of mainInputs) {
                if (inp.id === '_r_0_') {
                    return { type: 'user', element: inp, iframe: null };
                }
                if (inp.id === '_r_2_') {
                    return { type: 'pass', element: inp, iframe: null };
                }
            }
            
            return findInIframes();
        """)
        
        if result:
            print(f"JavaScript encontró elementos en iframe")
            # Recuperar los elementos encontrados
            try:
                # Buscar nuevamente usando Selenium después de saber dónde está
                iframes = driver.find_elements(By.TAG_NAME, "iframe")
                for iframe in iframes:
                    try:
                        driver.switch_to.frame(iframe)
                        inputs = driver.find_elements(By.XPATH, "//input")
                        for inp in inputs:
                            if inp.get_attribute('id') == '_r_0_':
                                input_user = inp
                                print("✅ Usuario encontrado después de JavaScript")
                            if inp.get_attribute('id') == '_r_2_':
                                input_pass = inp
                                print("✅ Contraseña encontrada después de JavaScript")
                        if input_user and input_pass:
                            # Buscar botón
                            botones = driver.find_elements(By.XPATH, "//button")
                            for btn in botones:
                                if 'smm-auth-submit' in btn.get_attribute('class'):
                                    boton_login = btn
                                    print("✅ Botón encontrado después de JavaScript")
                                    break
                            if boton_login:
                                break
                        driver.switch_to.default_content()
                    except:
                        driver.switch_to.default_content()
            except Exception as e:
                print(f"Error recuperando elementos: {e}")
    
    # ============================================
    # VERIFICAR FINAL
    # ============================================
    if not input_user:
        # Último intento: Buscar por texto en la página
        print("\nBuscando por texto en la página...")
        driver.switch_to.default_content()
        
        # Buscar cualquier elemento que contenga "Email"
        elementos_texto = driver.find_elements(By.XPATH, "//*[contains(text(), 'Email')]")
        print(f"Elementos con texto 'Email': {len(elementos_texto)}")
        
        # Si encontramos texto "Email", buscar el input cercano
        for elem in elementos_texto:
            try:
                # Buscar el siguiente input después del elemento
                input_cercano = elem.find_element(By.XPATH, "./following::input[1]")
                if input_cercano:
                    input_user = input_cercano
                    print("✅ Usuario encontrado por texto cercano")
                    break
            except:
                pass
        
        if not input_user:
            print("\n=== ERROR FINAL: No se encontraron los elementos ===")
            # Guardar información de debug
            with open('debug_final_completo.html', 'w', encoding='utf-8') as f:
                f.write(driver.page_source)
            driver.save_screenshot('debug_final_screenshot.png')
            
            # Mostrar todos los iframes y su contenido
            driver.switch_to.default_content()
            iframes = driver.find_elements(By.TAG_NAME, "iframe")
            print(f"\nTotal iframes: {len(iframes)}")
            for i, iframe in enumerate(iframes):
                try:
                    driver.switch_to.frame(iframe)
                    print(f"\nIframe {i}:")
                    print(f"  URL: {driver.current_url}")
                    print(f"  Title: {driver.title}")
                    print(f"  Inputs: {len(driver.find_elements(By.XPATH, '//input'))}")
                    driver.switch_to.default_content()
                except:
                    print(f"Iframe {i}: No se pudo acceder")
            
            raise Exception("No se pudo encontrar el campo de usuario después de todos los intentos")
    
    # ============================================
    # INGRESAR CREDENCIALES
    # ============================================
    print("\n--- Ingresando credenciales ---")
    
    # Asegurarse de estar en el iframe correcto si fue encontrado
    if iframe_encontrado is not None:
        driver.switch_to.default_content()
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        if iframe_encontrado < len(iframes):
            driver.switch_to.frame(iframes[iframe_encontrado])
            print(f"Cambiado al iframe {iframe_encontrado} para ingresar datos")
    
    # Limpiar y escribir
    try:
        input_user.clear()
        input_user.send_keys(user)
        print("✅ Usuario ingresado")
    except:
        # Si falla, usar JavaScript
        driver.execute_script(f"arguments[0].value = '{user}';", input_user)
        print("✅ Usuario ingresado con JavaScript")
    
    try:
        input_pass.clear()
        input_pass.send_keys(password)
        print("✅ Contraseña ingresada")
    except:
        driver.execute_script(f"arguments[0].value = '{password}';", input_pass)
        print("✅ Contraseña ingresada con JavaScript")
    
    # Buscar y hacer clic en el botón
    if boton_login:
        try:
            # Habilitar el botón si está deshabilitado
            driver.execute_script("arguments[0].disabled = false;", boton_login)
            boton_login.click()
            print("✅ Login enviado")
        except:
            # Si falla, usar JavaScript
            driver.execute_script("arguments[0].click();", boton_login)
            print("✅ Login enviado con JavaScript")
    
    # Esperar a que cargue
    time.sleep(5)
    
    # Verificar si el login fue exitoso
    try:
        driver.switch_to.default_content()
        # Buscar un elemento que indique que estamos logueados
        elementos_logout = driver.find_elements(By.XPATH, "//*[contains(text(), 'Sign Out')]")
        if elementos_logout:
            print("✅ Login exitoso - usuario logueado")
        else:
            print("⚠️ No se pudo confirmar login exitoso, continuando...")
    except:
        pass
    
    print("✅ Proceso de login completado")
    
except Exception as e:
    print(f"Error en el proceso de login: {str(e)}")
    try:
        driver.save_screenshot("error_login_completo.png")
        with open('error_html_completo.html', 'w', encoding='utf-8') as f:
            f.write(driver.page_source)
        print("Archivos de debug guardados")
    except:
        pass
    raise

time.sleep(random.uniform(8, 14))

wait = WebDriverWait(driver,10)

# =========================
# FUNCIONES
# =========================

def page_not_found(driver):
    try:
        driver.find_element(By.XPATH,'//div[contains(@class,"PriceWrap")]')
        return False
    except NoSuchElementException:
        return True

def extract_price_data(driver, url):
    driver.get(url)
    time.sleep(3)

    if page_not_found(driver):
        return None, None
    
    container = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH,'//div[contains(@class,"__PriceWrap")]')))

    first_price = container.find_element(By.XPATH,'.//div[contains(@class,"avg")]').text

    high = None
    low = None

    try:
        high = container.find_element(By.XPATH,'.//div[contains(@class,"list")]/div[1]/label[2]').text
    except:
        pass

    try:
        low = container.find_element(By.XPATH,'.//div[contains(@class,"list")]/div[2]/label[2]').text
    except:
        pass

    time.sleep(3)

    if low is not None and high is not None:
        price_range = f"{low}-{high}"
    else:
        price_range = None

    return first_price, price_range

# =========================
# LITHIUM CARBONATE
# =========================

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
    price, range_price = extract_price_data(driver,url)
    data_carbonate.append(price)
    data_carbonate.append(range_price)

df_lithium_carbonate = pd.DataFrame([data_carbonate], columns=cols_carbonate)

# =========================
# LITHIUM HYDROXIDE
# =========================

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
    price, range_price = extract_price_data(driver,url)
    data_hydroxide.append(price)
    data_hydroxide.append(range_price)

df_lithium_hydroxide = pd.DataFrame([data_hydroxide], columns=cols_hydroxide)

# =========================
# LITHIUM METAL
# =========================

urls_metal = ["https://www.metal.com/Lithium/202304250001",
              "https://www.metal.com/Lithium/202304250002"]

cols_metal = ["Industrial-Grade Lithium Metal (Weekly) Price",
              "Industrial-Grade Lithium Metal (Weekly) Price Range",
              "Battery-Grade Lithium Metal (Weekly) Price",
              "Battery-Grade Lithium Metal (Weekly) Price Range"]

data_metal = []

for url in urls_metal:
    price, range_price = extract_price_data(driver,url)
    data_metal.append(price)
    data_metal.append(range_price)

df_lithium_metal = pd.DataFrame([data_metal], columns=cols_metal)

# =========================
# OTHER CHEMICALS
# =========================

urls_other = ["https://www.metal.com/Lithium/202110220001",
              "https://www.metal.com/Lithium/202307040006"]

cols_other = ["LiPF6 (Domestic) Price",
              "LiPF6 (Domestic) Price Range",
              "Battery-Grade Lithium Fluoride Price",
              "Battery-Grade Lithium Fluoride Price Range"]

data_other = []

for url in urls_other:
    price, range_price = extract_price_data(driver,url)
    data_other.append(price)
    data_other.append(range_price)

df_other = pd.DataFrame([data_other], columns=cols_other)

del (cols_carbonate, cols_hydroxide, cols_metal, cols_other, data_carbonate, data_hydroxide, data_metal, data_other, price, range_price, url, urls_carbonate, urls_hydroxide, urls_metal, urls_other)

# =========================
# RARE EARTH OXIDES
# =========================

driver.get("https://www.metal.com/Rare-Earth-Oxides")
wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,".ant-table-content table")))
table = driver.find_element(By.CSS_SELECTOR,".ant-table-content table")
df_rare_earth = pd.read_html(StringIO(table.get_attribute("outerHTML")))[0]
df_rare_earth['Name'] = df_rare_earth['Name'].str.replace(r'SMM.*$', '', regex=True).str.strip()
df_rare_earth = df_rare_earth.rename(columns={
    "Name": "Price_description",
    "Average": "Avg."
})

driver.quit()

file_name = "Reporte_Diario.xlsx"

engine = "xlsxwriter"
try:
    __import__("xlsxwriter")
except ImportError:
    engine = "openpyxl"

engine = "xlsxwriter"

with pd.ExcelWriter(file_name, engine=engine) as writer:

    df_lithium_carbonate.to_excel(writer, sheet_name="Lithium carbonate", index=False)
    df_lithium_hydroxide.to_excel(writer, sheet_name="Lithium hydroxide", index=False)
    df_lithium_metal.to_excel(writer, sheet_name="Lithium metal", index=False)
    df_other.to_excel(writer, sheet_name="Other", index=False)
    df_rare_earth.to_excel(writer, sheet_name="REO", index=False)

    workbook  = writer.book

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

sender = os.environ["EMAIL_USER"]
password = os.environ["EMAIL_PASS"]
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
    smtp.login(sender, password)
    smtp.send_message(msg)
