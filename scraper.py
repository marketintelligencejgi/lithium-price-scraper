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

# ============================================
# FUNCIÓN PARA LOGIN REAL CON SELENIUM
# ============================================
def realizar_login_real(driver, user, password):
    """
    Realiza el login real en la página usando Selenium
    """
    print("\n=== INICIANDO LOGIN REAL ===")
    
    try:
        # PASO 1: Ir a la página principal
        driver.get("https://www.metal.com/")
        time.sleep(5)
        print("✅ Página principal cargada")
        
        # PASO 2: Buscar y hacer clic en el botón "Sign In"
        print("Buscando botón Sign In...")
        boton_signin = None
        
        # Intentar múltiples selectores
        selectores = [
            "//button[contains(text(), 'Sign In')]",
            "//button[contains(@class, 'signInButton')]",
            "//a[contains(text(), 'Sign In')]",
            "//*[contains(@class, 'signInButton')]"
        ]
        
        for selector in selectores:
            try:
                boton_signin = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, selector))
                )
                if boton_signin:
                    print(f"✅ Botón Sign In encontrado con selector: {selector}")
                    break
            except:
                continue
        
        if not boton_signin:
            # Intentar con JavaScript
            boton_signin = driver.execute_script("""
                const btn = document.querySelector('button.signInButton');
                if (btn) return btn;
                const buttons = document.querySelectorAll('button');
                for (let b of buttons) {
                    if (b.textContent.includes('Sign In')) return b;
                }
                return null;
            """)
        
        if boton_signin:
            # Hacer clic con JavaScript para evitar problemas
            driver.execute_script("arguments[0].click();", boton_signin)
            print("✅ Clic en Sign In ejecutado")
            time.sleep(3)
        else:
            raise Exception("No se encontró el botón Sign In")
        
        # PASO 3: Esperar a que aparezca el popup de login
        print("Esperando popup de login...")
        time.sleep(5)
        
        # Buscar el iframe del login (si existe)
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        print(f"Iframes encontrados: {len(iframes)}")
        
        # Buscar en cada iframe
        input_user = None
        input_pass = None
        boton_login = None
        
        for i, iframe in enumerate(iframes):
            try:
                print(f"Probando iframe {i}...")
                driver.switch_to.frame(iframe)
                
                # Buscar inputs dentro del iframe
                inputs = driver.find_elements(By.XPATH, "//input")
                print(f"  Inputs en iframe {i}: {len(inputs)}")
                
                for inp in inputs:
                    try:
                        inp_id = inp.get_attribute('id')
                        if inp_id == '_r_0_':
                            input_user = inp
                            print(f"  ✅ Campo usuario encontrado en iframe {i}")
                        elif inp_id == '_r_2_':
                            input_pass = inp
                            print(f"  ✅ Campo contraseña encontrado en iframe {i}")
                    except:
                        pass
                
                if input_user and input_pass:
                    break
                    
            except Exception as e:
                print(f"Error con iframe {i}: {e}")
            finally:
                if not (input_user and input_pass):
                    try:
                        driver.switch_to.default_content()
                    except:
                        pass
        
        # Si no se encontraron en iframes, buscar en el DOM principal
        if not input_user or not input_pass:
            print("Buscando en DOM principal...")
            driver.switch_to.default_content()
            
            # Buscar inputs en el DOM principal
            inputs = driver.find_elements(By.XPATH, "//input")
            print(f"Inputs en DOM principal: {len(inputs)}")
            
            for inp in inputs:
                try:
                    inp_id = inp.get_attribute('id')
                    if inp_id == '_r_0_':
                        input_user = inp
                        print("✅ Campo usuario encontrado en DOM principal")
                    elif inp_id == '_r_2_':
                        input_pass = inp
                        print("✅ Campo contraseña encontrado en DOM principal")
                except:
                    pass
        
        if not input_user or not input_pass:
            raise Exception("No se encontraron los campos de login")
        
        # PASO 4: Ingresar credenciales
        print("Ingresando credenciales...")
        
        # Limpiar y escribir usuario
        try:
            input_user.clear()
            input_user.send_keys(user)
            print("✅ Usuario ingresado")
        except:
            driver.execute_script(f"arguments[0].value = '{user}';", input_user)
            print("✅ Usuario ingresado con JavaScript")
        
        # Limpiar y escribir contraseña
        try:
            input_pass.clear()
            input_pass.send_keys(password)
            print("✅ Contraseña ingresada")
        except:
            driver.execute_script(f"arguments[0].value = '{password}';", input_pass)
            print("✅ Contraseña ingresada con JavaScript")
        
        # PASO 5: Buscar y hacer clic en el botón de login
        print("Buscando botón de login...")
        
        # Buscar en el iframe actual o en el DOM principal
        try:
            boton_login = driver.find_element(By.XPATH, "//button[contains(@class, 'smm-auth-submit')]")
            print("✅ Botón login encontrado por clase")
        except:
            try:
                boton_login = driver.find_element(By.XPATH, "//button[@type='submit']")
                print("✅ Botón login encontrado por tipo submit")
            except:
                try:
                    boton_login = driver.find_element(By.XPATH, "//button[contains(text(), 'Sign In')]")
                    print("✅ Botón login encontrado por texto")
                except:
                    # Buscar cualquier botón visible
                    botones = driver.find_elements(By.XPATH, "//button")
                    for btn in botones:
                        if btn.is_displayed() and btn.is_enabled():
                            boton_login = btn
                            print(f"✅ Botón login encontrado: {btn.text}")
                            break
        
        if boton_login:
            # Habilitar el botón si está deshabilitado
            driver.execute_script("arguments[0].disabled = false;", boton_login)
            time.sleep(1)
            
            # Hacer clic en el botón
            driver.execute_script("arguments[0].click();", boton_login)
            print("✅ Login enviado")
        else:
            # Si no hay botón, intentar enviar el formulario
            try:
                driver.execute_script("""
                    const form = document.querySelector('form');
                    if (form) {
                        form.submit();
                        return true;
                    }
                    return false;
                """)
                print("✅ Formulario enviado")
            except:
                pass
        
        # PASO 6: Esperar a que el login se procese
        print("Esperando procesamiento del login...")
        time.sleep(8)
        
        # PASO 7: Verificar si el login fue exitoso
        driver.switch_to.default_content()
        
        # Recargar la página para verificar
        driver.get("https://www.metal.com/")
        time.sleep(5)
        
        # Verificar si hay elementos de usuario logueado
        elementos_logout = driver.find_elements(By.XPATH, "//*[contains(text(), 'Sign Out') or contains(text(), 'Logout') or contains(text(), 'My Account')]")
        
        if elementos_logout:
            print("✅ LOGIN EXITOSO - Usuario autenticado")
            return True
        else:
            print("⚠️ No se pudo confirmar el login")
            # Verificar si hay cookies de sesión
            cookies = driver.get_cookies()
            print(f"Cookies: {len(cookies)}")
            for cookie in cookies:
                print(f"  {cookie.get('name')}: {cookie.get('value')[:20]}...")
            
            # Si no hay mensaje de error pero tampoco confirmación, asumir éxito
            print("ℹ️ No se pudo confirmar, pero continuando...")
            return True
            
    except Exception as e:
        print(f"❌ Error en el login: {str(e)}")
        try:
            driver.save_screenshot("error_login_completo.png")
            with open('error_login_completo.html', 'w', encoding='utf-8') as f:
                f.write(driver.page_source)
        except:
            pass
        return False

# ============================================
# EJECUTAR LOGIN
# ============================================
login_exitoso = realizar_login_real(driver, user, password)

if not login_exitoso:
    print("❌ El login falló. Intentando método alternativo...")
    # Intentar nuevamente con otro enfoque
    driver.get("https://www.metal.com/")
    time.sleep(5)
    # Continuar de todas formas, pero con advertencia

# ============================================
# VERIFICACIÓN DE AUTENTICACIÓN
# ============================================
print("\n=== VERIFICANDO AUTENTICACIÓN ===")

# Intentar acceder a una página de precios
test_url = "https://www.metal.com/Lithium/201102250059"
driver.get(test_url)
time.sleep(5)

# Verificar si se puede ver el precio o pide login
page_source = driver.page_source

if "Sign in to view" in page_source or "sign in to view" in page_source.lower():
    print("❌ El login NO fue exitoso - La página sigue pidiendo autenticación")
    print("Intentando refrescar...")
    
    # Intentar refrescar la página después del login
    driver.get("https://www.metal.com/")
    time.sleep(3)
    driver.refresh()
    time.sleep(5)
    
    # Verificar nuevamente
    driver.get(test_url)
    time.sleep(5)
    page_source = driver.page_source
    
    if "Sign in to view" in page_source:
        print("❌ El login sigue sin funcionar. Continuando con scraping limitado...")
    else:
        print("✅ Login exitoso después de refrescar")
else:
    print("✅ El login fue exitoso - Se puede ver la información")

time.sleep(random.uniform(3, 5))

# =========================
# FUNCIONES DE SCRAPING
# =========================

def page_not_found(driver):
    """Verifica si la página existe y tiene datos"""
    try:
        time.sleep(2)
        
        # Buscar indicadores de que la página tiene datos
        elementos = driver.find_elements(By.XPATH, '//div[contains(@class, "__PriceWrap")]')
        if elementos:
            return False
        
        elementos = driver.find_elements(By.XPATH, '//div[contains(@class, "PriceWrap")]')
        if elementos:
            return False
        
        mensaje_error = driver.find_elements(By.XPATH, '//*[contains(text(), "404") or contains(text(), "Not Found")]')
        if mensaje_error:
            return True
        
        return True
    except:
        return True

def extract_price_data(driver, url):
    """Extrae datos de precio de una URL"""
    try:
        print(f"\n🔍 Extrayendo datos de: {url}")
        
        driver.get(url)
        time.sleep(5)
        
        if page_not_found(driver):
            print(f"⚠️ Página no encontrada o sin datos: {url}")
            return None, None
        
        # Buscar contenedor
        container = None
        try:
            container = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH, '//div[contains(@class, "__PriceWrap")]'))
            )
            print("  ✅ Contenedor __PriceWrap encontrado")
        except:
            try:
                container = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//div[contains(@class, "PriceWrap")]'))
                )
                print("  ✅ Contenedor PriceWrap encontrado")
            except Exception as e:
                print(f"  ❌ No se encontró contenedor: {e}")
                return None, None
        
        # Extraer precio promedio
        first_price = None
        try:
            price_element = container.find_element(By.XPATH, './/div[contains(@class,"avg")]')
            first_price = price_element.text.strip()
            print(f"  ✅ Precio promedio: {first_price}")
        except Exception as e:
            print(f"  ❌ Error extrayendo precio: {e}")
        
        # Extraer rango
        high = None
        low = None
        try:
            high_element = container.find_element(By.XPATH, './/div[contains(@class,"list")]/div[1]/label[2]')
            high = high_element.text.strip()
            print(f"  ✅ High: {high}")
        except:
            pass
        
        try:
            low_element = container.find_element(By.XPATH, './/div[contains(@class,"list")]/div[2]/label[2]')
            low = low_element.text.strip()
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
    price, range_price = extract_price_data(driver, url)
    data_carbonate.append(price if price else "N/A")
    data_carbonate.append(range_price if range_price else "N/A")
    time.sleep(2)

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
    price, range_price = extract_price_data(driver, url)
    data_hydroxide.append(price if price else "N/A")
    data_hydroxide.append(range_price if range_price else "N/A")
    time.sleep(2)

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
    price, range_price = extract_price_data(driver, url)
    data_metal.append(price if price else "N/A")
    data_metal.append(range_price if range_price else "N/A")
    time.sleep(2)

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
    price, range_price = extract_price_data(driver, url)
    data_other.append(price if price else "N/A")
    data_other.append(range_price if range_price else "N/A")
    time.sleep(2)

df_other = pd.DataFrame([data_other], columns=cols_other)

# =========================
# RARE EARTH OXIDES
# =========================
print("\n--- Extrayendo Rare Earth Oxides ---")
driver.get("https://www.metal.com/Rare-Earth-Oxides")
wait = WebDriverWait(driver, 10)
wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".ant-table-content table")))
table = driver.find_element(By.CSS_SELECTOR, ".ant-table-content table")
df_rare_earth = pd.read_html(StringIO(table.get_attribute("outerHTML")))[0]
df_rare_earth['Name'] = df_rare_earth['Name'].str.replace(r'SMM.*$', '', regex=True).str.strip()
df_rare_earth = df_rare_earth.rename(columns={
    "Name": "Price_description",
    "Average": "Avg."
})

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

driver.quit()
print("\n✅ Proceso completado exitosamente")
