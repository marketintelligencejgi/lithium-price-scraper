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
import sys
import requests
from bs4 import BeautifulSoup

###----------------------------------------------------------------------> INICIO <----------------------------------------------------------------------###

user = os.environ["METAL_USER"]
password = os.environ["METAL_PASS"]

# Configuración optimizada para GitHub Actions
options = Options()

options.add_argument("--headless=new")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_argument("--disable-gpu")
options.add_argument("--window-size=1920,1080")

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
# FUNCIÓN PARA LOGIN - ENFOQUE DEFINITIVO
# ============================================
def realizar_login_definitivo(driver, user, password):
    """
    Realiza el login accediendo directamente al Shadow DOM
    """
    print("\n=== INICIANDO LOGIN DEFINITIVO ===")
    
    try:
        # PASO 1: Hacer clic en Sign In
        print("Buscando botón Sign In...")
        boton_signin = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Sign In')]"))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", boton_signin)
        time.sleep(0.5)
        driver.execute_script("arguments[0].click();", boton_signin)
        print("✅ Clic en Sign In")
        time.sleep(3)
        
        # PASO 2: Acceder al Shadow DOM correctamente
        print("Accediendo al Shadow DOM...")
        
        # Verificar si el Shadow DOM existe
        shadow_host_exists = driver.execute_script("""
            return document.querySelector('#smm-auth-widget-root') !== null;
        """)
        
        if not shadow_host_exists:
            print("❌ No se encontró el contenedor del Shadow DOM")
            return False
        
        print("✅ Contenedor Shadow DOM encontrado")
        
        # PASO 3: Interactuar con el Shadow DOM usando JavaScript
        print("Ingresando credenciales en Shadow DOM...")
        
        login_script = f"""
        (function() {{
            // Obtener el host del Shadow DOM
            const host = document.querySelector('#smm-auth-widget-root');
            if (!host) {{
                console.log('No se encontró el host');
                return 'host_not_found';
            }}
            
            // Obtener el Shadow Root
            const shadowRoot = host.shadowRoot;
            if (!shadowRoot) {{
                console.log('No se encontró el Shadow Root');
                return 'shadow_root_not_found';
            }}
            
            console.log('Shadow Root encontrado');
            
            // Buscar los inputs
            const userInput = shadowRoot.querySelector('#_r_0_');
            const passInput = shadowRoot.querySelector('#_r_2_');
            const loginBtn = shadowRoot.querySelector('button.smm-auth-submit');
            
            if (!userInput || !passInput) {{
                console.log('Inputs no encontrados');
                return 'inputs_not_found';
            }}
            
            console.log('Inputs encontrados');
            
            // Limpiar e ingresar valores
            userInput.value = '';
            passInput.value = '';
            
            // Ingresar valores
            userInput.value = '{user}';
            passInput.value = '{password}';
            
            // Disparar eventos para React
            userInput.dispatchEvent(new Event('input', {{ bubbles: true }}));
            userInput.dispatchEvent(new Event('change', {{ bubbles: true }}));
            passInput.dispatchEvent(new Event('input', {{ bubbles: true }}));
            passInput.dispatchEvent(new Event('change', {{ bubbles: true }}));
            
            console.log('Valores ingresados');
            
            // Buscar y habilitar el botón
            if (loginBtn) {{
                // Habilitar el botón si está deshabilitado
                loginBtn.disabled = false;
                loginBtn.removeAttribute('disabled');
                console.log('Botón habilitado');
                
                // Hacer clic en el botón
                loginBtn.click();
                console.log('Login enviado');
                return 'login_sent';
            }} else {{
                console.log('Botón no encontrado');
                return 'button_not_found';
            }}
        }})();
        """
        
        result = driver.execute_script(login_script)
        print(f"Resultado del login: {result}")
        
        if result == 'login_sent':
            print("✅ Login enviado correctamente")
        else:
            print(f"⚠️ Resultado inesperado: {result}")
        
        # PASO 4: Esperar a que el login se procese
        print("Esperando procesamiento del login...")
        time.sleep(10)
        
        # PASO 5: Verificar el login
        print("Verificando login...")
        
        # Recargar la página principal
        driver.get("https://www.metal.com/")
        time.sleep(5)
        
        # Verificar si hay elementos de usuario logueado
        elementos_logout = driver.find_elements(By.XPATH, "//*[contains(text(), 'Sign Out') or contains(text(), 'Logout') or contains(text(), 'My Account')]")
        
        if elementos_logout:
            print("✅ LOGIN EXITOSO - Usuario autenticado")
            return True
        
        # Verificar si hay cookies de sesión
        cookies = driver.get_cookies()
        session_cookie = None
        for cookie in cookies:
            if 'session' in cookie.get('name', '').lower() or 'auth' in cookie.get('name', '').lower():
                session_cookie = cookie
                break
        
        if session_cookie:
            print(f"✅ Cookie de sesión encontrada: {session_cookie.get('name')}")
            return True
        
        print("❌ No se pudo confirmar el login")
        return False
        
    except Exception as e:
        print(f"❌ Error en login: {str(e)}")
        return False

# ============================================
# FUNCIÓN PARA VERIFICAR ACCESO A DATOS
# ============================================
def verificar_acceso_datos(driver):
    """Verifica si se pueden ver los datos de precios"""
    print("\n=== VERIFICANDO ACCESO A DATOS ===")
    
    test_url = "https://www.metal.com/Lithium/201102250059"
    driver.get(test_url)
    time.sleep(8)
    
    page_source = driver.page_source
    
    if "Sign in to view" in page_source:
        print("❌ No se puede acceder a los datos - Pide autenticación")
        return False
    
    # Verificar si hay datos reales (números con formato de precio)
    import re
    price_pattern = r'\d+[,.]?\d*'
    numbers = re.findall(price_pattern, page_source)
    
    # Si hay más de 5 números en la página, probablemente hay datos
    if len(numbers) > 5:
        print(f"✅ Se puede acceder a los datos ({len(numbers)} números encontrados)")
        return True
    else:
        print("⚠️ La página no pide login pero hay pocos números")
        return False

# ============================================
# EJECUTAR LOGIN
# ============================================
login_exitoso = realizar_login_definitivo(driver, user, password)

if not login_exitoso:
    print("\n❌❌❌ LOGIN FALLIDO - DETENIENDO EJECUCIÓN ❌❌❌")
    driver.quit()
    sys.exit(1)

# Verificar acceso a datos
if not verificar_acceso_datos(driver):
    print("\n❌❌❌ NO SE PUEDE ACCEDER A LOS DATOS - DETENIENDO EJECUCIÓN ❌❌❌")
    driver.quit()
    sys.exit(1)

print("\n✅ Login verificado - Continuando con scraping...")

# =========================
# FUNCIONES DE SCRAPING
# =========================

def page_not_found(driver):
    """Verifica si la página existe y tiene datos"""
    try:
        time.sleep(2)
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
        time.sleep(8)
        
        # Verificar si la página pide login
        if "Sign in to view" in driver.page_source:
            print("  ❌ La página pide autenticación - Login no funcionó")
            return None, None
        
        if page_not_found(driver):
            print(f"⚠️ Página no encontrada: {url}")
            return None, None
        
        # Buscar contenedor
        container = None
        try:
            container = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, '//div[contains(@class, "__PriceWrap")]'))
            )
            print("  ✅ Contenedor __PriceWrap encontrado")
        except:
            try:
                container = WebDriverWait(driver, 15).until(
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
    data_carbonate.append(price if price else "")
    data_carbonate.append(range_price if range_price else "")
    time.sleep(3)

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
    data_hydroxide.append(price if price else "")
    data_hydroxide.append(range_price if range_price else "")
    time.sleep(3)

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
    data_metal.append(price if price else "")
    data_metal.append(range_price if range_price else "")
    time.sleep(3)

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
    data_other.append(price if price else "")
    data_other.append(range_price if range_price else "")
    time.sleep(3)

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
# VERIFICAR QUE SE EXTRAJERON DATOS
# ============================================
print("\n=== VERIFICANDO DATOS EXTRAÍDOS ===")

# Verificar si hay datos en los DataFrames
tiene_datos = False

# Función para verificar si un DataFrame tiene datos válidos
def df_tiene_datos(df):
    if df.empty:
        return False
    for col in df.columns:
        if df[col].notna().any() and (df[col] != "").any() and (df[col] != "N/A").any():
            return True
    return False

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
    driver.quit()
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

driver.quit()
print("\n✅ Proceso completado exitosamente - Email enviado con datos")
