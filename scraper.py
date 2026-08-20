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
# FUNCIÓN PARA LOGIN USANDO COOKIES
# ============================================
def login_con_cookies(driver, user, password):
    """
    Intenta hacer login usando diferentes estrategias
    """
    print("\n=== INTENTANDO LOGIN ===")
    
    # Estrategia 1: Intentar hacer clic en el botón Sign In y llenar el formulario
    try:
        print("Estrategia 1: Clic en Sign In...")
        boton = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Sign In')]"))
        )
        driver.execute_script("arguments[0].click();", boton)
        time.sleep(3)
        print("✅ Clic en Sign In")
        
        # Buscar los campos en el DOM (sin Shadow DOM)
        try:
            # Buscar en el DOM principal
            input_user = driver.find_element(By.ID, "_r_0_")
            input_pass = driver.find_element(By.ID, "_r_2_")
            print("✅ Campos encontrados en DOM principal")
            
            input_user.send_keys(user)
            input_pass.send_keys(password)
            
            # Buscar botón de login
            try:
                btn_login = driver.find_element(By.XPATH, "//button[contains(@class, 'smm-auth-submit')]")
                driver.execute_script("arguments[0].click();", btn_login)
            except:
                # Intentar con cualquier botón tipo submit
                btn_login = driver.find_element(By.XPATH, "//button[@type='submit']")
                driver.execute_script("arguments[0].click();", btn_login)
            
            print("✅ Login enviado")
            time.sleep(5)
            
            # Verificar login
            driver.get("https://www.metal.com/")
            time.sleep(3)
            
            # Buscar elementos de usuario logueado
            if driver.find_elements(By.XPATH, "//*[contains(text(), 'Sign Out')]"):
                print("✅ LOGIN EXITOSO (Estrategia 1)")
                return True
                
        except Exception as e:
            print(f"❌ Error en estrategia 1: {e}")
            
    except Exception as e:
        print(f"❌ Error en estrategia 1: {e}")
    
    # Estrategia 2: Intentar con JavaScript directamente
    try:
        print("\nEstrategia 2: JavaScript directo...")
        
        result = driver.execute_script(f"""
            // Buscar el botón Sign In
            const btn = document.querySelector('button.signInButton');
            if (btn) btn.click();
            
            // Esperar y buscar los campos
            setTimeout(() => {{
                // Buscar en toda la página
                const inputs = document.querySelectorAll('input');
                let userInput = null;
                let passInput = null;
                
                for (let inp of inputs) {{
                    if (inp.id === '_r_0_') userInput = inp;
                    if (inp.id === '_r_2_') passInput = inp;
                }}
                
                if (userInput && passInput) {{
                    userInput.value = '{user}';
                    passInput.value = '{password}';
                    
                    // Disparar eventos
                    userInput.dispatchEvent(new Event('input', {{ bubbles: true }}));
                    passInput.dispatchEvent(new Event('input', {{ bubbles: true }}));
                    
                    // Buscar y hacer clic en el botón
                    const submitBtn = document.querySelector('button.smm-auth-submit');
                    if (submitBtn) {{
                        submitBtn.disabled = false;
                        submitBtn.click();
                        return 'login_enviado';
                    }}
                }}
                return 'no_encontrado';
            }}, 2000);
            
            return 'ejecutando';
        """)
        
        print(f"Resultado JavaScript: {result}")
        time.sleep(8)
        
        # Verificar login
        driver.get("https://www.metal.com/")
        time.sleep(3)
        
        if driver.find_elements(By.XPATH, "//*[contains(text(), 'Sign Out')]"):
            print("✅ LOGIN EXITOSO (Estrategia 2)")
            return True
            
    except Exception as e:
        print(f"❌ Error en estrategia 2: {e}")
    
    # Estrategia 3: Usar requests para obtener cookies y luego inyectarlas
    try:
        print("\nEstrategia 3: Login con requests...")
        import requests
        from bs4 import BeautifulSoup
        
        session = requests.Session()
        
        # Obtener la página principal
        response = session.get("https://www.metal.com/")
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # Buscar token CSRF si existe
        csrf_token = None
        token_input = soup.find('input', {'name': 'csrf_token'})
        if token_input:
            csrf_token = token_input.get('value')
        
        # Preparar datos de login
        login_data = {
            'username': user,
            'password': password,
        }
        if csrf_token:
            login_data['csrf_token'] = csrf_token
        
        # Intentar login
        login_response = session.post('https://www.metal.com/api/login', data=login_data)
        
        if login_response.status_code == 200:
            print("✅ Login con requests exitoso")
            
            # Obtener cookies y agregarlas al driver
            cookies = session.cookies.get_dict()
            for name, value in cookies.items():
                driver.add_cookie({'name': name, 'value': value})
            
            # Recargar la página
            driver.get("https://www.metal.com/")
            time.sleep(3)
            
            if driver.find_elements(By.XPATH, "//*[contains(text(), 'Sign Out')]"):
                print("✅ LOGIN EXITOSO (Estrategia 3)")
                return True
        else:
            print(f"❌ Login con requests falló: {login_response.status_code}")
            
    except Exception as e:
        print(f"❌ Error en estrategia 3: {e}")
    
    print("❌ TODAS LAS ESTRATEGIAS DE LOGIN FALLARON")
    return False

# ============================================
# FUNCIÓN PARA VERIFICAR ACCESO A DATOS
# ============================================
def verificar_acceso_datos(driver):
    """Verifica si se pueden ver los datos de precios"""
    print("\n=== VERIFICANDO ACCESO A DATOS ===")
    
    test_url = "https://www.metal.com/Lithium/201102250059"
    driver.get(test_url)
    time.sleep(5)
    
    page_source = driver.page_source
    
    if "Sign in to view" in page_source:
        print("❌ No se puede acceder a los datos - Pide autenticación")
        return False
    else:
        # Verificar si hay datos reales
        if "Price" in page_source and any(c.isdigit() for c in page_source):
            print("✅ Se puede acceder a los datos")
            return True
        else:
            print("⚠️ La página no pide login pero no hay datos visibles")
            return False

# ============================================
# EJECUTAR LOGIN
# ============================================
login_exitoso = login_con_cookies(driver, user, password)

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

# Verificar Lithium Carbonate
if not df_lithium_carbonate.empty:
    # Verificar si hay valores no vacíos (no N/A o vacíos)
    for col in df_lithium_carbonate.columns:
        if df_lithium_carbonate[col].notna().any() and (df_lithium_carbonate[col] != "").any():
            tiene_datos = True
            break

# Verificar Lithium Hydroxide
if not tiene_datos and not df_lithium_hydroxide.empty:
    for col in df_lithium_hydroxide.columns:
        if df_lithium_hydroxide[col].notna().any() and (df_lithium_hydroxide[col] != "").any():
            tiene_datos = True
            break

# Verificar Lithium Metal
if not tiene_datos and not df_lithium_metal.empty:
    for col in df_lithium_metal.columns:
        if df_lithium_metal[col].notna().any() and (df_lithium_metal[col] != "").any():
            tiene_datos = True
            break

# Verificar Other
if not tiene_datos and not df_other.empty:
    for col in df_other.columns:
        if df_other[col].notna().any() and (df_other[col] != "").any():
            tiene_datos = True
            break

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
