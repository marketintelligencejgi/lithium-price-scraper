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
# LOGIN - ESPERANDO CARGA DINÁMICA DEL POPUP
# ============================================

try:
    # Primero, asegurarnos de que la página esté completamente cargada
    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, "//button[contains(text(), 'Sign In')]"))
    )
    print("Página cargada correctamente")
    
    # Hacer clic en Sign In con JavaScript
    print("Abriendo popup de login...")
    driver.execute_script("""
        const buttons = document.querySelectorAll('button');
        for (let btn of buttons) {
            if (btn.textContent.includes('Sign In')) {
                btn.click();
                break;
            }
        }
    """)
    print("Clic en Sign In ejecutado")
    
    # ============================================
    # ESPERAR A QUE EL CONTENIDO DEL POPUP SE CARGUE
    # ============================================
    print("Esperando carga del popup...")
    time.sleep(3)
    
    # Buscar el popup por role='dialog'
    popup = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//div[@role='dialog']"))
    )
    print("Popup encontrado")
    
    # ============================================
    # ESPERAR A QUE APAREZCAN LOS INPUTS DENTRO DEL POPUP
    # ============================================
    print("Esperando inputs dentro del popup...")
    
    # Método 1: Esperar específicamente a que haya inputs
    input_user = None
    intentos = 0
    max_intentos = 10
    
    while intentos < max_intentos and not input_user:
        try:
            # Buscar inputs dentro del popup
            inputs = popup.find_elements(By.XPATH, ".//input")
            print(f"Intento {intentos + 1}: {len(inputs)} inputs encontrados en popup")
            
            for inp in inputs:
                try:
                    input_type = inp.get_attribute('type')
                    placeholder = inp.get_attribute('placeholder')
                    print(f"  Input: type={input_type}, placeholder={placeholder}")
                    
                    # Buscar campo de email/usuario
                    if input_type == 'email' or (placeholder and ('email' in placeholder.lower() or 'usuario' in placeholder.lower())):
                        input_user = inp
                        print("¡Campo de usuario encontrado!")
                        break
                except:
                    pass
            
            if not input_user:
                # También buscar inputs de tipo text
                for inp in inputs:
                    try:
                        if inp.get_attribute('type') == 'text' and inp.is_displayed():
                            input_user = inp
                            print("¡Campo de usuario (text) encontrado!")
                            break
                    except:
                        pass
            
            if not input_user:
                # Esperar un poco más antes del siguiente intento
                time.sleep(1)
                intentos += 1
                # Actualizar referencia al popup (por si cambia)
                popup = driver.find_element(By.XPATH, "//div[@role='dialog']")
                
        except Exception as e:
            print(f"Error en intento {intentos + 1}: {e}")
            time.sleep(1)
            intentos += 1
    
    # ============================================
    # SI NO SE ENCUENTRA EL INPUT, FORZAR CON JAVASCRIPT
    # ============================================
    if not input_user:
        print("No se encontraron inputs en el popup, forzando con JavaScript...")
        
        # Usar JavaScript para encontrar y mostrar cualquier input oculto
        driver.execute_script("""
            // Buscar todos los inputs en la página
            const inputs = document.querySelectorAll('input');
            console.log('Total inputs:', inputs.length);
            
            // Forzar visibilidad de todos los inputs
            for (let inp of inputs) {
                // Hacer visible el input
                inp.style.display = 'block';
                inp.style.visibility = 'visible';
                inp.style.opacity = '1';
                
                // Si el input está dentro de un contenedor oculto, mostrar el contenedor
                let parent = inp.parentElement;
                while (parent) {
                    parent.style.display = 'block';
                    parent.style.visibility = 'visible';
                    parent.style.opacity = '1';
                    parent = parent.parentElement;
                }
            }
            
            // También mostrar todos los contenedores de formulario
            const forms = document.querySelectorAll('form');
            for (let form of forms) {
                form.style.display = 'block';
                form.style.visibility = 'visible';
                form.style.opacity = '1';
            }
        """)
        print("Forzada visibilidad de inputs")
        time.sleep(2)
        
        # Intentar encontrar inputs nuevamente
        try:
            # Buscar en toda la página
            all_inputs = driver.find_elements(By.XPATH, "//input")
            print(f"Total inputs en página: {len(all_inputs)}")
            
            for inp in all_inputs:
                try:
                    if inp.is_displayed():
                        input_type = inp.get_attribute('type')
                        placeholder = inp.get_attribute('placeholder')
                        print(f"Input visible: type={input_type}, placeholder={placeholder}")
                        
                        if input_type == 'email' or (placeholder and 'email' in placeholder.lower()):
                            input_user = inp
                            print("¡Campo de usuario encontrado después de forzar!")
                            break
                except:
                    pass
        except:
            pass
    
    # ============================================
    # SI AÚN NO SE ENCUENTRA, USAR SELENIUM PARA ESPERAR
    # ============================================
    if not input_user:
        print("Usando WebDriverWait para esperar inputs...")
        try:
            # Esperar a que aparezca cualquier input dentro del popup
            wait = WebDriverWait(driver, 20)
            input_user = wait.until(
                EC.presence_of_element_located((By.XPATH, "//div[@role='dialog']//input"))
            )
            print("Input encontrado con WebDriverWait!")
        except Exception as e:
            print(f"WebDriverWait falló: {e}")
    
    if not input_user:
        # Guardar debug final
        print("\n=== DEBUG FINAL ===")
        with open('debug_final.html', 'w', encoding='utf-8') as f:
            f.write(driver.page_source)
        driver.save_screenshot("debug_final.png")
        
        # Listar todos los elementos dentro del popup
        try:
            popup = driver.find_element(By.XPATH, "//div[@role='dialog']")
            html_inner = popup.get_attribute('innerHTML')
            print("Contenido del popup:")
            print(html_inner[:500])
        except:
            pass
        
        raise Exception("No se pudo encontrar el campo de usuario después de todos los intentos")
    
    # ============================================
    # ENCONTRAR CONTRASEÑA Y BOTÓN
    # ============================================
    print("\n--- Buscando contraseña y botón ---")
    
    # Buscar contraseña
    input_pass = None
    try:
        # Buscar en el popup
        input_pass = driver.find_element(By.XPATH, "//div[@role='dialog']//input[@type='password']")
        print("Contraseña encontrada en popup")
    except:
        try:
            # Buscar en toda la página
            input_pass = driver.find_element(By.XPATH, "//input[@type='password']")
            print("Contraseña encontrada en página")
        except:
            pass
    
    if not input_pass:
        raise Exception("No se pudo encontrar el campo de contraseña")
    
    # Buscar botón login
    boton_login = None
    try:
        # Buscar botón dentro del popup
        boton_login = driver.find_element(By.XPATH, "//div[@role='dialog']//button[@type='submit']")
        print("Botón encontrado en popup")
    except:
        try:
            # Buscar botón por texto
            boton_login = driver.find_element(By.XPATH, "//div[@role='dialog']//button[contains(text(), 'Sign In')]")
            print("Botón encontrado por texto")
        except:
            try:
                # Buscar en toda la página
                boton_login = driver.find_element(By.XPATH, "//button[@type='submit']")
                print("Botón encontrado en página")
            except:
                pass
    
    if not boton_login:
        raise Exception("No se pudo encontrar el botón de login")
    
    # ============================================
    # INGRESAR CREDENCIALES
    # ============================================
    print("\n--- Ingresando credenciales ---")
    
    # Limpiar y escribir
    input_user.clear()
    input_user.send_keys(user)
    print("Usuario ingresado")
    
    input_pass.clear()
    input_pass.send_keys(password)
    print("Contraseña ingresada")
    
    # Hacer clic en login
    boton_login.click()
    print("Login enviado")
    
    # Esperar a que cargue
    time.sleep(random.uniform(5, 8))
    
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
