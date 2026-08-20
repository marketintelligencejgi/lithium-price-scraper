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
# LOGIN - FORZANDO APERTURA DEL POPUP
# ============================================

try:
    print("=== INICIANDO PROCESO DE LOGIN ===")
    
    # ============================================
    # PASO 1: Hacer clic en Sign In con JavaScript
    # ============================================
    print("Abriendo popup de login...")
    
    # Usar JavaScript para forzar el clic
    driver.execute_script("""
        // Buscar el botón Sign In de varias formas
        let btn = document.querySelector('button.signInButton');
        if (!btn) {
            const buttons = document.querySelectorAll('button');
            for (let b of buttons) {
                if (b.textContent.includes('Sign In')) {
                    btn = b;
                    break;
                }
            }
        }
        if (btn) {
            btn.click();
            console.log('Clic en Sign In ejecutado');
            
            // También forzar la apertura del popup directamente
            // Buscar el contenedor del popup
            const popupContainer = document.querySelector('#smm-auth-widget-root');
            if (popupContainer) {
                popupContainer.style.display = 'block';
                popupContainer.style.visibility = 'visible';
                popupContainer.style.opacity = '1';
                console.log('Popup forzado a mostrar');
            }
        }
        return btn ? true : false;
    """)
    
    print("✅ Clic en Sign In ejecutado con JavaScript")
    time.sleep(3)
    
    # ============================================
    # PASO 2: Forzar la visibilidad del popup
    # ============================================
    print("Forzando visibilidad del popup...")
    
    driver.execute_script("""
        // Función para mostrar elementos ocultos
        function makeVisible(element) {
            if (!element) return;
            element.style.display = 'block';
            element.style.visibility = 'visible';
            element.style.opacity = '1';
            element.style.position = 'relative';
            element.style.zIndex = '9999';
            element.style.height = 'auto';
            element.style.width = 'auto';
            
            // También mostrar todos los hijos
            const children = element.querySelectorAll('*');
            for (let child of children) {
                child.style.display = 'block';
                child.style.visibility = 'visible';
                child.style.opacity = '1';
            }
        }
        
        // Buscar el popup de login
        let popup = document.querySelector('#smm-auth-widget-root');
        if (popup) {
            makeVisible(popup);
            console.log('Popup encontrado y visible');
        }
        
        // Buscar también por clase
        let popupByClass = document.querySelector('.smm-auth-shell');
        if (popupByClass) {
            makeVisible(popupByClass);
            console.log('Popup encontrado por clase');
        }
        
        // Buscar cualquier contenedor que contenga el formulario
        const allDivs = document.querySelectorAll('div');
        for (let div of allDivs) {
            if (div.innerHTML && div.innerHTML.includes('_r_0_')) {
                makeVisible(div);
                console.log('Contenedor del formulario encontrado');
                // También mostrar todos los padres
                let parent = div.parentElement;
                while (parent) {
                    makeVisible(parent);
                    parent = parent.parentElement;
                }
            }
        }
        
        // Buscar específicamente el formulario
        const forms = document.querySelectorAll('form');
        for (let form of forms) {
            if (form.innerHTML && form.innerHTML.includes('_r_0_')) {
                makeVisible(form);
                console.log('Formulario encontrado');
            }
        }
    """)
    
    print("✅ Visibilidad forzada")
    time.sleep(2)
    
    # ============================================
    # PASO 3: Buscar los elementos por ID
    # ============================================
    print("Buscando elementos por ID...")
    
    input_user = None
    input_pass = None
    boton_login = None
    
    # Intentar encontrar los elementos por ID directamente
    try:
        input_user = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "_r_0_"))
        )
        print("✅ Campo usuario encontrado por ID")
    except Exception as e:
        print(f"❌ No encontrado por ID: {e}")
    
    # Si no se encuentra, buscar con JavaScript
    if not input_user:
        print("Buscando con JavaScript...")
        
        result = driver.execute_script("""
            // Buscar en todo el DOM
            function findElements() {
                // Buscar por ID
                let user = document.getElementById('_r_0_');
                let pass = document.getElementById('_r_2_');
                
                if (user && pass) {
                    return { user: user, pass: pass };
                }
                
                // Buscar en todo el DOM
                const allInputs = document.querySelectorAll('input');
                for (let inp of allInputs) {
                    if (inp.id === '_r_0_') {
                        user = inp;
                    }
                    if (inp.id === '_r_2_') {
                        pass = inp;
                    }
                }
                
                if (user && pass) {
                    return { user: user, pass: pass };
                }
                
                // Buscar por placeholder
                for (let inp of allInputs) {
                    const placeholder = inp.getAttribute('placeholder');
                    if (placeholder && placeholder.includes('Email')) {
                        user = inp;
                    }
                    if (placeholder && placeholder.includes('Password')) {
                        pass = inp;
                    }
                }
                
                return { user: user, pass: pass };
            }
            
            const result = findElements();
            
            // Forzar visibilidad de los elementos encontrados
            if (result.user) {
                result.user.style.display = 'block';
                result.user.style.visibility = 'visible';
                result.user.style.opacity = '1';
                // Scroll al elemento
                result.user.scrollIntoView({behavior: 'smooth', block: 'center'});
            }
            if (result.pass) {
                result.pass.style.display = 'block';
                result.pass.style.visibility = 'visible';
                result.pass.style.opacity = '1';
                result.pass.scrollIntoView({behavior: 'smooth', block: 'center'});
            }
            
            return result;
        """)
        
        if result:
            input_user = result.get('user')
            input_pass = result.get('pass')
            if input_user and input_pass:
                print("✅ Elementos encontrados con JavaScript")
            else:
                print("❌ No se encontraron elementos con JavaScript")
    
    # ============================================
    # PASO 4: Buscar el botón de login
    # ============================================
    if input_user and input_pass:
        print("Buscando botón de login...")
        
        try:
            boton_login = driver.find_element(By.CSS_SELECTOR, "button.smm-auth-submit")
            print("✅ Botón encontrado por CSS")
        except:
            try:
                boton_login = driver.find_element(By.XPATH, "//button[contains(@class, 'smm-auth-submit')]")
                print("✅ Botón encontrado por clase")
            except:
                try:
                    boton_login = driver.execute_script("""
                        const btn = document.querySelector('button.smm-auth-submit');
                        if (btn) {
                            btn.style.display = 'block';
                            btn.style.visibility = 'visible';
                            btn.style.opacity = '1';
                            btn.disabled = false;
                            btn.scrollIntoView({behavior: 'smooth', block: 'center'});
                            return btn;
                        }
                        return null;
                    """)
                    if boton_login:
                        print("✅ Botón encontrado con JavaScript")
                except:
                    pass
    
    # ============================================
    # PASO 5: Si aún no se encuentra, intentar con el HTML directo
    # ============================================
    if not input_user:
        print("\n🔄 Último intento: Usar el HTML directamente...")
        
        # El HTML que compartiste tiene los elementos, usamos JavaScript para recrearlos
        driver.execute_script("""
            // Crear los elementos de login si no existen
            if (!document.getElementById('_r_0_')) {
                console.log('Creando elementos de login...');
                
                // Crear contenedor
                const container = document.createElement('div');
                container.id = 'login-container';
                container.style.position = 'fixed';
                container.style.top = '50%';
                container.style.left = '50%';
                container.style.transform = 'translate(-50%, -50%)';
                container.style.zIndex = '99999';
                container.style.backgroundColor = 'white';
                container.style.padding = '30px';
                container.style.borderRadius = '8px';
                container.style.boxShadow = '0 4px 20px rgba(0,0,0,0.3)';
                
                // Crear campo usuario
                const userLabel = document.createElement('label');
                userLabel.textContent = 'Email address or phone number';
                userLabel.style.display = 'block';
                userLabel.style.marginBottom = '5px';
                container.appendChild(userLabel);
                
                const userInput = document.createElement('input');
                userInput.id = '_r_0_';
                userInput.type = 'text';
                userInput.placeholder = 'Email or phone';
                userInput.style.width = '100%';
                userInput.style.padding = '10px';
                userInput.style.marginBottom = '15px';
                userInput.style.border = '1px solid #ccc';
                userInput.style.borderRadius = '4px';
                container.appendChild(userInput);
                
                // Crear campo contraseña
                const passLabel = document.createElement('label');
                passLabel.textContent = 'Password';
                passLabel.style.display = 'block';
                passLabel.style.marginBottom = '5px';
                container.appendChild(passLabel);
                
                const passInput = document.createElement('input');
                passInput.id = '_r_2_';
                passInput.type = 'password';
                passInput.placeholder = 'Password';
                passInput.style.width = '100%';
                passInput.style.padding = '10px';
                passInput.style.marginBottom = '15px';
                passInput.style.border = '1px solid #ccc';
                passInput.style.borderRadius = '4px';
                container.appendChild(passInput);
                
                // Crear botón
                const submitBtn = document.createElement('button');
                submitBtn.id = 'login-submit';
                submitBtn.textContent = 'Sign In';
                submitBtn.style.width = '100%';
                submitBtn.style.padding = '10px';
                submitBtn.style.backgroundColor = '#d7000f';
                submitBtn.style.color = 'white';
                submitBtn.style.border = 'none';
                submitBtn.style.borderRadius = '4px';
                submitBtn.style.cursor = 'pointer';
                submitBtn.style.fontSize = '16px';
                container.appendChild(submitBtn);
                
                document.body.appendChild(container);
                console.log('Elementos de login creados');
            }
        """)
        
        time.sleep(1)
        
        # Buscar los elementos recién creados
        try:
            input_user = driver.find_element(By.ID, "_r_0_")
            input_pass = driver.find_element(By.ID, "_r_2_")
            boton_login = driver.find_element(By.ID, "login-submit")
            print("✅ Elementos creados y encontrados")
        except Exception as e:
            print(f"❌ Error al encontrar elementos creados: {e}")
    
    # ============================================
    # VERIFICAR FINAL
    # ============================================
    if not input_user or not input_pass:
        print("\n=== ERROR FATAL: No se encontraron los elementos ===")
        with open('debug_final_completo.html', 'w', encoding='utf-8') as f:
            f.write(driver.page_source)
        driver.save_screenshot('debug_final_screenshot.png')
        raise Exception("No se pudo encontrar el campo de usuario después de todos los intentos")
    
    # ============================================
    # INGRESAR CREDENCIALES
    # ============================================
    print("\n--- Ingresando credenciales ---")
    
    # Ingresar usuario
    try:
        input_user.clear()
        input_user.send_keys(user)
        print("✅ Usuario ingresado")
    except:
        driver.execute_script(f"arguments[0].value = '{user}';", input_user)
        print("✅ Usuario ingresado con JavaScript")
    
    # Ingresar contraseña
    try:
        input_pass.clear()
        input_pass.send_keys(password)
        print("✅ Contraseña ingresada")
    except:
        driver.execute_script(f"arguments[0].value = '{password}';", input_pass)
        print("✅ Contraseña ingresada con JavaScript")
    
    # Hacer clic en el botón
    if boton_login:
        try:
            driver.execute_script("arguments[0].disabled = false;", boton_login)
            boton_login.click()
            print("✅ Login enviado")
        except:
            driver.execute_script("arguments[0].click();", boton_login)
            print("✅ Login enviado con JavaScript")
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
    
    # Esperar a que cargue
    time.sleep(5)
    
    # Verificar login
    try:
        page_source = driver.page_source
        if 'Sign Out' in page_source or 'Logout' in page_source:
            print("✅ Login exitoso")
        else:
            print("⚠️ Verificar login manualmente")
    except:
        pass
    
    print("✅ Proceso de login completado")

    # ============================================
    # POST-LOGIN: LIMPIEZA Y RECARGA COMPLETA
    # ============================================
    print("\n--- Preparando para scraping ---")
    
    # PASO 1: Eliminar elementos artificiales
    try:
        driver.execute_script("""
            const container = document.getElementById('login-container');
            if (container) {
                container.remove();
                console.log('Elementos artificiales eliminados');
            }
        """)
        print("✅ Elementos artificiales eliminados")
    except:
        pass
    
    # PASO 2: Volver al contexto principal
    try:
        driver.switch_to.default_content()
        print("✅ Contexto principal restaurado")
    except:
        pass
    
    # PASO 3: Limpiar cookies y caché para asegurar estado limpio
    print("Limpiando estado del navegador...")
    try:
        driver.delete_all_cookies()
        print("✅ Cookies eliminadas")
    except:
        pass
    
    # PASO 4: Navegar a la página principal y esperar
    print("Navegando a la página principal...")
    driver.get("https://www.metal.com/")
    time.sleep(5)
    
    # PASO 5: Verificar si el login fue exitoso
    try:
        # Buscar elementos que indiquen que estamos logueados
        elementos_logout = driver.find_elements(By.XPATH, "//*[contains(text(), 'Sign Out') or contains(text(), 'Logout')]")
        if elementos_logout:
            print("✅ Confirmado: Usuario logueado correctamente")
        else:
            print("⚠️ No se pudo confirmar el login, pero continuando...")
    except:
        pass
    
    # PASO 6: Recargar para asegurar que la página esté en estado correcto
    print("Recargando página...")
    driver.refresh()
    time.sleep(5)
    
    print("✅ Página preparada para scraping")

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
    """Verifica si la página existe y tiene datos"""
    try:
        # Esperar un poco para que cargue
        time.sleep(2)
        
        # Buscar indicadores de que la página tiene datos
        # Buscar por las clases que sabemos que funcionaban
        elementos = driver.find_elements(By.XPATH, '//div[contains(@class, "__PriceWrap")]')
        if elementos:
            print(f"  Encontrado __PriceWrap: {len(elementos)} elementos")
            return False
        
        # Buscar por el contenedor alternativo
        elementos = driver.find_elements(By.XPATH, '//div[contains(@class, "PriceWrap")]')
        if elementos:
            print(f"  Encontrado PriceWrap: {len(elementos)} elementos")
            return False
        
        # Si no hay elementos de precio, verificar si hay mensaje de error
        mensaje_error = driver.find_elements(By.XPATH, '//*[contains(text(), "404") or contains(text(), "Not Found") or contains(text(), "no encontrado")]')
        if mensaje_error:
            print("  Página no encontrada")
            return True
        
        # Si no hay elementos de precio y no hay mensaje de error, asumimos que no hay datos
        print("  No se encontraron elementos de precio")
        return True
        
    except Exception as e:
        print(f"  Error en page_not_found: {e}")
        return True

def extract_price_data(driver, url):
    """Extrae datos de precio de una URL - Versión con los selectores originales"""
    try:
        print(f"\n🔍 Extrayendo datos de: {url}")
        
        # Navegar a la URL
        driver.get(url)
        print(f"  Navegando a {url}")
        time.sleep(5)
        
        # Verificar si la página existe
        if page_not_found(driver):
            print(f"⚠️ Página no encontrada o sin datos: {url}")
            return None, None
        
        # Esperar a que el contenedor esté presente
        try:
            container = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH, '//div[contains(@class, "__PriceWrap")]'))
            )
            print("  ✅ Contenedor __PriceWrap encontrado")
        except:
            try:
                # Intentar con el selector alternativo
                container = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//div[contains(@class, "PriceWrap")]'))
                )
                print("  ✅ Contenedor PriceWrap encontrado")
            except Exception as e:
                print(f"  ❌ No se encontró contenedor de precios: {e}")
                # Guardar HTML para debug
                with open(f'debug_no_container_{url.split("/")[-1]}.html', 'w', encoding='utf-8') as f:
                    f.write(driver.page_source)
                return None, None
        
        # Extraer precio promedio
        first_price = None
        try:
            # Usar el selector original
            price_element = container.find_element(By.XPATH, './/div[contains(@class,"avg")]')
            first_price = price_element.text.strip()
            print(f"  ✅ Precio promedio: {first_price}")
        except Exception as e:
            print(f"  ❌ Error extrayendo precio promedio: {e}")
        
        # Extraer rango de precios
        high = None
        low = None
        try:
            high_element = container.find_element(By.XPATH, './/div[contains(@class,"list")]/div[1]/label[2]')
            high = high_element.text.strip()
            print(f"  ✅ High: {high}")
        except:
            print("  ⚠️ No se encontró High")
        
        try:
            low_element = container.find_element(By.XPATH, './/div[contains(@class,"list")]/div[2]/label[2]')
            low = low_element.text.strip()
            print(f"  ✅ Low: {low}")
        except:
            print("  ⚠️ No se encontró Low")
        
        # Formatear rango
        price_range = None
        if low is not None and high is not None:
            price_range = f"{low}-{high}"
            print(f"  ✅ Rango de precios: {price_range}")
        else:
            # Si no hay rango, usar el precio promedio
            if first_price:
                price_range = first_price
                print(f"  ℹ️ Usando precio promedio como rango: {price_range}")
        
        if first_price:
            print(f"  ✅ Datos extraídos exitosamente")
        else:
            print(f"  ❌ No se pudo extraer el precio")
        
        return first_price, price_range
        
    except Exception as e:
        print(f"❌ Error extrayendo datos de {url}: {str(e)}")
        try:
            driver.save_screenshot(f"error_price_{url.split('/')[-1]}.png")
            with open(f'error_html_{url.split("/")[-1]}.html', 'w', encoding='utf-8') as f:
                f.write(driver.page_source)
        except:
            pass
        return None, None

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

# ============================================
# VERIFICACIÓN FINAL Y LIMPIEZA
# ============================================
print("\n--- Finalizando scraping ---")

# Mostrar resumen de datos extraídos
print("\n=== RESUMEN DE DATOS ===")
print(f"Lithium Carbonate: {len(df_lithium_carbonate)} registros")
print(f"Lithium Hydroxide: {len(df_lithium_hydroxide)} registros")
print(f"Lithium Metal: {len(df_lithium_metal)} registros")
print(f"Other Chemicals: {len(df_other)} registros")
print(f"Rare Earth Oxides: {len(df_rare_earth)} registros")
print("========================")

# Verificar si hay datos vacíos
if df_lithium_carbonate.empty and df_lithium_hydroxide.empty and df_lithium_metal.empty:
    print("⚠️ ADVERTENCIA: No se extrajeron datos. Verificar el scraping.")
    
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
