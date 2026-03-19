from selenium.webdriver.chrome.service import Service
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from datetime import datetime
import pandas as pd
import time
import smtplib
from email.message import EmailMessage
import os

###----------------------------------------------------------------------> INICIO <----------------------------------------------------------------------###

user = os.environ["METAL_USER"]
password = os.environ["METAL_PASS"]

# =========================
# CONFIGURACIÓN CHROME (GITHUB)
# =========================

options = Options()
options.add_argument("--headless=new")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--window-size=1920,1080")

# evitar detección selenium
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option("useAutomationExtension", False)

options.add_argument(
"user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36"
)

service = Service(ChromeDriverManager().install())

driver = webdriver.Chrome(service=service, options=options)

driver.execute_script(
"Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
)

wait = WebDriverWait(driver,10)

driver.get('https://www.metal.com/')

wait.until(EC.presence_of_element_located((By.XPATH,'/html/body/div[2]/main/header')))

# Sign in
boton = driver.find_element(By.XPATH, '/html/body/div[2]/main/header/div[2]/div/div/div[2]/div/div[1]')
boton.click()

input_user = wait.until(EC.presence_of_element_located((By.XPATH,'//*[@id="account"]')))
input_pass = driver.find_element(By.XPATH, '//*[@id="password"]')
boton = driver.find_element(By.XPATH, '//*[@id="action"]/div[1]/div[1]/div/div[1]/form/div[4]/div/div/div/div/button')

input_user.send_keys(user)
input_pass.send_keys(password)
boton.click()

del(user, password, input_user, input_pass, boton)

wait.until(EC.presence_of_element_located((By.XPATH,'//div[contains(@class,"PriceWrap")]')))

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

    try:
        container = WebDriverWait(driver,10).until(
            EC.presence_of_element_located((By.XPATH,'//div[contains(@class,"__PriceWrap")]'))
        )
    except:
        return None, None

    if page_not_found(driver):
        return None, None
    
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

table = WebDriverWait(driver,10).until(
    EC.presence_of_element_located((By.CSS_SELECTOR,".ant-table-content table"))
)

df_rare_earth = pd.read_html(table.get_attribute("outerHTML"))[0]

df_rare_earth['Name'] = df_rare_earth['Name'].str.replace(r'SMM.*$', '', regex=True).str.strip()

driver.quit()

file_name = "Reporte_Diario.xlsx"

with pd.ExcelWriter(file_name, engine="xlsxwriter") as writer:

    df_lithium_carbonate.to_excel(writer, sheet_name="Lithium carbonate", index=False)

    df_lithium_hydroxide.to_excel(writer, sheet_name="Lithium hydroxide", index=False)

    df_lithium_metal.to_excel(writer, sheet_name="Lithium metal", index=False)

    df_other.to_excel(writer, sheet_name="Other", index=False)

    df_rare_earth.to_excel(writer, sheet_name="REO", index=False)

sender = os.environ["EMAIL_USER"]
password = os.environ["EMAIL_PASS"]

receiver = "market.intelligence@JGI.be"

msg = EmailMessage()

msg["Subject"] = "Price Tracking Data"
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