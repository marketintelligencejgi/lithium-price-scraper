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

print("=== BUSCANDO PASSWORD EN EL DOM VISIBLE ===", flush=True)

resultado = driver.execute_script("""
    const elementos = document.querySelectorAll('*');
    const encontrados = [];

    for (const elemento of elementos) {

        const texto = (elemento.innerText || '').trim();

        if (
            elemento.offsetParent !== null &&
            texto.includes('Password')
        ) {
            encontrados.push({
                tag: elemento.tagName,
                id: elemento.id,
                class: elemento.className,
                texto: texto.substring(0, 200)
            });
        }
    }

    return encontrados;
""")

print(
    f"Elementos visibles que contienen Password: {len(resultado)}",
    flush=True
)

for elemento in resultado:
    print(elemento, flush=True)

print("=== BUSCANDO WELCOME EN EL DOM VISIBLE ===", flush=True)

resultado = driver.execute_script("""
    const elementos = document.querySelectorAll('*');
    const encontrados = [];

    for (const elemento of elementos) {

        const texto = (elemento.innerText || '').trim();

        if (
            elemento.offsetParent !== null &&
            texto.includes('Welcome to SMM')
        ) {
            encontrados.push({
                tag: elemento.tagName,
                id: elemento.id,
                class: elemento.className,
                texto: texto.substring(0, 200)
            });
        }
    }

    return encontrados;
""")

print(
    f"Elementos visibles que contienen Welcome: {len(resultado)}",
    flush=True
)

for elemento in resultado:
    print(elemento, flush=True)

dialogs = driver.find_elements(
    By.CSS_SELECTOR,
    "[role='dialog']"
)

print(
    f"Elementos role=dialog: {len(dialogs)}",
    flush=True
)

dialogs = driver.find_elements(
    By.CSS_SELECTOR,
    "dialog"
)

print(
    f"Elementos <dialog>: {len(dialogs)}",
    flush=True
)

