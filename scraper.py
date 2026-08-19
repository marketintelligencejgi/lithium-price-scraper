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

print("=== ANALISIS DEL DOM ===", flush=True)

resultado = driver.execute_script("""
    function analizar(root, ruta) {

        let resultado = {
            encontrado: false,
            ruta: ruta,
            shadowRoots: []
        };

        // Buscar el modal
        if (root.querySelector) {

            const modal = root.querySelector('.modalWrapper');

            if (modal) {
                resultado.encontrado = true;
                resultado.ruta = ruta + " -> .modalWrapper";
                return resultado;
            }

            // Buscar elementos con Shadow Root
            const elementos = root.querySelectorAll('*');

            for (let i = 0; i < elementos.length; i++) {

                const elemento = elementos[i];

                if (elemento.shadowRoot) {

                    resultado.shadowRoots.push(
                        elemento.tagName +
                        " class=" +
                        elemento.className +
                        " id=" +
                        elemento.id
                    );

                    const subresultado = analizar(
                        elemento.shadowRoot,
                        ruta + " -> SHADOW(" + elemento.tagName + ")"
                    );

                    if (subresultado.encontrado) {
                        return subresultado;
                    }

                    resultado.shadowRoots =
                        resultado.shadowRoots.concat(
                            subresultado.shadowRoots
                        );
                }
            }
        }

        return resultado;
    }

    return analizar(document, "DOCUMENT");
""")

print(
    f"Modal encontrado: {resultado.encontrado}",
    flush=True
)

print(
    f"Ruta: {resultado.ruta}",
    flush=True
)

print(
    f"Shadow Roots encontrados: {resultado.shadowRoots.length}",
    flush=True
)

for shadow in resultado.shadowRoots:
    print(
        "Shadow Root: " + shadow,
        flush=True
    )
