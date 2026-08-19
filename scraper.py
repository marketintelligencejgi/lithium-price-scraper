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

print("=== INSPECCION CDP ===", flush=True)

dom = driver.execute_cdp_cmd(
    "DOM.getDocument",
    {
        "depth": -1,
        "pierce": True
    }
)

def buscar_nodos(node):
    encontrados = []

    if node.get("nodeName") == "DIV":
        attrs = node.get("attributes", [])

        # attributes viene como:
        # ["class", "modalWrapper", "id", "..."]
        attrs_dict = dict(
            zip(attrs[::2], attrs[1::2])
        )

        if (
            attrs_dict.get("class") == "modalWrapper"
            or "modalWrapper" in attrs_dict.get("class", "")
        ):
            encontrados.append(node)

    for child in node.get("children", []):
        encontrados.extend(buscar_nodos(child))

    for shadow in node.get("shadowRoots", []):
        encontrados.extend(buscar_nodos(shadow))

    return encontrados


encontrados = buscar_nodos(dom["root"])

print(
    f"modalWrapper encontrados por CDP: {len(encontrados)}",
    flush=True
)

for nodo in encontrados:
    print(
        f"NODE encontrado: {nodo.get('nodeName')}",
        flush=True
    )

    print(
        f"BackendNodeId: {nodo.get('backendNodeId')}",
        flush=True
    )

print("=== BUSCANDO ELEMENTOS DEL LOGIN ===", flush=True)

def buscar_login(node):
    encontrados = []

    node_name = node.get("nodeName", "")
    attrs = node.get("attributes", [])

    attrs_dict = dict(
        zip(attrs[::2], attrs[1::2])
    )

    clase = attrs_dict.get("class", "")
    autocomplete = attrs_dict.get("autocomplete", "")
    tipo = attrs_dict.get("type", "")
    nombre = attrs_dict.get("name", "")

    if (
        autocomplete == "username"
        or nombre == "password"
        or "smm-auth-submit" in clase
    ):
        encontrados.append({
            "nodeName": node_name,
            "class": clase,
            "autocomplete": autocomplete,
            "type": tipo,
            "name": nombre,
            "backendNodeId": node.get("backendNodeId")
        })

    for child in node.get("children", []):
        encontrados.extend(buscar_login(child))

    for shadow in node.get("shadowRoots", []):
        encontrados.extend(buscar_login(shadow))

    return encontrados


login_nodes = buscar_login(dom["root"])

print(
    f"Elementos de login encontrados por CDP: {len(login_nodes)}",
    flush=True
)

for nodo in login_nodes:
    print(
        nodo,
        flush=True
    )

