import time
import pyautogui
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import re
import openpyxl
from openpyxl.styles import Font

chrome_options = Options()
chrome_options.add_argument("--start-maximized")

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

driver.get("https://www.creamy.com.br")

WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.TAG_NAME, "body")))

time.sleep(3)

for _ in range(3):
    pyautogui.press('tab')
    time.sleep(0.01)

pyautogui.press('enter')

time.sleep(1)

elementos = driver.find_elements(By.XPATH, '//*[starts-with(@id, "radix-") and contains(@id, "-trigger-")]/a')
if elementos:
    elementos[0].click()
else:
    driver.quit()

time.sleep(5)

contador_1 = 0

for _ in range(3):
    for _ in range(3):
        driver.execute_script("window.scrollBy(0, window.innerHeight * 1.1);")
        time.sleep(1)

    if contador_1 <= 1:
        try:
            elemento = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//div[@class="vtex-button__label flex items-center justify-center h-100 ph5 "]'))
            )
            elemento.click()
        except Exception:
            pass

    time.sleep(3)

    contador_1 += 1

elementos = driver.find_elements(By.XPATH, '//div[@class="flex-layout-row flex-layout-row--shelf__product-infos flex-layout-row--justify-start flex-layout-row--align-center"]')

if elementos:
    print(f"Encontrados {len(elementos)} elementos.")
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Relatório"
    
    headers = ["Nome/Descrição", "Preco/Desconto/Parcela", "Avaliação"]
    ws.append(headers)
    
    for cell in ws[1]:
        cell.font = Font(bold=True)
    
    for i, elemento in enumerate(elementos, 1):
        texto = elemento.text.replace('\n', ' ')
        
        partes = re.match(r'^(.*?) (R\$ .*?) (\(.*?\))$', texto)
        if partes:
            nome = partes.group(1)
            preco_desconto_parcela = partes.group(2)
            avaliacao = partes.group(3)
            
            if "%" in preco_desconto_parcela:
                preco_desconto_parcela = re.sub(r'(R\$ [\d,]+)', r'\1 por', preco_desconto_parcela, 1)
            
            texto_formatado = f"{nome} {preco_desconto_parcela} {avaliacao}"
            print(f"Elemento {i}: {texto_formatado}")
            
            ws.append([nome, preco_desconto_parcela, avaliacao])
    
    for column in ws.columns:
        max_length = 0
        column = list(column)
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[column[0].column_letter].width = adjusted_width
    
    wb.save("Relatório.xlsx")
    print("Arquivo Relatório.xlsx salvo com sucesso.")
else:
    print("Nenhum elemento encontrado.")

time.sleep(1)

print(f"Total de elementos encontrados: {len(elementos)}")

driver.quit()
