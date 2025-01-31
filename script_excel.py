import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from PyPDF2 import PdfReader
from selenium.common.exceptions import NoSuchElementException
import os
import time


file_path = "Lotes_Export.xlsx"; ## Nome do arquivo que está sendo utilizado para pegar os dados.
folder_path = 'D:\\Projetos\\script_python\\pdf'
iterator_row = 0;
linhas_pdf = [];

try:
    df = pd.read_excel(file_path, engine="openpyxl") ## Utilizando a propriedade (nrows=) você consegue selecionar a quantidade de linhas que vai ler no excel
except Exception as e:
    print("Erro ao ler o arquivo:", e)

service = Service(ChromeDriverManager().install())
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--headless")
chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36")
prefs = {
    "download.default_directory": "D:\\Projetos\\script_python\\pdf",
    "download.prompt_for_download": False,
    "plugins.always_open_pdf_externally": True
}
chrome_options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(service=service, options=chrome_options)

def extract_text_from_pdf(pdf_path):
    with open(pdf_path, 'rb') as file:
        reader = PdfReader(file)
        text = ""
        
        # Itera por todas as páginas e extrai o texto
        for page in reader.pages:
            text += page.extract_text()
    
    return text

# Função para pegar linhas específicas
def get_specific_lines(pdf_path, line_numbers):
    try:
        text = extract_text_from_pdf(pdf_path)
        lines = text.split('\n')
        specific_lines = [lines[i] for i in line_numbers if i < len(lines)]
        return specific_lines
    except Exception as e:
        print(f"Erro ao processar o PDF: {e}")
        return []

def buscar_com_selenium(numero_inscricao):
    try:
        driver.get("https://portal.cidadao.conam.com.br/cacapava/certidao_venal.php")

        search_box = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.NAME, "frm_inscricao_imobiliaria"))
        )   
        numero_inscricao_str = str(numero_inscricao)
        search_box.send_keys("0" + numero_inscricao_str)
        search_box.send_keys(Keys.RETURN)

        time.sleep(1)
        try:
            div_element = driver.find_element(By.CLASS_NAME, "form-actions")
            print_button = div_element.find_element(By.TAG_NAME, "button")
            print_button.click()
        except NoSuchElementException:
            print(f"Erro: Não foi possível encontrar o botão de download para o número de inscrição {numero_inscricao}. Pulando a iteração.")
            return []

        time.sleep(0.5)

        # Exemplo: obter as linhas 2 e 5
        line_numbers = [11, 10, 14, 15, 9, 16, 13]

        files = [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]
        file_paths = [os.path.join(folder_path, f) for f in files]
        latest_pdf = max(file_paths, key=os.path.getmtime)

        if os.path.exists(latest_pdf):
            print(f"Arquivo {latest_pdf} existe e pode ser processado")
        else:
            print(f"Erro: {latest_pdf} não foi baixado corretamente.")

        lines = get_specific_lines(latest_pdf, line_numbers)

        return lines
    except Exception as e:
        print(f"Erro ao buscar informações para {numero_inscricao}: {e}")
        return []


for linha_excel in df.iterrows():
    init_fill_column = 3;
    return_lines = buscar_com_selenium(df["Insc_Imob"].iloc[iterator_row]);

    if not return_lines:
        print(f"Sem resultados para o número de inscrição {df['Insc_Imob'].iloc[iterator_row]}. Pulando a iteração.")
        continue  # Pula para a próxima iteração

    for line in return_lines:
        df[df.columns[init_fill_column]] = df[df.columns[init_fill_column]].astype(object)
        df.loc[iterator_row, df.columns[init_fill_column]] = str(line)
        init_fill_column += 1;
    
    if(iterator_row == len(df)):
        break;

    if iterator_row % 10 == 0:
        driver.quit()
        driver = webdriver.Chrome(service=service, options=chrome_options)

    if iterator_row % 5000 == 0:
        # 3. Salvar no excel a cada 5 mil linhas
        print(f"O excel foi salvo na linha {iterator_row}!")
        df.to_excel("Lotes_Export_dados.xlsx", index=False) ## Nome da saída do arquivo com os dados nas colunas de informações adquiridas do PDF

    iterator_row += 1;

# 3. Salvar de volta no Excel
df.to_excel("Lotes_Export_dados.xlsx", index=False) ## Nome da saída do arquivo com os dados nas colunas de informações adquiridas do PDF