import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# Caminho para o GeckoDriver
gecko_driver_path = r"C:\Users\Julliana\Desktop\Projetos_ Freelancer\Freelancer Contabilidade\dist\geckodriver.exe"

# Configuração do Selenium para o Firefox
firefox_options = Options()
firefox_options.add_argument("--start-maximized")  # Inicia o navegador maximizado
driver = webdriver.Firefox(options=firefox_options)

# Abre a página da web
url = "https://cadastro-produtos-devaprender.netlify.app/index.html"
driver.get(url)

# Espera para garantir que a página seja totalmente carregada
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "product_name")))

# Entrar na planilha
workbook = openpyxl.load_workbook(r"C:\Users\Julliana\Desktop\Projetos_ Freelancer\Freelancer Contabilidade\dist\produtos_ficticios.xlsx")
sheet_produtos = workbook['Produtos']

# Função para esperar até que um elemento esteja visível e clicável
def wait_until_clickable(element):
    try:
        return WebDriverWait(driver, 20).until(EC.element_to_be_clickable(element))
    except TimeoutException:
        print(f"Elemento {element} não encontrado ou não clicável dentro do tempo especificado.")
        raise

# Copiar informações de um campo e colar no seu campo correspondente
for linha in sheet_produtos.iter_rows(min_row=2):
    wait_until_clickable((By.ID, "product_name")).send_keys(linha[0].value)
    wait_until_clickable((By.ID, "description")).send_keys(linha[1].value)
    wait_until_clickable((By.ID, "category")).send_keys(linha[2].value)
    wait_until_clickable((By.ID, "product_code")).send_keys(linha[3].value)
    wait_until_clickable((By.ID, "weight")).send_keys(linha[4].value)
    wait_until_clickable((By.ID, "dimensions")).send_keys(linha[5].value)
    
    # Botão Próximo
    wait_until_clickable((By.CSS_SELECTOR, ".btn.btn-primary.me-2")).click()

    # Copiar informações de um campo e colar no seu campo correspondente
    wait_until_clickable((By.ID, "price")).send_keys(linha[6].value)
    wait_until_clickable((By.ID, "stock")).send_keys(linha[7].value)
    wait_until_clickable((By.ID, "expiry_date")).send_keys(linha[8].value)
    wait_until_clickable((By.ID, "color")).send_keys(linha[9].value)
    
    # Mapeamento dos valores de tamanho para seus seletores CSS
    tamanho_selectors = {
        'Pequeno': "#size > option:nth-child(1)",
        'Médio': "#size > option:nth-child(2)",
        'Grande': "#size > option:nth-child(3)",
    }

    # Selecionar a opção com base no tamanho
    tamanho = linha[10].value
    # Clicar no campo de seleção de tamanho
    wait_until_clickable((By.CSS_SELECTOR, "#size")).click()
    
    # Esperar até que o seletor específico do tamanho esteja presente
    wait_until_clickable((By.CSS_SELECTOR, tamanho_selectors[tamanho]))

    # Selecionar a opção com base no tamanho
    if tamanho in tamanho_selectors:
        driver.find_element(By.CSS_SELECTOR, tamanho_selectors[tamanho]).click()

    # Copiar informações de um campo e colar no seu campo correspondente
    wait_until_clickable((By.ID, "material")).send_keys(linha[10].value)

    # Botão Próximo
    wait_until_clickable((By.CSS_SELECTOR, ".btn.btn-primary.me-2")).click()

    wait_until_clickable((By.ID, "manufacturer")).send_keys(linha[11].value)
    wait_until_clickable((By.ID, "country")).send_keys(linha[12].value)
    wait_until_clickable((By.ID, "remarks")).send_keys(linha[13].value)
    wait_until_clickable((By.ID, "barcode")).send_keys(linha[14].value)
    wait_until_clickable((By.ID, "warehouse_location")).send_keys(linha[15].value)
    
    # Clicar no botão "Concluir"
    wait_until_clickable((By.CSS_SELECTOR, ".btn.btn-primary.me-2")).click()

    # Aguardar o alerta estar presente
    alert = WebDriverWait(driver, 10).until(EC.alert_is_present())

    # Aceitar o alerta (clicar em OK)
    alert.accept()

    # Adicionar mais um produto
    wait_until_clickable((By.CSS_SELECTOR, ".btn.btn-primary")).click()

# Fechar o navegador ao final
driver.quit()









































    
          
