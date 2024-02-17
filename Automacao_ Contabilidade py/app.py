import openpyxl
import pyperclip
import pyautogui
from time import sleep
# Entrar na planilha
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']
# Copiar infor de um campo e colar no seu campo correspondente
for linha in sheet_produtos.iter_rows(min_row =2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(139,360,duration=0.5)
    pyautogui.hotkey('ctrl','v')

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(131,451,duration=0.5)
    pyautogui.hotkey('ctrl','v')

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(133,589,duration=0.5)
    pyautogui.hotkey('ctrl','v')

    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(132,675,duration=0.5)
    pyautogui.hotkey('ctrl','v')

    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(133,765,duration=0.5)
    pyautogui.hotkey('ctrl','v')

    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(132,852,duration=0.5)
    pyautogui.hotkey('ctrl','v')

    # Botão Próximo
    pyautogui.click(162,906,duration=0.5)
    pyautogui.hotkey('ctrl','v')

    # Botão Preço
    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(136,397,duration=0.5)
    pyautogui.hotkey('ctrl','v')

    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(139,481,duration=0.5)
    pyautogui.hotkey('ctrl','v')

    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(143,571,duration=0.5)
    pyautogui.hotkey('ctrl','v')

    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(138,654,duration=0.5)
    pyautogui.hotkey('ctrl','v')


    tamanho = linha[10].value
    pyautogui.click(148,740,duration=0.5)

    if tamanho =='Pequeno': 
        pyautogui.click(168,770,duration=0.5)
    elif tamanho == 'Médio':
        pyautogui.click(170,796,duration=0.5)
    else:
        pyautogui.click(168,819,duration=0.5)

    # Se for "pequeno", clicar em uma pos
    # Se for "medio", clicar em uma pos
    # Se for "grande", clicar em uma pos


    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(130,825,duration=0.5)
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(162,906,duration=0.5)
    pyautogui.hotkey('ctrl','v')

    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(136,432,duration=0.5)
    pyautogui.hotkey('ctrl','v')

    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(142,513,duration=0.5)
    pyautogui.hotkey('ctrl','v')

    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(134,598,duration=0.5)
    pyautogui.hotkey('ctrl','v') 

    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(138,738,duration=0.5)
    pyautogui.hotkey('ctrl','v') 

    localizacao_armazem = linha[16].value 
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(135,826,duration=0.5)
    pyautogui.hotkey('ctrl','v')

    #Botão concluir
    pyautogui.click(161,883,duration=0.5)

    #Botão OK (alert)
    pyautogui.click(639,231,duration=0.5)
 
    # Add mais um produto
    pyautogui.click(477,639,duration=0.5)

    
          
# Repetir esses passos para outros campos até preencher tudo da página
# Clicar em próxima
# Repetir os mesmos passos e ir para a próxima página (pg 2)
# Repetir os mesmos passos e finalizar o cadastro daquele produto e clicar em concluir 
# Clicar em ok para finalizar o processo
# Clicar no ok mais uama vez na msg de confirmação de salvamento no BD
# Clicar em "Add mais um e repetir o processo até finalizar a planilha" 

# PyAutoGUI (Automação de clicks e e teclado)
# Openpyxl (leitura e automação de planilhas)
