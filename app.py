import openpyxl
import pyperclip
import pyautogui
from time import sleep

# Entrar na Planilha
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']
# Copiar informação de um campo e colar no seu campo correspondente
for linha in sheet_produtos.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(319,336,duration=1)
    pyautogui.hotkey('ctrl', 'v')

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(318,422,duration=1)
    pyautogui.hotkey('ctrl', 'v')

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(322,557,duration=1)
    pyautogui.hotkey('ctrl', 'v')

    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(320,650,duration=1)
    pyautogui.hotkey('ctrl', 'v')

    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(320,730,duration=1)
    pyautogui.hotkey('ctrl', 'v')

    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(320,816,duration=1)
    pyautogui.hotkey('ctrl', 'v')

    pyautogui.click(350,873, duration=1)
    sleep(3)

    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(326,361,duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(325,446,duration=1)
    pyautogui.hotkey('ctrl', 'v')
    
    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(328,533,duration=1)
    pyautogui.hotkey('ctrl', 'v')

    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(325,618,duration=1)
    pyautogui.hotkey('ctrl', 'v')

    tamanho = linha[10].value
    pyautogui.click(329,704,duration=1)

    if tamanho == 'Pequeno':
        pyautogui.click(351,739,duration=1)
    elif tamanho == 'Médio':
        pyautogui.click(350,760,duration=1)
    else:
        pyautogui.click(350,783,duration=1)
    
    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(333,791,duration=1)
    pyautogui.hotkey('ctrl', 'v')

    pyautogui.click(331,850,duration=1)
    sleep(3)

    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(320,379,duration=1)
    pyautogui.hotkey('ctrl', 'v')

    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(319,466,duration=1)
    pyautogui.hotkey('ctrl', 'v')

    observacoes = linha[14].value
    pyperclip.copy(observacoes)
    pyautogui.click(319,559,duration=1)
    pyautogui.hotkey('ctrl', 'v')

    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(317,685,duration=1)
    pyautogui.hotkey('ctrl', 'v')

    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(319,772,duration=1)
    pyautogui.hotkey('ctrl', 'v')

    pyautogui.click(342,829,duration=1)

    pyautogui.click(1132,191,duration=1)

    pyautogui.click(953,607,duration=1)

# Repetir esses passos para outros campos até preencher campos daquela página
# Clicar em próxima
# Repetir os mesmos passos e ir para a próxima pagina(página 2)
# Repetir os mesmos passo e finalizar o cadastro daquele produto e clicar em concluir
# Clicar em ok, para finalizar o processo
# Clicar no ok mais uma vez na mensagem de confirmação de salvamento no banco de dados
# Clicar em "adicionar mais um e repetir o processo ate finalizar a planilha"
