# ler dados da planilha escolhida
# inserir cada célula em determinados campos
# No lugar das unidades "1"; mapeamento de cursor do mause.
# Em "vendas", colocar o diretorio do arquivo em .xlsx e o nome da planilha.
import openpyxl

workbook = openpyxl.load_workbook('#Planilha')
vendas_sheet = workbook['vendas']

for linha in vendas_sheet.iter_rows(min_rows=2):
    #nome
    pyautogui.click(1,1,duration=1.5)
    pyautogui.write(linha[0].value)
    #produto
    pyautogui.click(1,1,duration=1.5)
    pyautogui.write(linha[1].value)
    #quantidade
    pyautogui.click(1,1,duration=1.5)
    pyautogui.write(linha[2].value)
    #categoria
    pyautogui.click(1,1,duration=1.5)
    pyautogui.write(linha[3].value)
    pyautogui.click(1,1,duration=1.5)
    pyautogui.click(1,1,duration=1.5)
    
