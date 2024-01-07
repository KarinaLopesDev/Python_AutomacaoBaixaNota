import keyboard
import webbrowser
import openpyxl
import pyautogui
import time
import pygetwindow as gw
import PySimpleGUI as sg

# lista para identificar os finalizados
finalizados = []

# Abrir uma URL em um navegador padrão
url = 'https://example/'
webbrowser.open(url)

# Esperar 3 segundos
time.sleep(3)

# Primeiro click
pyautogui.click(x=50, y=453)

# Esperar 2 segundos
time.sleep(2)

keyboard.press_and_release('tab')
keyboard.press_and_release('enter')
keyboard.press_and_release('tab')
keyboard.press_and_release('enter')

# Esperar 2 segundos
time.sleep(2)

# Segundo click
pyautogui.click(x=1041, y=838)

keyboard.press_and_release('tab')
keyboard.press_and_release('tab')
keyboard.press_and_release('tab')
keyboard.press_and_release('tab')

# Abrir o arquivo da planilha
workbook = openpyxl.load_workbook('C:\BaixaNota\exemplo.xlsx')

# Selecionar a aba desejada
sheet = workbook['ANA']

coluna_contagem = sheet['A']

linhas_preenchidas = 0

# Iterar sobre as células da coluna e contar as preenchidas
for cell in coluna_contagem:
    if cell.value is not None and cell.value != '':
        linhas_preenchidas += 1

for row_num in range(1, linhas_preenchidas + 1):
    nota = str(sheet.cell(row=row_num, column=1).value)

    # Simular a digitação do conteúdo da variável
    pyautogui.typewrite(nota)

    # Esperar 2 segundos
    time.sleep(3)

    keyboard.press_and_release('tab')
    keyboard.press_and_release('enter')

    # Esperar 3 segundo
    time.sleep(3)
    keyboard.press_and_release('tab')
    keyboard.press_and_release('enter')

    # Esperar 1 segundos
    time.sleep(1)

    keyboard.press_and_release('tab')
    keyboard.press_and_release('tab')
    keyboard.press_and_release('enter')

    # Esperar 1 segundo
    time.sleep(1)
    pyautogui.typewrite("limp")

    keyboard.press_and_release('enter')
    keyboard.press_and_release('tab')
    keyboard.press_and_release('tab')
    keyboard.press_and_release('tab')

    # Esperar 1 segundo
    time.sleep(1)
    keyboard.press_and_release('enter')

    time.sleep(3)
    keyboard.press_and_release('enter')

    # click na tela
    pyautogui.click(x=1041, y=838)

    time.sleep(3)

    # Encontrar janela do Chrome
    chrome_window = gw.getWindowsWithTitle('Google Chrome')[0]  # Obtem a primeira janela com o título 'Google Chrome'

    # Tornar a janela do Chrome ativa
    chrome_window.activate()

    pyautogui.click(x=420, y=396)
    pyautogui.click(x=420, y=396)

# Exibir mensagem
sg.popup('Planilha finalizada', title='Aviso')