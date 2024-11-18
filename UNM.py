import os
import pyautogui
import pyperclip
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from time import sleep

# Interromper a execução do programa ao arrastar o mouse para o lado superior esquerdo da tela.
pyautogui.FAILSAFE = True

# Definir as posições para os campos
campos = {
    'Username': (978, 140),
    'Real Name': (978, 167),
    'E-mail': (978, 224),
    'Description': (978, 257),
}

# Criar e configurar o workbook para salvar os dados e a imagem
wb = Workbook()
ws = wb.active
ws.append(['Username', 'Real Name', 'E-mail', 'Description', 'Activity Info Screenshot'])  # Cabeçalhos

# Função para copiar e colar dados no Excel
def copiar_e_colar(campo, linha_excel, coluna_excel):
    # Clicar no campo
    pyautogui.click(campos[campo])
    pyautogui.write('0')
    
    # Selecionar e copiar
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'c')
    
    # Esperar um pouco para garantir que os dados foram copiados
    sleep(0.5)

    # Colar o valor no Excel (usando pyperclip para pegar o texto copiado)
    ws.cell(row=linha_excel, column=coluna_excel + 1, value=pyperclip.paste())

# Inicializar a linha onde serão colados os dados
linha_excel = 2

# Executar automação para copiar campos e colar no Excel
for i in range(2):  # Ajuste o range para o número de usuários a serem extraídos
    # Clicar no usuário e abrir o menu
    pyautogui.click(92, 675)
    pyautogui.rightClick(92, 675)
    pyautogui.click(139, 753)

    # Copiar e colar os dados de cada campo
    for coluna_excel, campo in enumerate(campos.keys()):
        copiar_e_colar(campo, linha_excel, coluna_excel)

    # Clicar em "Close" para fechar a janela do usuário
    pyautogui.click(984, 627)
    # Clicar em "Activity Info"
    pyautogui.click(787, 125)

    # Tirar um screenshot da região da tela desejada
    screenshot_path = f"activity_info_{linha_excel}.png"  # Nome único para cada iteração
    screenshot = pyautogui.screenshot(region=(496, 156, 130, 31))
    screenshot.save(screenshot_path)  # Salva temporariamente a imagem

    # Inserir a imagem na coluna correta (E para seguir o exemplo)
    img = Image(screenshot_path)
    ws.add_image(img, f"E{linha_excel}")  # Ajusta conforme a coluna de destino

   # Clicar no usuário
    pyautogui.click(92, 675)

    # Ir para baixo para selecionar o próximo usuário
    pyautogui.press('down')

    # Avançar para a próxima linha na planilha para o próximo usuário
    linha_excel += 1

# Salvar a planilha após todos os dados e imagens serem inseridos
wb.save("dados_usuarios.xlsx")
