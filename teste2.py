import pyautogui
import pandas as pd
from time import sleep

# Interromper a execução do programa ao arrastar o mouse para o lado superior esquerdo da tela.
pyautogui.FailSafeException

# Definir as posições para os campos
campos = {
    'Username': (978, 140),
    'Real Name': (978, 167),
    'E-mail': (978, 224),
    'Description': (978, 257),
    #'User Group': (496, 114),
}

for i in range(1):
    # Clicar no usuário
    pyautogui.click(92, 675)
    # Abrir menu do usuário
    pyautogui.rightClick(92, 675)
    # Clicar em "Copy"
    pyautogui.click(139, 753)
    #Clicar no campo de "Username"
    pyautogui.click(590, 140)

    # Função para copiar dados e colar no Excel
    def copiar_e_colar(campo, coluna_excel):
        # Clicar no campo
        pyautogui.click(campos[campo])
        pyautogui.write('0')

        # Selecionar tudo e copiar
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.hotkey('ctrl', 'c')

        # Esperar um pouco para garantir que os dados foram copiados
        sleep(0.5)

        # Clicar na célula do Excel na coluna correspondente
        pyautogui.click(1427,651)
        for _ in range(coluna_excel):
            pyautogui.press('right')
        pyautogui.hotkey('ctrl', 'v')
        
    # Copiar e colar os dados
    for i, campo in enumerate(campos.keys()):
        copiar_e_colar(campo, i)
        if i >= 3:
            pyautogui.keyDown('down')  # Mover para a próxima linha

    # Clicar em "Close" para fechar a janela do usuário
    pyautogui.click(984, 627)
    # Clicar em "Activity Info"
    pyautogui.click(787, 125)
    # Tirar print da tela
    screenshot = pyautogui.screenshot(region=(496, 156, 1065, 31))
    # Clicar na primeira célula
    pyautogui.click(1427,651)
    # Apertar "Right" 4 vezes
    for i in range(4):
        pyautogui.press('right')
    # Colar o print
    pyautogui.hotkey('ctrl', 'v')

    # Clicar no usuário
    pyautogui.click(92, 675)
    # Ir para baixo para selecionar o próximo usuário
    pyautogui.press('down')