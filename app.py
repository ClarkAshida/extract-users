import pyautogui
import pandas as pd
from time import sleep

'''
Script para automação de extração de dados dos usuários do UNM
'''

#Interromper a execução do programa ao arrastar o mouse para o lado superior esquerdo da tela.
pyautogui.FailSafeException

#Clicar no usuário
pyautogui.click(92, 675)
#Abrir menu do usuário
pyautogui.rightClick(92, 675)
#Clicar em "Copy"
pyautogui.click(139, 753)

#Clicar no campo de "Username"
pyautogui.click(590, 140)
#Selecionar tudo
pyautogui.hotkey('ctrl', 'a')
#Copiar
pyautogui.hotkey('ctrl', 'c')

#Clicar na célula do Excel
pyautogui.click(1427,651)
pyautogui.hotkey('ctrl', 'v')

#Clicar no campo de "Real Name"
pyautogui.click(590, 167)

#Clicar no campo "Contatc (E-mail)"
pyautogui.click(590, 224)

#Clicar no campo "Description"
pyautogui.click(590, 257)

#Clicar em "User_group"
pyautogui.click(496, 114)
#Clicar na Coluna de grupos de usuário
pyautogui.click(534, 151)

#Clicar em "Activity Info"
pyautogui.click(790, 126)
#Tirar print da tela
screenshot = pyautogui.screenshot(region=(0, 0, 1920, 1080))
screenshot.save('activity_info.png')

#Clicar em "Close"
pyautogui.click(984, 627)

#Ir para baixo
pyautogui.press('down')
