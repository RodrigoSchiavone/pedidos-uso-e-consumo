import pandas as pd
import pyautogui
import time
import openpyxl

# Configurações iniciais
base = pd.read_excel('uso-e-consumo.xlsx')
print("iniciando em 5 segundos")
pyautogui.sleep(5)

for index, linha in base.iterrows():
    # Convertendo para string para garantir que o pyautogui consiga digitar
    pedido = str(linha['Pedido de Compra ']) 
    conta = str(linha['conta'])
    subconta = str(linha['sub-conta'])
    quantidade = int(linha['qtd']) 
    
    print(f"--- Iniciando Pedido: {pedido} ---")

    pyautogui.hotkey('ctrl', 't')
    pyautogui.sleep(1)
    pyautogui.write(pedido)
    pyautogui.sleep(1)
    pyautogui.press('enter')
    pyautogui.sleep(3)

    print(f"{quantidade} Itens no pedido")    
    
    for i in range(quantidade):
        pyautogui.press('tab')
        pyautogui.sleep(2)
        pyautogui.press('enter')
        pyautogui.sleep(5)

        # CORREÇÃO: Adicionado ":" ao final do for
        for j in range(3): 
            pyautogui.press('tab')
            pyautogui.sleep(1)

        pyautogui.write(subconta)
        pyautogui.sleep(1)
        pyautogui.press('tab')
        pyautogui.sleep(1)
        pyautogui.write(conta)
        pyautogui.sleep(1)

        # CORREÇÃO: Adicionado ":" ao final do for
        # Mudei a variável para 'k' para não confundir com o loop externo
        for k in range(3):
            pyautogui.hotkey('ctrl','tab')
            pyautogui.sleep(1)
        
        pyautogui.write("PC")
        pyautogui.sleep(1)
        pyautogui.hotkey('shift','tab')
        pyautogui.sleep(1)
        pyautogui.press('enter')

        # Espera longa para processamento do sistema
        pyautogui.sleep(10)

        pyautogui.press('tab')
        pyautogui.sleep(1)
        pyautogui.press('tab')
        pyautogui.sleep(1)
        pyautogui.press('enter')
        pyautogui.sleep(1)
        pyautogui.hotkey('shift','tab')
        pyautogui.sleep(1)
        pyautogui.press('down')
        pyautogui.sleep(1)