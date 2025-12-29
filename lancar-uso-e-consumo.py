import pandas as pd
import pyautogui
import time
import openpyxl

# Configurações iniciais
# Pausa de 1 segundo entre cada comando do pyautogui para segurança
base = pd.read_excel('uso-e-consumo.xlsx')
print("iniciando em 5 segundos")
pyautogui.sleep(5)


for index, linha in base.iterrows():
    pedido = linha['Pedido de Compra ']
    conta = linha['conta']
    subconta = linha['sub-conta']

    quantidade = int(linha['qtd']) # Convertemos para inteiro para garantir o uso no range
    
    print(f"--- Iniciando Pedido: {pedido} ---")

    pyautogui.hotkey('alt', 't')

    pyautogui.sleep(1.5)

    pyautogui.write(pedido)

    pyautogui.sleep(1.5)

    pyautogui.press('enter')
    
    pyautogui.sleep(2)

    print(f"{quantidade} Itens no pedido")    
    # Criar um loop que roda 'quantidade' vezes
    for i in range(quantidade):
        
        pyautogui.press('tab')

        pyautogui.sleep(2)

        pyautogui.press('enter')

        pyautogui.sleep(5)

        for i in range(3)
            pyautogui.press('tab')
            pyautogui.sleep(0.5)

        pyautogui.write(subconta)

        pyautogui.sleep(1)

        pyautogui.write(conta)

        pyautogui.sleep(1)

        for i in range(3)
            pyautogui.hotkey('ctrl','tab')
            pyautogui.sleep(0.5)

        
        pyautogui.write("PC")
        pyautogui.sleep(1)

        pyautogui.hotkey('shift','tab')

        pyautogui.sleep(1)

        pyautogui.press('enter')

        pyautogui.sleep(10)

        pyautogui.press('tab')
        pyautogui.sleep(0.5)
        pyautogui.press('tab')
        pyautogui.sleep(0.5)

        pyautogui.press('enter')

        pyautogui.sleep(1)

        pyautogui.hotkey('shift','tab')

        pyautogui.sleep(1)