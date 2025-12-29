import pandas as pd
import pyautogui
import time
import os
from datetime import timedelta

# Tempos estimados
cicloPedido = 4.5
cicloItem = 49

# Configurações iniciais
base = pd.read_excel('uso-e-consumo.xlsx')
base.columns = base.columns.str.strip() # Remove espaços extras nos nomes das colunas

total_pedidos = base['Pedido de Compra '].nunique()
total_itens = base['qtd'].sum()

# Tempo total estimado inicial
tempo_restante = (total_itens * cicloItem) + (total_pedidos * cicloPedido)

def painel_status(pedido_num, pedido_atual, total_p, item_do_pedido, total_item_pedido, itens_globais, total_g, restante):
    # Formata os segundos para HH:MM:SS
    tempo_hms = str(timedelta(seconds=int(restante)))
    
    os.system('cls' if os.name == 'nt' else 'clear')
    print("========================================")
    print(f"         STATUS DA AUTOMAÇÃO          ")
    print("========================================")
    print(f"Pedido: {pedido_num} ({pedido_atual}/{total_p})")
    print(f"No Pedido Atual: Item {item_do_pedido} de {total_item_pedido}")
    print(f"Progresso Total: {itens_globais}/{total_g} itens")
    print(f"Tempo Restante Estimado: {tempo_hms}")
    print("========================================")

print("Iniciando em 5 segundos...")
time.sleep(5)

itens_processados_global = 0
pedido_contador = 0

for index, linha in base.iterrows():
    pedido_contador += 1
    pedido = str(linha['Pedido de Compra ']) 
    conta = str(linha['conta'])
    subconta = str(linha['sub-conta'])
    quantidade = int(linha['qtd']) 

    # --- INÍCIO DO PEDIDO ---
    pyautogui.hotkey('ctrl', 't')
    pyautogui.sleep(0.5)
    pyautogui.write(pedido)
    pyautogui.sleep(0.5)
    pyautogui.press('enter')
    pyautogui.sleep(3)
    
    # Deduz o tempo gasto na abertura do pedido
    tempo_restante -= cicloPedido

    for i in range(1, quantidade + 1):
        # Atualiza o painel no início de cada item
        painel_status(pedido, pedido_contador, total_pedidos, i, quantidade, itens_processados_global, total_itens, tempo_restante)

        pyautogui.press('tab')
        pyautogui.sleep(0.5)
        pyautogui.press('enter')
        pyautogui.sleep(5)

        for j in range(3): 
            pyautogui.press('tab')
            pyautogui.sleep(0.5)

        pyautogui.write(subconta)
        pyautogui.sleep(0.5)
        pyautogui.press('tab')
        pyautogui.sleep(0.5)
        pyautogui.write(conta)
        pyautogui.sleep(0.5)

        for k in range(2):
            pyautogui.hotkey('ctrl','tab')
            pyautogui.sleep(0.5)
        
        pyautogui.write("PC")
        pyautogui.sleep(0.5)
        pyautogui.hotkey('shift','tab')
        pyautogui.sleep(0.5)
        pyautogui.press('enter')

        # Espera longa
        pyautogui.sleep(35)

        pyautogui.press('tab')
        pyautogui.sleep(0.5)
        pyautogui.press('tab')
        pyautogui.sleep(0.5)
        pyautogui.press('enter')
        pyautogui.sleep(0.5)
        pyautogui.hotkey('shift','tab')
        pyautogui.sleep(0.5)
        pyautogui.press('down')
        pyautogui.sleep(0.5)

        # Atualização de contadores e tempo ao finalizar o item
        itens_processados_global += 1
        tempo_restante -= cicloItem

print("\nProcesso concluído com sucesso!")