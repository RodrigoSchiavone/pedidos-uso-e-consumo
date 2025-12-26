import pandas as pd
import pyautogui
import time
import openpyxl

# Configurações iniciais
# Pausa de 1 segundo entre cada comando do pyautogui para segurança
pyautogui.PAUSE = 1.0 

def executar_automacao():
    # 1. Ler o arquivo Excel
    df = pd.read_excel('uso-e-consumo.xlsx')
    
    # Limpar nomes das colunas (remover espaços extras se houver)
    df.columns = df.columns.str.strip()

    # Agrupar por pedido para garantir que o Alt+T só aconteça uma vez por pedido
    # Assumindo que o pedido pode ter várias linhas no Excel
    pedidos = df.groupby('Pedido de Compra')

    print("A automação começará em 5 segundos. Abra a janela do software destino.")
    time.sleep(1)

    for num_pedido, grupo in pedidos:
        # --- INÍCIO DO LOOP DO PEDIDO ---
        # Apertar Alt+T
        pyautogui.hotkey('alt', 't')
        time.sleep(1)
        
        # Escrever o número do pedido de compra
        # Removendo pontos se o Excel ler como string formatada
        pedido_str = str(num_pedido).replace('.', '')
        time.sleep(1)
        pyautogui.write(pedido_str)
        time.sleep(1)
        pyautogui.press('enter') 
        time.sleep(1) # Espera carregar o pedido

        # --- SEGUNDO LOOP: PARA CADA SKU (LINHA) DO PEDIDO ---
        for index, linha in grupo.iterrows():
            # A coluna 'qtd' define quantas vezes o processo se repete para este SKU
            repeticoes_sku = int(linha['qtd'])
            time.sleep(1)
            
            sub_conta = str(linha['sub-conta'])
            time.sleep(1)
            conta = str(linha['conta'])
            time.sleep(1)

            for _ in range(repeticoes_sku):
                # Sequência de comandos solicitada
                pyautogui.press('tab')
                pyautogui.press('enter')
                time.sleep(1)
                
                for _ in range(3):
                    pyautogui.press('tab')
                    time.sleep(1)
                
                pyautogui.write(sub_conta)
                time.sleep(1)
                pyautogui.press('tab')
                time.sleep(1)
                pyautogui.write(conta)
                time.sleep(1)
                
                pyautogui.hotkey('ctrl', 'tab')
                time.sleep(1)
                pyautogui.hotkey('ctrl', 'tab')
                time.sleep(1)
                
                pyautogui.write('PC')
                time.sleep(1)
                
                pyautogui.hotkey('shift', 'tab')
                time.sleep(1)
                pyautogui.press('enter')
                time.sleep(1)

                # Espera longa solicitada após processar o SKU
                print(f"Aguardando 20 segundos (Pedido: {num_pedido})...")
                time.sleep(20)

                pyautogui.press('tab')
                time.sleep(1)
                pyautogui.press('tab')
                time.sleep(1)
                pyautogui.press('enter')
                time.sleep(1)
                # Preparar para o próximo SKU ou próxima repetição
                pyautogui.hotkey('shift', 'tab')
                time.sleep(1)
                pyautogui.press('down')
                time.sleep(1)

    print("Processo concluído!")

if __name__ == "__main__":
    executar_automacao()