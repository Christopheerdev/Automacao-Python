import pyautogui
import time

print("Posicione o mouse... capturando em 4 segundos.")
time.sleep(4)

x, y = pyautogui.position()
print(f"Posição atual do ponteiro: X={x}, Y={y}")
