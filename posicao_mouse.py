import pyautogui
import time

print("🖱️ Mova o mouse para a posição desejada...")
time.sleep(5)  # tempo pra você posicionar o mouse

posicao = pyautogui.position()
print(f"📍 Posição do mouse: {posicao}")
