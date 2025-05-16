import pyautogui
import time

# Give a brief pause to switch to the application where the hotkey works
time.sleep(1)

# Simulate pressing Windows + Control + C
pyautogui.hotkey('win', 'ctrl', 'c')

print("Simulated Windows + Control + C")
