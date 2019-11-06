import time
import pyautogui
import pyperclip

def copy_clipboard():
    pyautogui.hotkey('ctrl', 'c')
    time.sleep(.01)
    return pyperclip.paste()
