import pyautogui
import pygetwindow as gw
import time
import pyperclip

button_location = pyautogui.locateOnScreen('up_arrow.png')

if button_location is None:
    print('not found')
else:
    print('found')

pix = pyautogui.pixelMatchesColor(430, 100, (0, 120, 215))

while not pix:
    print('doesnt match')
    pix = pyautogui.pixelMatchesColor(430, 100, (0, 120, 215))
    if pix:
        break
    time.sleep(0.5)
    

print('matches')

