import pyautogui
import time
import os


def go_up():
    button_location = pyautogui.locateOnScreen('up_arrow.png', region=(875,207,29,42))
    top_location = pyautogui.locateOnScreen('top_arrow.png', region=(875,207,29,100))

    if button_location is not None:
        if top_location is None:
            print("GOING UP")
            pyautogui.click(892, 250, clicks=4, interval=0.05)
            pyautogui.click(890,227)
