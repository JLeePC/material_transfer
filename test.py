import pyautogui
import pygetwindow as gw
import time
import pyperclip
import csv
import os
import PIL
import getpass
from PIL import Image, ImageGrab, ImageFilter
import pytesseract
from file_routing import file_routing
from copy_clipboard import copy_clipboard
from go_up import go_up
from combine_workbook import combine_workbook

pytesseract.pytesseract.tesseract_cmd = r'C:\Users\jlee.NTPV\AppData\Local\Tesseract-OCR\tesseract.exe'

"""
part_list = []
stock_text_list = []
wip_text_list = []
released_text_list = []
on_order_text_list = []
stocking_text_list = []

custom_oem_psm_config = r'--oem 3 --psm 6'

arrow_location = pyautogui.locateOnScreen('arrow.png', region=(31,232,21,739))
arrow_top = arrow_location.top-2
arrow_height = arrow_location.height+4
arrow_bottom = arrow_top + arrow_height
                
#wip
wip_text_im = ImageGrab.grab(bbox =(251,arrow_top,311,arrow_bottom))
wip_size = tuple(4*x for x in wip_text_im.size)
wip_text_im = wip_text_im.resize(wip_size, Image.ANTIALIAS)
wip_text_im_gs = wip_text_im.convert('LA')
wip_text_string = pytesseract.image_to_string(wip_text_im_gs, config=custom_oem_psm_config)
wip_text_list.append(wip_text_string)

print(wip_text_string)

list_list = [stock_text_list,wip_text_list,released_text_list,on_order_text_list]
for i in range(len(list_list[1])):
    current_list = list_list[1]
    string = current_list[i]
    print(string)
    len_string = len(string)
    front_string = len_string - 3
    if "." not in string[front_string]:
        end_string = string[:front_string] + "." + string[front_string:]
        print(end_string)
        current_list.pop(i)
        current_list.insert(i,end_string)
print(wip_text_list)
"""


pyautogui.doubleClick(210,101)
job_no = copy_clipboard()

os.chdir(r"D:\MIsys Data")

with open('Master.csv', 'a', newline='') as csv_file:
    fieldnames = ['Time','Job number','+/-','Item','Part Number','Heat number','Amount','WIP Before','WIP After','Stock','Stock Amount','Released','On Order']
    csv_writer = csv.DictWriter(csv_file, fieldnames=fieldnames)

    for write in range(0,max_range):

        item_writer = str(item_list[write]).zfill(2)
        amount_writer = amount_transfer_list[write]
        part_writer = part_list[write]
        red_x_writer = red_x_list[write]
        job_writer = job_no
        transfer_time_writer = transfer_time_list[write]
        heat_number_writer = heat_list[write]
        stock_text_writer = stock_list[write]
        wip_text_writer = wip_text_list[write]
        released_text_writer = released_text_list[write]
        on_order_text_writer = on_order_text_list[write]
        wip_after_writer = wip_after_list[write]
        wip_writer = wip_list[write]

        if len(on_order_text_writer) == 0:
            on_order_text_writer = '0.00'
        
        if ',' in str(wip_text_writer):
            wip_text_writer = wip_text_writer.replace(",",".")
        if ',' in str(released_text_writer):
            released_text_writer = released_text_writer.replace(",",".")
        if ',' in str(on_order_text_writer):
            on_order_text_writer = on_order_text_writer.replace(",",".")
        if ',' in str(wip_after_writer):
            wip_after_writer = wip_after_writer.replace(",",".")

        amount_writer = round(amount_writer,3)
        stock_text_writer = round(stock_text_writer,3)
        wip_text_writer = round(float(wip_text_writer),3)
        released_text_writer = round(float(released_text_writer),3)
        on_order_text_writer = round(float(on_order_text_writer),3)
        wip_after_writer = round(wip_after_writer,3)
        
        csv_writer.writerow({'Time': transfer_time_writer, 'Job number': job_writer, '+/-': wip_writer,
                             'Item': item_writer, 'Part Number': part_writer,
                             'Heat number': heat_number_writer, 'Amount': amount_writer, 'WIP Before': wip_text_writer,
                             'WIP After': wip_after_writer, 'Stock': red_x_writer, 'Stock Amount': stock_text_writer,
                             'Released': released_text_writer, 'On Order': on_order_text_writer})

os.chdir(r"C:\Users\jlee.NTPV\Documents\GitHub\material_transfer")
