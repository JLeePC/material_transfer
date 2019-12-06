#! python3
# MISys semi-auto transfer

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

username = "NTPV\jlee"
password = getpass.getpass("Enter Password: ")
path = r'\\NTPV-SERVER2008\ntpv data'

mount_command = "net use {} /user:{} {} /persistent:no".format(path,username,password)
os.system(mount_command)

pytesseract.pytesseract.tesseract_cmd = r'C:\Users\jlee.NTPV\AppData\Local\Tesseract-OCR\tesseract.exe'

class EmptyInput(Exception):
    pass

def remove_values_from_list(the_list, val):
   return [value for value in the_list if value != val]

ask_for_next_job = False
pyautogui.PAUSE = 0.1
print('Ctrl_c to quit.')

try:
    while True:
        pyautogui.PAUSE = 0.1
        pwht = False
        # Go to next job
        if ask_for_next_job is True:
            pyautogui.click(1423,15)
            next_job = input('\nEnter next job: ')
            if len(next_job) <= 4:
                next_job = next_job.zfill(2)
                pyautogui.click(214,102)
                back = len(next_job)
                for number_of_backspace in range(back):
                    pyautogui.typewrite(['\b'])
                    time.sleep(0.05)
                pyautogui.typewrite(str(next_job))
                pyautogui.typewrite(['tab'])
            else:
                pyautogui.doubleClick(214,102)
                pyautogui.typewrite(str(next_job))
                pyautogui.typewrite(['tab'])

            time.sleep(0.5)
            pyautogui.click(41,154) # header
            time.sleep(1)
            waiting = True
            while waiting:
                released_by = pyautogui.locateOnScreen('released_by.png', region=(370,250,130,50))
                if released_by is not None:
                    waiting = False
                    break
            pyautogui.doubleClick(214,102)
            next_job = copy_clipboard()
            time.sleep(0.1)
            pyautogui.doubleClick(228,313) # job no.

            job_no = copy_clipboard()
            
            job = False

            while not job:
                pyautogui.doubleClick(228,313) # job no.
                job_no = copy_clipboard()
                if job_no ==  next_job.upper():
                    job = True
                    break
                time.sleep(0.25)
                
        total_start = time.time()
        
        time.sleep(1)
        released_button = pyautogui.locateOnScreen('released.png', region=(614,86,170,29))
        on_hold_button = pyautogui.locateOnScreen('on_hold.png', region=(712,115,91,24))
        time.sleep(0.25)

        if released_button is not None and on_hold_button is None:

            pyautogui.click(150,150)
            time.sleep(0.5)
            pyautogui.click(351,243)
            pyautogui.typewrite(['up'])
            pyautogui.typewrite(['up'])
            print('\nLooking for side bar.')
            button_location = pyautogui.locateOnScreen('up_arrow.png', region=(875,207,29,42))
            time.sleep(1)
            
            if button_location is None:
                print('Up arrow not found')
            else:
                print('Up arrow found')
                go_up()
                pyautogui.click(351,243)

            labor_lines = False
                
            if button_location is not None:
                print('\nLooking for LABOR.')

                part = ImageGrab.grab(bbox =(81,233,193,970))
                new_size = tuple(4*x for x in part.size)
                part = part.resize(new_size, Image.ANTIALIAS)
                part_bl = part.filter(ImageFilter.GaussianBlur(radius = 1))
                part_gs = part_bl.convert('LA')
                custom_oem_psm_config = r'--oem 3 --psm 6'
                part_string = pytesseract.image_to_string(part_gs, config=custom_oem_psm_config)
                part_string_list = part_string.split("\n")

                line_number = ImageGrab.grab(bbox =(52,234,80,967))
                new_size = tuple(4*x for x in line_number.size)
                line_number = line_number.resize(new_size, Image.ANTIALIAS)
                line_number_bl = line_number.filter(ImageFilter.GaussianBlur(radius = 1))
                line_number_gs = line_number_bl.convert('LA')
                custom_oem_psm_config = r'--oem 3 --psm 6'
                line_number_string = pytesseract.image_to_string(line_number_gs, config=custom_oem_psm_config)
                line_number_string_list = line_number_string.split("\n")

                line_number_string_list = remove_values_from_list(line_number_string_list, "")

                if 'PWHT' in part_string:
                    pwht = True
                    print('There is PWHT in this job.')
                
                if 'LABOR' in part_string:
                    page = 1
                    labor_lines = True
                    print('There is LABOR in this job.')

                page = 2
                labor_loop = True
                while labor_loop:
                    pyautogui.click(891,961, clicks=35, interval=0.05) # down arrow
                    time.sleep(0.25)
                    part = ImageGrab.grab(bbox =(81,233,193,970))
                    new_size = tuple(4*x for x in part.size)
                    part = part.resize(new_size, Image.ANTIALIAS)
                    part_bl = part.filter(ImageFilter.GaussianBlur(radius = 1))
                    part_gs = part_bl.convert('LA')
                    custom_oem_psm_config = r'--oem 3 --psm 6'
                    part_string2 = pytesseract.image_to_string(part_gs, config=custom_oem_psm_config)
                    part_string_list2 = part_string2.split("\n")

                    line_number = ImageGrab.grab(bbox =(52,234,80,967))
                    new_size = tuple(4*x for x in line_number.size)
                    line_number = line_number.resize(new_size, Image.ANTIALIAS)
                    line_number_bl = line_number.filter(ImageFilter.GaussianBlur(radius = 1))
                    line_number_gs = line_number_bl.convert('LA')
                    custom_oem_psm_config = r'--oem 3 --psm 6'
                    line_number_string2 = pytesseract.image_to_string(line_number_gs, config=custom_oem_psm_config)
                    line_number_string_list2 = line_number_string2.split("\n")

                    line_number_string_list2 = remove_values_from_list(line_number_string_list2, "")
                    
                    bottom_arrow = pyautogui.locateOnScreen('bottom_arrow.png', region=(880,940,30,40))
                    
                    if 'LABOR' in part_string2:
                        labor_lines = True
                        print('There is LABOR in this job.')

                    if 'PWHT' in part_string2:
                        pwht = True
                        print('There is PWHT in this job.')

                    line_number_string_list2 = line_number_string2.split("\n")

                    if line_number_string_list2[0] in line_number_string_list[len(line_number_string_list)-1]:
                        print("At Bottom")
                        labor_loop = False
                    else:
                        part_string_list = part_string_list + part_string_list2
                        line_number_string_list = line_number_string_list + line_number_string_list2

            else:
                print('\nLooking for LABOR.')

                part = ImageGrab.grab(bbox =(81,233,193,970))
                new_size = tuple(4*x for x in part.size)
                part = part.resize(new_size, Image.ANTIALIAS)
                part_bl = part.filter(ImageFilter.GaussianBlur(radius = 1))
                part_gs = part_bl.convert('LA')
                custom_oem_psm_config = r'--oem 3 --psm 6'
                part_string = pytesseract.image_to_string(part_gs, config=custom_oem_psm_config)
                part_string_list = part_string.split("\n")

                line_number = ImageGrab.grab(bbox =(52,234,80,967))
                new_size = tuple(4*x for x in line_number.size)
                line_number = line_number.resize(new_size, Image.ANTIALIAS)
                line_number_bl = line_number.filter(ImageFilter.GaussianBlur(radius = 1))
                line_number_gs = line_number_bl.convert('LA')
                custom_oem_psm_config = r'--oem 3 --psm 6'
                line_number_string = pytesseract.image_to_string(line_number_gs, config=custom_oem_psm_config)
                line_number_string_list = line_number_string.split("\n")

                line_number_string_list = remove_values_from_list(line_number_string_list, "")
                
                if 'LABOR' in part_string:
                    labor_lines = True
                    print('There is LABOR in this job.')
                else:
                    labor_lines = False
                    print('There is no LABOR in this job.')
                if 'PWHT' in part_string_list:
                    pwht = True
                    print('There is PWHT in this job.')
                
            pyautogui.PAUSE = 0.1

            string_list_len = len(part_string_list)

            go_up()

            if pwht:
                pyautogui.doubleClick(130,243)
                current_part_no = copy_clipboard()
                pyautogui.PAUSE = 0.1
                if 'PWHT' in part_string_list[0]:
                    pyautogui.click(920,290)
                    time.sleep(0.25)
                    pyautogui.typewrite(['tab'])
                    pyautogui.hotkey('shift','tab')
                    time.sleep(0.25)
                    current_part_no9 = copy_clipboard()
                    up_count = 0
                    go_up_loop = True
                    while go_up_loop:
                        pyautogui.typewrite(['up'])
                        current_part_no9 = copy_clipboard()
                        if 'LABOR' not in current_part_no9:
                            go_up_loop = False
                            break
                        up_count = up_count + 1
                    for down in range(up_count+1):
                        pyautogui.typewrite(['down'])
                    pyautogui.click(922,247,clicks=up_count,interval=0.025)
                if 'PWHT' in part_string_list[string_list_len - 1]:
                    if button_location:
                        go_up()
                        pyautogui.doubleClick(130,243)

                    else:
                        pyautogui.doubleClick(130,243)
                        down_loop = True
                        while down_loop:
                            pyautogui.PAUSE = 0.03
                            im1 = pyautogui.screenshot()
                            for i in range(10):
                                pyautogui.typewrite(['down'])
                            im2 = pyautogui.screenshot()
                            if im1 == im2:
                                down_loop = False
                                break
                    time.sleep(0.5)
                    current_part_no = copy_clipboard()
                    time.sleep(0.1)
                    if 'PWHT' in current_part_no:
                        pyautogui.typewrite(['tab'])
                        pyautogui.hotkey('shift','tab')
                        time.sleep(0.25)
                        current_part_no9 = copy_clipboard()
                        up_count = 0
                        go_up = True
                        while go_up:
                            pyautogui.PAUSE = 0.1
                            pyautogui.typewrite(['up'])
                            current_part_no9 = copy_clipboard()
                            if 'LABOR' not in current_part_no9:
                                go_up = False
                                break
                            up_count = up_count + 1
                        for down in range(up_count+1):
                            pyautogui.typewrite(['down'])
                        pyautogui.click(922,247,clicks=up_count,interval=0.025)

            labor_no = str('LABOR')

            down_loop = True

            if 'LABOR' in part_string_list[string_list_len - 1]:
                print('Labor at bottom')
                labor_lines = False
            if 'LABOR' in part_string_list[0]:
                print('Labor at top')
                labor_lines = True
                
            if labor_lines:
                go_up()
                pyautogui.doubleClick(130,243)
                print('\nLooking for LABOR lines.')

                for i in range(0,string_list_len):
                    current_part = part_string_list[i]
                    if "LABOR" in current_part:
                        total_lines = i
                        for down in range(total_lines):
                            pyautogui.typewrite(['down'])
                        time.sleep(0.5)
                        pyautogui.click(919,291)
                        break
                
                pyautogui.doubleClick(130,243)
                
                print('First LABOR line is: ' + part_string_list[total_lines])
                first_labor = part_string_list[total_lines]

                go_up()
                pyautogui.doubleClick(130,243)
                print('Lines before Labor line: ' + str(total_lines))
                labor_copy = False
                total_down = True
                while labor_copy is False:
                    if total_lines != 0:
                        if total_down:
                            print('\nGoing down.')
                            for i in range(0,total_lines):
                                pyautogui.PAUSE = 0.03
                                pyautogui.typewrite(['down'])
                                time.sleep(0.03)
                            time.sleep(1)
                            print('Down complete.')

                    pyautogui.PAUSE = 0.1
                    current_part_no = copy_clipboard()
                    time.sleep(0.1)
                    current = current_part_no.split('-')
                    time.sleep(0.1)
                    if current_part_no == first_labor:
                        labor_copy = True
                        go_up()
                        print('\nFound labor copy.')
                        break
                    if labor_no in current:
                        pyautogui.click(919,291)
                        time.sleep(.1)
                        go_up()
                        pyautogui.doubleClick(130,243)
                        print('Moved ' + str(current_part_no))
                        total_down = True
                        time.sleep(0.5)
                        continue
                    pyautogui.PAUSE = 0.05
                    print('Current part: ' +current_part_no)
                    pyautogui.typewrite(['down'])
                    total_down = False

            
            if pwht:
                if 'PWHT' in part_string_list[0]:
                    go_up()
                    pyautogui.doubleClick(130,243)
                    current_part_no = copy_clipboard()
                    pyautogui.PAUSE = 0.1
                    if 'PWHT' in current_part_no:
                        pyautogui.click(921,292)
                        time.sleep(0.25)
                if 'PWHT' in part_string_list[string_list_len - 1]:
                    go_up()
                    while down_loop:
                        pyautogui.PAUSE = 0.03
                        im1 = pyautogui.screenshot()
                        for i in range(10):
                            pyautogui.typewrite(['down'])
                        im2 = pyautogui.screenshot()
                        if im1 == im2:
                            down_loop = False
                            break
                    current_part_no = copy_clipboard()
                    pyautogui.PAUSE = 0.1
                    if 'PWHT' in current_part_no:
                        pyautogui.typewrite(['tab'])
                        pyautogui.hotkey('shift','tab')
                        time.sleep(0.25)
                        current_part_no9 = copy_clipboard()
                        up_count = 0
                        go_up_loop = True
                        while go_up_loop:
                            pyautogui.typewrite(['up'])
                            current_part_no9 = copy_clipboard()
                            if 'LABOR' not in current_part_no9:
                                go_up_loop = False
                                break
                            up_count = up_count + 1
                        for down in range(up_count+1):
                            pyautogui.typewrite(['down'])
                        pyautogui.click(922,247,clicks=up_count,interval=0.025)

            go_up()

            pyautogui.click(1423,15)

            add_part = input('\nAre there part numbers to add?: ')
            
            if '1' in str(add_part):
                pyautogui.PAUSE = 0.1
                add_part = True
                go_up_count = True
                up_count = 0
                while add_part:
                    new_part = input('What is the new part number?: ')
                    pyautogui.click(920,312) # new line
                    time.sleep(0.25)
                    pyautogui.typewrite(str(new_part))
                    time.sleep(0.5)
                    pyautogui.typewrite(['tab'])
                    time.sleep(0.25)
                    pyautogui.hotkey('shift','tab')
                    time.sleep(0.25)
                    current_part_no9 = copy_clipboard()
                    up_count = 0
                    if go_up_count:
                        while go_up_count:
                            pyautogui.typewrite(['up'])
                            current_part_no9 = copy_clipboard()
                            if 'LABOR' not in current_part_no9 and 'PWHT' not in current_part_no9:
                                go_up_count = False
                                break
                            up_count = up_count + 1
                        for down in range(up_count+1):
                            pyautogui.typewrite(['down'])
                        pyautogui.click(922,247,clicks=up_count,interval=0.025)
                    else:
                        pyautogui.click(920,247)
                        pyautogui.click(920,271)
                    pyautogui.click(1423,15)
                    add_more = input('Are there more parts to add?: ')
                    if '0' in str(add_more):
                        add_part = False
                        break
                    else:
                        go_up_count = False
            
            change_part = input('Are there part numbers that need to change?: ')
            
            if '1' in str(change_part):
                pyautogui.PAUSE = 0.05
                change_part = True
                while change_part:
                    go_up()
                    pyautogui.click(1423,15)
                    change_line = input('Which line needs to change?: ').zfill(2)
                    new_part = input('What is the new part number?: ')
                    go_up()
                    if '01' in change_line:
                        pyautogui.doubleClick(130,243)
                        pyautogui.click(919,289) # bottom
                        time.sleep(0.5)
                        pyautogui.typewrite(['tab'])
                        time.sleep(0.25)
                        current_part_no = copy_clipboard()
                        time.sleep(0.25)
                        pyautogui.typewrite('0')
                        time.sleep(0.5)
                        pyautogui.hotkey('shift','tab')
                        time.sleep(0.5)
                        current_part_no9 = copy_clipboard()
                        go_up_switch = True
                        up_count = 0
                        while go_up_switch:
                            pyautogui.typewrite(['up'])
                            current_part_no9 = copy_clipboard()
                            time.sleep(0.25)
                            if 'LABOR' not in current_part_no9 and 'PWHT' not in current_part_no9:
                                go_up_switch = False
                                break
                            up_count = up_count + 1
                        for down in range(up_count+1):
                            pyautogui.typewrite(['down'])
                        pyautogui.click(922,247,clicks=up_count,interval=0.025)
                        pyautogui.click(920,312) # new line
                        pyautogui.typewrite(str(new_part))
                        pyautogui.typewrite(['tab'])
                        pyautogui.typewrite(current_part_no)
                        time.sleep(0.5)
                        pyautogui.click(920,227) # top
                        go_up()
                    else:
                        go_up()
                        pyautogui.doubleClick(130,243)
                        change_line_range = int(change_line) - 1
                        for number_of_down in range(change_line_range):
                            pyautogui.PAUSE = 0.05
                            pyautogui.typewrite(['down'])
                        pyautogui.click(919,289) # bottom
                        time.sleep(0.5)
                        pyautogui.typewrite(['tab'])
                        time.sleep(0.1)
                        current_part_no = copy_clipboard()
                        time.sleep(0.1)
                        pyautogui.typewrite('0')
                        time.sleep(0.5)
                        pyautogui.hotkey('shift','tab')
                        time.sleep(0.5)
                        current_part_no9 = copy_clipboard()
                        go_up_switch = True
                        up_count = 0
                        while go_up_switch:
                            pyautogui.typewrite(['up'])
                            current_part_no9 = copy_clipboard()
                            time.sleep(0.25)
                            if 'LABOR' not in current_part_no9 and 'PWHT' not in current_part_no9:
                                go_up_switch = False
                                break
                            up_count = up_count + 1
                        for down in range(up_count+1):
                            pyautogui.typewrite(['down'])
                        pyautogui.click(922,247,clicks=up_count,interval=0.025)
                        pyautogui.click(920,312) # new line
                        pyautogui.typewrite(str(new_part))
                        pyautogui.typewrite(['tab'])
                        pyautogui.typewrite(current_part_no)
                        time.sleep(0.5)
                        pyautogui.click(920,227) # top
                        time.sleep(0.25)
                        for number_of_down in range(change_line_range):
                            pyautogui.click(920,269)
                    pyautogui.click(1423,15)
                    change_more = input('Are there more lines to change?: ')
                    if '0' in change_more:
                        change_part = False
                        break
            
            go_up()
                
            pyautogui.click(1423,15)
            pyautogui.PAUSE = 0.05
            #line_skip = int(input('How many lines need to be skipped?: '))
            line_skip = 0
            print('\nPress Ctrl-C to quit.')
            #pyautogui.click(1423,15)
            # start loop to get the line and amount information
            # type stop to move on
            item_list = []
            amount_list = []
            heat_list = []
            last_heat_number = ''
            stop_loop = False
            while not stop_loop:
                try:
                    item_input = input("Please enter the item number (enter + to quit): ")
                    try:
                        if '-' in item_input:
                            length = len(item_list) - 1
                            item_list.pop(length)
                            amount_list.pop(length)
                            heat_list.pop(length)
                            continue
                        if '+' in str(item_input) and len(item_list) == 0:
                            raise EmptyInput
                        if '+' in str(item_input):
                            stop_loop = True
                            break
                        if len(item_input) == 0:
                            raise EmptyInput
                    except EmptyInput:
                        print("Input is empty.\n")
                        continue
                    except ValueError:
                        continue
                    amount_input = input("Please enter the amount: ")
                    try:
                        if len(amount_input) == 0:
                            raise EmptyInput
                    except EmptyInput:
                        print("Input is empty.\n")
                        continue

                    heat_number = input("Please enter the heat number: ")

                    if '+' in heat_number:
                        heat_number = last_heat_number

                    item_list.append(int(item_input))
                    amount_list.append(float(amount_input))
                    heat_list.append(str(heat_number.upper()))
                    last_heat_number = heat_number
                except ValueError:
                    len_pop = len(item_list)
                    item_list.pop(len_pop-1)
                    print("\nValue error.")
                    continue
            
            go_up()
            pyautogui.PAUSE = 0.1
            max_range = len(item_list)
            pyautogui.click(1423,15)
            skip_me = '1'
            pyautogui.doubleClick(219,243)
            if '1' in str(skip_me):
                start_change = time.time()
                if max_range ==1:
                    item_change = item_list[0]
                    amount_change = amount_list[0]
                    heat_change = heat_list[0]
                    line = int(line_skip) + int(item_change) -1
                    for number_of_down in range(line):
                        pyautogui.typewrite(['down'])
                        time.sleep(0.05)
                    pyautogui.typewrite(str(amount_change))
                    pyautogui.typewrite(['tab'])
                    pyautogui.typewrite(str(heat_change))
                    #pyautogui.typewrite(['tab'])
                    #pyautogui.typewrite(str(heat_change))
                    #pyautogui.hotkey('shift','tab')
                    pyautogui.hotkey('shift','tab')
                    
                else:
                    item_change_1 = item_list[0]
                    amount_change_1 = amount_list[0]
                    heat_change_1 = heat_list[0]
                    line_1 = int(line_skip) + int(item_change_1) -1
                    for number_of_down in range(line_1):
                        pyautogui.typewrite(['down'])
                    pyautogui.typewrite(str(amount_change_1))
                    pyautogui.typewrite(['tab'])
                    pyautogui.typewrite(str(heat_change_1))
                    #pyautogui.typewrite(['tab'])
                    #pyautogui.typewrite(str(heat_change_1))
                    #pyautogui.hotkey('shift','tab')
                    pyautogui.hotkey('shift','tab')

                    last_line = item_change_1
                    
                    for change in range(1,max_range):
                        item_change = item_list[change]
                        amount_change = amount_list[change]
                        heat_change = heat_list[change]
                        line = int(item_change) - last_line
                        if line > 0:
                            for number_of_down in range(line):
                                pyautogui.typewrite(['down'])
                                time.sleep(0.05)
                        elif line < 0:
                            for number_of_up in range(abs(line)):
                                pyautogui.typewrite(['up'])
                                time.sleep(0.05)
                        pyautogui.typewrite(str(amount_change))
                        pyautogui.typewrite(['tab'])
                        pyautogui.typewrite(str(heat_change))
                        #pyautogui.typewrite(['tab'])
                        #pyautogui.typewrite(str(heat_change))
                        #pyautogui.hotkey('shift','tab')
                        pyautogui.hotkey('shift','tab')
                        last_line = item_change
            
            time.sleep(0.25)
            go_up()
            part_list = []
            stock_text_list = []
            wip_text_list = []
            released_text_list = []
            on_order_text_list = []
            stocking_text_list = []
            sleep = 0.1
            pyautogui.doubleClick(130,243)
            pyautogui.PAUSE = 0.1
            custom_oem_psm_config = r'--oem 3 --psm 6'
            if max_range ==1:
                item_change = item_list[0]
                line = int(line_skip) + int(item_change) -1
                for number_of_down in range(line):
                    pyautogui.typewrite(['down'])
                    time.sleep(0.05)
                part_no = copy_clipboard()
                arrow_location = pyautogui.locateOnScreen('arrow.png', region=(31,232,21,739))
                if arrow_location is None:
                    arrow_location = pyautogui.locateOnScreen('yellow_arrow.png', region=(31,232,21,739))
                if arrow_location is None:
                    arrow_location = pyautogui.locateOnScreen('pencil_arrow.png', region=(31,232,21,739))
                arrow_top = arrow_location.top-2
                arrow_height = arrow_location.height+4
                arrow_bottom = arrow_top + arrow_height
                # stock
                stock_text_im = ImageGrab.grab(bbox =(538,arrow_top,599,arrow_bottom))
                stock_size = tuple(4*x for x in stock_text_im.size)
                stock_text_im = stock_text_im.resize(stock_size, Image.ANTIALIAS)
                stock_text_im_gs = stock_text_im.convert('LA')
                stock_text_string = pytesseract.image_to_string(stock_text_im_gs, config=custom_oem_psm_config)
                stock_text_list.append(stock_text_string)
                #wip
                wip_text_im = ImageGrab.grab(bbox =(251,arrow_top,311,arrow_bottom))
                wip_size = tuple(4*x for x in wip_text_im.size)
                wip_text_im = wip_text_im.resize(wip_size, Image.ANTIALIAS)
                wip_text_im_gs = wip_text_im.convert('LA')
                wip_text_string = pytesseract.image_to_string(wip_text_im_gs, config=custom_oem_psm_config)
                wip_text_list.append(wip_text_string)
                #released
                released_text_im = ImageGrab.grab(bbox =(599,arrow_top,657,arrow_bottom))
                released_size = tuple(4*x for x in released_text_im.size)
                released_text_im = released_text_im.resize(released_size, Image.ANTIALIAS)
                released_text_im_gs = released_text_im.convert('LA')
                released_text_string = pytesseract.image_to_string(released_text_im_gs, config=custom_oem_psm_config)
                released_text_list.append(released_text_string)
                #on_order
                on_order_text_im = ImageGrab.grab(bbox =(657,arrow_top,723,arrow_bottom))
                on_order_size = tuple(4*x for x in on_order_text_im.size)
                on_order_text_im = on_order_text_im.resize(on_order_size, Image.ANTIALIAS)
                on_order_text_im_gs = on_order_text_im.convert('LA')
                on_order_text_string = pytesseract.image_to_string(on_order_text_im_gs, config=custom_oem_psm_config)
                on_order_text_list.append(on_order_text_string)
                #stocking
                stocking_text_im = ImageGrab.grab(bbox =(816,arrow_top,886,arrow_bottom))
                stocking_size = tuple(4*x for x in stocking_text_im.size)
                stocking_text_im = stocking_text_im.resize(stocking_size, Image.ANTIALIAS)
                stocking_text_im_gs = stocking_text_im.convert('LA')
                stocking_text_string = pytesseract.image_to_string(stocking_text_im_gs, config=custom_oem_psm_config)
                stocking_text_list.append(stocking_text_string)
                
                #time.sleep(0.25)
                print("\nPart: {}".format(part_no))
                print("WIP: {}".format(wip_text_string))
                print("Item Stock: {}".format(stock_text_string))
                print("Released: {}".format(released_text_string))
                print("On Order: {}".format(on_order_text_string))
                print("Stocking: {}".format(stocking_text_string))
                part_list.append(str(part_no))
                
            else:
                item_change_1 = item_list[0]
                line_1 = int(line_skip) + int(item_change_1) -1
                for number_of_down in range(line_1):
                    pyautogui.typewrite(['down'])
                    time.sleep(0.1)
                time.sleep(0.1)
                part_no = copy_clipboard()
                arrow_location = pyautogui.locateOnScreen('arrow.png', region=(31,232,21,739))
                if arrow_location is None:
                    arrow_location = pyautogui.locateOnScreen('yellow_arrow.png', region=(31,232,21,739))
                if arrow_location is None:
                    arrow_location = pyautogui.locateOnScreen('pencil_arrow.png', region=(31,232,21,739))
                arrow_top = arrow_location.top-2
                arrow_height = arrow_location.height+4
                arrow_bottom = arrow_top + arrow_height
                # stock
                stock_text_im = ImageGrab.grab(bbox =(538,arrow_top,599,arrow_bottom))
                stock_size = tuple(4*x for x in stock_text_im.size)
                stock_text_im = stock_text_im.resize(stock_size, Image.ANTIALIAS)
                stock_text_im_gs = stock_text_im.convert('LA')
                stock_text_string = pytesseract.image_to_string(stock_text_im_gs, config=custom_oem_psm_config)
                stock_text_list.append(stock_text_string)
                #wip
                wip_text_im = ImageGrab.grab(bbox =(251,arrow_top,311,arrow_bottom))
                wip_size = tuple(4*x for x in wip_text_im.size)
                wip_text_im = wip_text_im.resize(wip_size, Image.ANTIALIAS)
                wip_text_im_gs = wip_text_im.convert('LA')
                wip_text_string = pytesseract.image_to_string(wip_text_im_gs, config=custom_oem_psm_config)
                wip_text_list.append(wip_text_string)
                #released
                released_text_im = ImageGrab.grab(bbox =(599,arrow_top,657,arrow_bottom))
                released_size = tuple(4*x for x in released_text_im.size)
                released_text_im = released_text_im.resize(released_size, Image.ANTIALIAS)
                released_text_im_gs = released_text_im.convert('LA')
                released_text_string = pytesseract.image_to_string(released_text_im_gs, config=custom_oem_psm_config)
                released_text_list.append(released_text_string)
                #on_order
                on_order_text_im = ImageGrab.grab(bbox =(657,arrow_top,723,arrow_bottom))
                on_order_size = tuple(4*x for x in on_order_text_im.size)
                on_order_text_im = on_order_text_im.resize(on_order_size, Image.ANTIALIAS)
                on_order_text_im_gs = on_order_text_im.convert('LA')
                on_order_text_string = pytesseract.image_to_string(on_order_text_im_gs, config=custom_oem_psm_config)
                on_order_text_list.append(on_order_text_string)
                #stocking
                stocking_text_im = ImageGrab.grab(bbox =(816,arrow_top,886,arrow_bottom))
                stocking_size = tuple(4*x for x in stocking_text_im.size)
                stocking_text_im = stocking_text_im.resize(stocking_size, Image.ANTIALIAS)
                stocking_text_im_gs = stocking_text_im.convert('LA')
                stocking_text_string = pytesseract.image_to_string(stocking_text_im_gs, config=custom_oem_psm_config)
                stocking_text_list.append(stocking_text_string)
                
                #time.sleep(0.25)
                print("\nPart: {}".format(part_no))
                print("WIP: {}".format(wip_text_string))
                print("Item Stock: {}".format(stock_text_string))
                print("Released: {}".format(released_text_string))
                print("On Order: {}".format(on_order_text_string))
                print("Stocking: {}".format(stocking_text_string))
                part_list.append(str(part_no))
                last_line = item_change_1
                
                for change in range(1,max_range):
                    item_change = item_list[change]
                    amount_change = amount_list[change]
                    line = int(item_change) - last_line
                    if line > 0:
                        for number_of_down in range(line):
                            pyautogui.typewrite(['down'])
                            time.sleep(0.1)
                    elif line < 0:
                        for number_of_up in range(abs(line)):
                            pyautogui.typewrite(['up'])
                            time.sleep(0.1)
                    part_no = copy_clipboard()
                    arrow_location = pyautogui.locateOnScreen('arrow.png', region=(31,232,21,739))
                    if arrow_location is None:
                        arrow_location = pyautogui.locateOnScreen('yellow_arrow.png', region=(31,232,21,739))
                    if arrow_location is None:
                        arrow_location = pyautogui.locateOnScreen('pencil_arrow.png', region=(31,232,21,739))
                    arrow_top = arrow_location.top-2
                    arrow_height = arrow_location.height+4
                    arrow_bottom = arrow_top + arrow_height
                    # stock
                    stock_text_im = ImageGrab.grab(bbox =(538,arrow_top,599,arrow_bottom))
                    stock_size = tuple(4*x for x in stock_text_im.size)
                    stock_text_im = stock_text_im.resize(stock_size, Image.ANTIALIAS)
                    stock_text_im_gs = stock_text_im.convert('LA')
                    stock_text_string = pytesseract.image_to_string(stock_text_im_gs, config=custom_oem_psm_config)
                    stock_text_list.append(stock_text_string)
                    #wip
                    wip_text_im = ImageGrab.grab(bbox =(251,arrow_top,311,arrow_bottom))
                    wip_size = tuple(4*x for x in wip_text_im.size)
                    wip_text_im = wip_text_im.resize(wip_size, Image.ANTIALIAS)
                    wip_text_im_gs = wip_text_im.convert('LA')
                    wip_text_string = pytesseract.image_to_string(wip_text_im_gs, config=custom_oem_psm_config)
                    wip_text_list.append(wip_text_string)
                    #released
                    released_text_im = ImageGrab.grab(bbox =(599,arrow_top,657,arrow_bottom))
                    released_size = tuple(4*x for x in released_text_im.size)
                    released_text_im = released_text_im.resize(released_size, Image.ANTIALIAS)
                    released_text_im_gs = released_text_im.convert('LA')
                    released_text_string = pytesseract.image_to_string(released_text_im_gs, config=custom_oem_psm_config)
                    released_text_list.append(released_text_string)
                    #on_order
                    on_order_text_im = ImageGrab.grab(bbox =(657,arrow_top,723,arrow_bottom))
                    on_order_size = tuple(4*x for x in on_order_text_im.size)
                    on_order_im = on_order_text_im.resize(on_order_size, Image.ANTIALIAS)
                    on_order_text_im_gs = on_order_text_im.convert('LA')
                    on_order_text_string = pytesseract.image_to_string(on_order_text_im_gs, config=custom_oem_psm_config)
                    on_order_text_list.append(on_order_text_string)
                    #stocking
                    stocking_text_im = ImageGrab.grab(bbox =(816,arrow_top,886,arrow_bottom))
                    stocking_size = tuple(4*x for x in stocking_text_im.size)
                    stocking_text_im = stocking_text_im.resize(stocking_size, Image.ANTIALIAS)
                    stocking_text_im_gs = stocking_text_im.convert('LA')
                    stocking_text_string = pytesseract.image_to_string(stocking_text_im_gs, config=custom_oem_psm_config)
                    stocking_text_list.append(stocking_text_string)
                    
                    #time.sleep(0.25)
                    print("\nPart: {}".format(part_no))
                    print("WIP: {}".format(wip_text_string))
                    print("Item Stock: {}".format(stock_text_string))
                    print("Released: {}".format(released_text_string))
                    print("On Order: {}".format(on_order_text_string))
                    print("Stocking: {}".format(stocking_text_string))
                    part_list.append(str(part_no))
                    last_line = item_change
                            
            # if amount to skip is >0 then go down past labor
            labor_range = item_list[0]
            labor_range = int(labor_range)
            if len(item_list) == 0:
                for number_of_down_labor in range(line_skip):
                    pyautogui.typewrite(['down'])
            
            pyautogui.click(1423,15)
            if '1' in str(skip_me):
                end_change = time.time()-start_change
                print('\nChange value time: ' + str(round(end_change, 3)) + ' Seconds')
                
            # click save
            pyautogui.click(75,65)
            time.sleep(0.5)
            # open transfer window
            pyautogui.click(550,60)
            time.sleep(1)
            okWindow = gw.getWindowsWithTitle('Start Manufacturing Order Detail')
            while len(gw.getWindowsWithTitle('Start Manufacturing Order Detail')) == 0:
                time.sleep(1)
            time.sleep(0.5)
            time.sleep(0.1)
            pyautogui.PAUSE = 0.1
            
            start_transfer = time.time()
        # material transfer loop
            red_x_list = []
            transfer_time_list = []
            wip_after_list = []
            transfers_list = []
            stock_list = []
            amount_transfer_list = []
            wip_list = []
            pyautogui.PAUSE = 0.1
            for transfer in range(0,max_range):
                pyautogui.click(645,684)
                transfer_switch = False
                item_transfer = int(item_list[transfer]) + line_skip
                amount_transfer = float(amount_list[transfer])
                stock = str(stock_text_list[transfer])
                part = str(part_list[transfer])
                start_wip = float(wip_text_list[transfer])
                released = float(released_text_list[transfer])
                stocking = str(stocking_text_list[transfer])

                wip = '+'

                if len(stock) == 0:
                    stock = '0.000'

                if ',' in str(stock) and len(stock) >= 9:
                    stock = stock.replace(",","",1)
                
                if ',' in str(stock):
                    stock = stock.replace(",",".")

                if ' ' in str(stock):
                    stock = stock.replace(" ",".")

                stock = float(stock)

                if part in transfers_list:
                    for search in range(0,len(transfers_list)):
                        item = transfers_list[search]
                        required = amount_transfer_list[search]
                        stock = float(stock_list[search])
                        if item == part and stock > 0:
                            if '+' in wip:
                                stock = float(stock) - float(required)
                            else:
                                stock = float(stock) + float(required)

                if float(stock) > 0:
                    red_x_look = True
                    transfer_switch = True
                    if start_wip == amount_transfer:
                        wip = '0'
                        print('\nItem {} already in WIP'.format(item_transfer))
                        red_x_list.append('Already in WIP')
                        transfer_switch = False
                    elif amount_transfer > stock:
                        wip = '0'
                        red_x_look = False
                        amount_transfer = stock
                        print('\nNot enough of item {} in stock'.format(item_transfer))
                        red_x_list.append('Not enough in stock')

                    if amount_transfer > start_wip:
                        amount_transfer = amount_transfer - start_wip

                    if start_wip > float(amount_list[transfer]):
                        wip = '-'
                        pyautogui.click(831,682)
                        amount_transfer = start_wip - float(amount_list[transfer])
                else:
                    if amount_transfer == 0:
                        wip = '0'
                        print('\nTransfer amount is 0')
                        red_x_list.append('Transfer amount is 0')
                        transfer_switch = False
                    elif start_wip == amount_transfer:
                        wip = '0'
                        print('\nItem {} already in WIP'.format(item_transfer))
                        red_x_list.append('Already in WIP')
                        transfer_switch = False
                    else:
                        print("None in stock")
                        wip = '0'
                        red_x_list.append('None in stock')
                    
                transfer_time = time.strftime('%Y-%m-%d, %H:%M:%S', time.localtime())
                transfer_time_list.append(transfer_time)

                if transfer_switch:
                    # activate window
                    pyautogui.click(827,252)
                    # double click search bar
                    pyautogui.doubleClick(590,453)
                    # input that line into the search bar
                    pyautogui.typewrite(str(item_transfer))
                    # double click the input bar
                    pyautogui.doubleClick(615,611)
                    # input desired amount into the transfer bar
                    pyautogui.typewrite(str(amount_transfer))
                    # click transfer
                    pyautogui.click(621,779)
                    
                    time.sleep(1)
                    # ok window
                    okWindow = gw.getWindowsWithTitle('Start Manufacturing Order Detail')
                    while len(gw.getWindowsWithTitle('Start Manufacturing Order Detail')) == 1:
                        time.sleep(0.25)
                        
                    time.sleep(0.25)
                    
                    red_x = pyautogui.locateOnScreen('red_x.png', region=(650,450,100,100))

                    time.sleep(0.25)

                    pyautogui.click(1034,597)
                    pyautogui.click(868,568)
                    print('\nTransfering {} {} for item {}'.format(round(amount_transfer,3), stocking, item_transfer))
                    
                    if red_x_look:
                        if red_x is not None:
                            print('\nNot enough of item {} in stock'.format(item_transfer))
                            red_x_list.append('Not enough in stock')
                            wip_after = float(wip_text_list[transfer])
                        else:
                            print('\nEnough of item {} in stock'.format(item_transfer))
                            red_x_list.append('Enough in stock')
                            wip_after = start_wip + amount_transfer
                    
                    transfers_list.append(part)
                
                else:
                    wip_after = float(wip_text_list[transfer])
                
                wip_after_list.append(float(wip_after))
                stock_list.append(float(stock))
                amount_transfer_list.append(float(amount_transfer))
                wip_list.append(wip)
            # close transfer window
            pyautogui.click(1004, 776)

            while len(gw.getWindowsWithTitle('Start Manufacturing Order Detail')) != 0:
                time.sleep(0.5)

            pyautogui.doubleClick(210,101)
            job_no = copy_clipboard()
            
            """
            file_routing()

            current_date = time.strftime('%Y-%m-%d, %Hh %Mm %Ss', time.localtime())

            with open('{}.csv'.format(current_date), 'w', newline='') as csv_file:
                fieldnames = ['Time','Job number','+/-','Item','Part Number','Heat number','Amount','WIP Before','WIP After','Stock','Stock Amount','Released','On Order']
                csv_writer = csv.DictWriter(csv_file, fieldnames=fieldnames)

                csv_writer.writeheader()

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
            """

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

            end_transfer = time.time() - start_transfer
            print('\nTransfer time: ' + str(round(end_transfer, 2)))
            if '1' in skip_me:
                elapsed_time = round(end_transfer + end_change, 2)
                print('\nElapsed automation time: ' + str(elapsed_time) + ' Seconds')
            total_time = time.time() - total_start
            minutes = 0
            while total_time >= 60:
                total_time = total_time - 60
                minutes = minutes + 1
            
            print('\nTotal Time: ' + str(minutes) + ' Minutes ' + str(round(total_time, 2)) + ' Seconds')
            print('\nCompleted Loop.')
            print('\n----------------------------------------')

        else:
            if released_button is None:
                print('\nThe job is closed.')
            if on_hold_button is not None:
                print('\nThe job is on hold.')
            print('\n----------------------------------------')
        ask_for_next_job = True

except KeyboardInterrupt:
    pass

combine_workbook()

os.chdir(r"C:\Users\jlee.NTPV\Documents\GitHub\material_transfer")

print('\nDone')
