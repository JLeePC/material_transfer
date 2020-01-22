#! python3

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
from copy_clipboard import copy_clipboard
from go_up import go_up
from combine_workbook import combine_workbook

try:
    os.chdir(r"\\NTPV-SERVER2008\ntpv data\Justyn's MISys")
except:
    # login to be able to copy the backup of the xlsx to the Z: server
    username = "NTPV\jlee"
    password = getpass.getpass("Enter Password: ")
    path = r'\\NTPV-SERVER2008\ntpv data'

    mount_command = "net use {} /user:{} {} /persistent:no".format(path,username,password)
    os.system(mount_command)

pytesseract.pytesseract.tesseract_cmd = r'C:\Users\jlee.NTPV\AppData\Local\Tesseract-OCR\tesseract.exe'

os.chdir(r"C:\Users\jlee.NTPV\Documents\GitHub\material_transfer")

# Exceptions for data input
class EmptyInput(Exception):
    pass

# def for removing values from item list reading
def remove_values_from_list(the_list, val):
   return [value for value in the_list if value != val]

ask_for_next_job = False
pyautogui.PAUSE = 0.1
print('\nCtrl_c to quit.')

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
            # wait for released_by.png to show up to be sure that the page loaded
            waiting = True
            while waiting:
                released_by = pyautogui.locateOnScreen('released_by.png', region=(370,250,130,50))
                if released_by is not None:
                    waiting = False
                    break
            """
            pyautogui.doubleClick(214,102)
            next_job = copy_clipboard()
            time.sleep(0.1)
            pyautogui.doubleClick(228,313) # job no.

            # double check that the correct job is loaded
            job_no = copy_clipboard()
            
            job = False

            while not job:
                pyautogui.doubleClick(228,313) # job no.
                job_no = copy_clipboard()
                if job_no ==  next_job.upper():
                    job = True
                    break
                time.sleep(0.25)
            """
                
        total_start = time.time()
        
        # check to see if the job is closed or on hold
        # if either are true prohibit changes to the mo
        time.sleep(1)
        released_button = pyautogui.locateOnScreen('released.png', region=(614,86,170,29))
        on_hold_button = pyautogui.locateOnScreen('on_hold.png', region=(712,115,91,24))
        time.sleep(0.25)

        if released_button is not None and on_hold_button is None:
            
            # click in a box and go up twice to make sure the lines are loaded properly
            # if you are scrolled down in a job and go to a job without a scroll bar it wont be at the top like it should
            pyautogui.click(150,150)
            time.sleep(0.5)
            pyautogui.click(351,243)
            pyautogui.typewrite(['up'])
            pyautogui.typewrite(['up'])
            # check to see if there is a scroll bar
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
            
            # grab all the item no. to know where the LABOR lines and PWHT lines are
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
                # scroll down 35 lines to the next page and repeat till at the bottom
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
                    
                    if len(line_number_string_list2) <= 35:
                        print("At Bottom")
                        labor_loop = False
                    if not line_number_string_list2[0] in line_number_string_list[len(line_number_string_list)-1]:
                        part_string_list = part_string_list + part_string_list2
                        line_number_string_list = line_number_string_list + line_number_string_list2

            # else for if there is only one page
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

            # go through the list of parts and delete any that are empty
            count = 0
            for i in range(0,len(part_string_list)-1):
                current_part = part_string_list[i-count]
                if len(current_part) == 0:
                    part_string_list.pop(i-count)
                    count += 1

            string_list_len = len(part_string_list)

            go_up()

            
            down_loop = True
            move_labor_lines = False
            if labor_lines:
                move_labor_lines = True
                # check if LABOR is at the bottom
                if 'LABOR' in part_string_list[string_list_len - 1]:
                    print('Labor at bottom')
                    move_labor_lines = False
                # check if LABOR is at the top
                if 'LABOR' in part_string_list[0]:
                    print('Labor at top')
                    move_labor_lines = True

                # check if there are any part numbers inside the LABOR lines
                first_labor = 0
                for i in range(0,len(part_string_list)):
                    current_part = part_string_list[i]
                    if "LABOR" in current_part and first_labor == 0:
                        first_labor = i
                    if "LABOR" in current_part:
                        last_labor = i
                for i in range(first_labor, last_labor):
                    current_part = part_string_list[i]
                    if not "LABOR" in current_part:
                        move_labor_lines = True

            if pwht and move_labor_lines:
                pwht_lines = True
                # if PWHT lines are at the bottom dont do anything
                if 'PWHT' in part_string_list[string_list_len - 1]:
                    print('PWHT at bottom')
                    pwht_lines = False
                # move PWHT lines to the bottom
                if pwht_lines:
                    go_up()
                    pyautogui.doubleClick(130,243)
                    print('\nLooking for PWHT lines.')

                    # look for the line that PWHT is at
                    for i in range(0,string_list_len):
                        current_part = part_string_list[i]
                        if "PWHT" in current_part:
                            print("PWHT found")
                            go_up()
                            pyautogui.doubleClick(130,243)
                            total_lines = i
                            # go from top down to PWHT
                            for down in range(total_lines):
                                pyautogui.typewrite(['down'])
                            time.sleep(0.5)
                            # move PWHT to bottom
                            pyautogui.click(919,291)
                            # TODO if labor is at bottom automatically move the PWHT above it so it can skip moving all the labor below it
            else:
                pwht_lines = False
            
            # move LABOR to bottom
            if move_labor_lines:
                go_up()
                pyautogui.doubleClick(130,243)
                print('\nLooking for LABOR lines.')

                # look for first LABOR line
                for i in range(0,string_list_len):
                    current_part = part_string_list[i]
                    if "LABOR" in current_part:
                        total_lines = i
                        if pwht_lines:
                            total_lines = total_lines -1
                        for down in range(total_lines):
                            pyautogui.typewrite(['down'])
                        time.sleep(0.5)
                        pyautogui.click(919,291)
                        break
                
                pyautogui.doubleClick(130,243)
                if pwht_lines:
                    print('First LABOR line is: ' + part_string_list[total_lines+1])
                    first_labor = part_string_list[total_lines+1]
                else:
                    print('First LABOR line is: ' + part_string_list[total_lines])
                    first_labor = part_string_list[total_lines]

                # move the rest of the LABOR lines
                go_up()
                pyautogui.doubleClick(130,243)
                print('Lines before Labor line: ' + str(total_lines))
                labor_copy = False
                total_down = True
                while labor_copy is False:
                    # if LABOR isnt at the top
                    # also it saves the line count to get to the previous labor line so it can go fast back to there
                    if total_lines != 0:
                        if total_down:
                            # go down fast till it gets to the line that it knows has LABOR
                            print('\nGoing down.')
                            for i in range(0,total_lines):
                                pyautogui.PAUSE = 0.03
                                pyautogui.typewrite(['down'])
                                time.sleep(0.03)
                            time.sleep(0.5)
                            print('Down complete.')

                    pyautogui.PAUSE = 0.1
                    # check to see if LABOR is at the current line
                    current_part_no = copy_clipboard()
                    current = current_part_no.split('-')
                    # if the current part matches the first labor line it knows its done moving labor lines
                    if current_part_no == first_labor:
                        labor_copy = True
                        go_up()
                        print('\nFound labor copy.')
                        break
                    # if labor in current move it to the bottom
                    if "LABOR" in current:
                        pyautogui.click(919,291)
                        time.sleep(.1)
                        go_up()
                        pyautogui.doubleClick(130,243)
                        print('Moved ' + str(current_part_no))
                        total_down = True
                        time.sleep(0.5)
                        continue
                    # if labor is not in the current part it knows to go down and check each line till it finds more labor
                    else:
                        total_down = False
                    pyautogui.PAUSE = 0.05
                    print('Current part: ' +current_part_no)
                    pyautogui.typewrite(['down'])
                    total_lines += 1

            """
            # once i add the todo from the previous pwht i can delete this
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
            #---------------
            """

            # add part numbers to mo
            go_up()
            pyautogui.click(1423,15)
            # 1 means yes, 0 means no
            add_part = input('\nAre there part numbers to add?: ')
            if '1' in str(add_part):
                pyautogui.PAUSE = 0.1
                add_part = True
                go_up_count = True
                up_count = 0
                while add_part:
                    # ask for part no.
                    new_part = input('What is the new part number?: ')
                    pyautogui.click(920,312) # new line
                    time.sleep(0.25)
                    pyautogui.typewrite(str(new_part))
                    time.sleep(1.5)
                    pyautogui.typewrite(['tab'])
                    time.sleep(0.25)
                    pyautogui.hotkey('shift','tab')
                    time.sleep(0.25)
                    # move new part aove labor lines and owht lines
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
            
            # ask if there are any part no. to change
            change_part = input('Are there part numbers that need to change?: ')
            if '1' in str(change_part):
                pyautogui.PAUSE = 0.05
                change_part = True
                while change_part:
                    go_up()
                    pyautogui.click(1423,15)
                    # had to zfill so i can make sure not to only do the top line is 10, 11 etc. are input
                    change_line = input('Which line needs to change?: ').zfill(2)
                    new_part = input('What is the new part number?: ')
                    go_up()
                    # if its the first line
                    if '01' in change_line:
                        pyautogui.doubleClick(130,243)
                        pyautogui.click(919,289) # bottom
                        time.sleep(0.5)
                        pyautogui.typewrite(['tab'])
                        time.sleep(0.25)
                        current_part_no = copy_clipboard() # copy amount required
                        time.sleep(0.25)
                        pyautogui.typewrite('0') # zero it out
                        time.sleep(0.5)
                        pyautogui.hotkey('shift','tab')
                        time.sleep(0.5)
                        current_part_no9 = copy_clipboard()
                        go_up_switch = True
                        up_count = 0
                        # loop to put it above labor and pwht
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
                        pyautogui.click(922,247,clicks=up_count,interval=0.025) # click to move it above labor and pwht
                        pyautogui.click(920,312) # new line
                        pyautogui.typewrite(str(new_part))
                        pyautogui.typewrite(['tab'])
                        pyautogui.typewrite(current_part_no)
                        time.sleep(0.5)
                        pyautogui.click(920,227) # move new item to top
                        go_up()
                    # if its not the first line
                    else:
                        go_up()
                        pyautogui.doubleClick(130,243)
                        change_line_range = int(change_line) - 1
                        # go down to part
                        for number_of_down in range(change_line_range):
                            pyautogui.PAUSE = 0.05
                            pyautogui.typewrite(['down'])
                        pyautogui.click(919,289) # move part to bottom
                        time.sleep(0.5)
                        pyautogui.typewrite(['tab'])
                        time.sleep(0.1)
                        current_part_no = copy_clipboard() # copy amount
                        time.sleep(0.1)
                        pyautogui.typewrite('0') # zero out old part
                        time.sleep(0.5)
                        pyautogui.hotkey('shift','tab')
                        time.sleep(0.5)
                        current_part_no9 = copy_clipboard()
                        go_up_switch = True
                        up_count = 0
                        # count how many lines to get above labor and pwht
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
                        pyautogui.click(922,247,clicks=up_count,interval=0.025) # click to move part above labor and pwht
                        pyautogui.click(920,312) # create new line
                        pyautogui.typewrite(str(new_part)) # type new part
                        pyautogui.typewrite(['tab'])
                        pyautogui.typewrite(current_part_no) # type copied amount
                        time.sleep(0.5)
                        pyautogui.click(920,227) # move to top
                        time.sleep(0.25)
                        # move new part to the same line the old part was at
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
            print('\nPress Ctrl-C to quit.')
            item_list = []
            amount_list = []
            heat_list = []
            last_heat_number = ''
            stop_loop = False
            # loop to gather all data
            while not stop_loop:
                try:
                    item_input = input("Please enter the item number (enter + to quit): ") # line number
                    try:
                        # if '-' in part delete last input item
                        if '-' in item_input:
                            length = len(item_list) - 1
                            item_list.pop(length)
                            amount_list.pop(length)
                            heat_list.pop(length)
                            continue
                        # if '+' in part and no data was input use all data from last job no
                        if '+' in str(item_input) and len(item_list) == 0:
                            stop_loop = True
                            item_list = last_item_list
                            amount_list = last_amount_list
                            heat_list = last_heat_list
                            break
                        # if '+' in part data input is complete
                        if '+' in str(item_input) and len(item_list) != 0:
                            stop_loop = True
                            break
                        if len(item_input) == 0:
                            raise EmptyInput
                    except EmptyInput:
                        print("Input is empty.\n")
                        continue
                    except ValueError:
                        continue
                    # amount to transfer
                    amount_input = input("Please enter the amount: ")
                    try:
                        if len(amount_input) == 0:
                            raise EmptyInput
                        if '/' in amount_input:
                            amount_input = amount_input.replace("/","")
                            amount_input = round(float(amount_input)/12,3)
                    except EmptyInput:
                        print("Input is empty.\n")
                        continue
                    
                    # heat number for part
                    heat_number = input("Please enter the heat number: ")

                    if '+' in heat_number:
                        heat_number = last_heat_number

                    item_list.append(int(item_input))
                    amount_list.append(str(amount_input))
                    heat_list.append(str(heat_number.upper()))
                    last_heat_number = heat_number
                except ValueError:
                    len_pop = len(item_list)
                    item_list.pop(len_pop-1)
                    print("\nValue error.")
                    continue
            
            # save all data for if i want to reuse the same data on the next job
            last_item_list = item_list
            last_amount_list = amount_list
            last_heat_list = heat_list

            # this is all for gathering data of information i cant highlight the text of to copy to clipboard
            # plus its faster this way and surprisingly accurate
            # i feel ike i can compress this a lot but for now its long
            start_change = time.time()
            time.sleep(0.25)
            go_up()
            part_list = []
            stock_text_list = []
            wip_text_list = []
            released_text_list = []
            on_order_text_list = []
            stocking_text_list = []
            sleep = 0.1
            pyautogui.doubleClick(130,243) # click part no.
            pyautogui.PAUSE = 0.1
            custom_oem_psm_config = r'--oem 3 --psm 6' # not sure what this does but until i added it it wouldnt read properly
            max_range = len(item_list) # get length of list
            # if only 1 item in list
            if max_range ==1:
                item_change = item_list[0]
                amount_change = amount_list[0]
                heat_change = heat_list[0]
                line = int(item_change) -1
                # loop to go to the line
                for number_of_down in range(line):
                    pyautogui.typewrite(['down'])
                    time.sleep(0.05)
                part_no = copy_clipboard() # copy part number to add to list
                # search for all 3 different possible arrows to get the position of the line
                arrow_location = pyautogui.locateOnScreen('arrow.png', region=(31,232,21,739))
                if arrow_location is None:
                    arrow_location = pyautogui.locateOnScreen('yellow_arrow.png', region=(31,232,21,739))
                if arrow_location is None:
                    arrow_location = pyautogui.locateOnScreen('pencil_arrow.png', region=(31,232,21,739))
                arrow_top = arrow_location.top-2
                arrow_height = arrow_location.height+4
                arrow_bottom = arrow_top + arrow_height
                # all of there are the same just different x coordinates
                # stock
                stock_text_im = ImageGrab.grab(bbox =(539,arrow_top,599,arrow_bottom)) # grab im to read
                stock_size = tuple(4*x for x in stock_text_im.size) # get size of image
                stock_text_im = stock_text_im.resize(stock_size, Image.ANTIALIAS) # increase size of image
                stock_text_im_gs = stock_text_im.convert('LA') # convert image to black and white
                stock_text_string = pytesseract.image_to_string(stock_text_im_gs, config=custom_oem_psm_config) # read the image
                stock_text_list.append(stock_text_string) # add it to string
                #wip
                wip_text_im = ImageGrab.grab(bbox =(252,arrow_top,311,arrow_bottom))
                wip_size = tuple(4*x for x in wip_text_im.size)
                wip_text_im = wip_text_im.resize(wip_size, Image.ANTIALIAS)
                wip_text_im_gs = wip_text_im.convert('LA')
                wip_text_string = pytesseract.image_to_string(wip_text_im_gs, config=custom_oem_psm_config)
                wip_text_list.append(wip_text_string)
                #released
                released_text_im = ImageGrab.grab(bbox =(600,arrow_top,657,arrow_bottom))
                released_size = tuple(4*x for x in released_text_im.size)
                released_text_im = released_text_im.resize(released_size, Image.ANTIALIAS)
                released_text_im_gs = released_text_im.convert('LA')
                released_text_string = pytesseract.image_to_string(released_text_im_gs, config=custom_oem_psm_config)
                released_text_list.append(released_text_string)
                #on_order
                on_order_text_im = ImageGrab.grab(bbox =(658,arrow_top,723,arrow_bottom))
                on_order_size = tuple(4*x for x in on_order_text_im.size)
                on_order_text_im = on_order_text_im.resize(on_order_size, Image.ANTIALIAS)
                on_order_text_im_gs = on_order_text_im.convert('LA')
                on_order_text_string = pytesseract.image_to_string(on_order_text_im_gs, config=custom_oem_psm_config)
                on_order_text_list.append(on_order_text_string)
                #stocking
                stocking_text_im = ImageGrab.grab(bbox =(817,arrow_top,886,arrow_bottom))
                stocking_size = tuple(4*x for x in stocking_text_im.size)
                stocking_text_im = stocking_text_im.resize(stocking_size, Image.ANTIALIAS)
                stocking_text_im_gs = stocking_text_im.convert('LA')
                stocking_text_string = pytesseract.image_to_string(stocking_text_im_gs, config=custom_oem_psm_config)
                stocking_text_list.append(stocking_text_string)
                
                # pring all the informaion that i grabbed 
                print("\nPart: {}".format(part_no))
                print("WIP: {}".format(wip_text_string))
                print("Item Stock: {}".format(stock_text_string))
                print("Released: {}".format(released_text_string))
                print("On Order: {}".format(on_order_text_string))
                print("Stocking: {}".format(stocking_text_string))
                part_list.append(str(part_no)) # add part number to list

                # this is where i change the amount required and the heat number
                pyautogui.typewrite(['tab']) # tab over to required amount
                if '+' not in str(amount_change):
                    pyautogui.typewrite(str(amount_change)) # type amount
                # if heat number was input
                if len(heat_change) > 0:
                    pyautogui.typewrite(['tab']) # tab over to heat number
                    pyautogui.typewrite(str(heat_change)) # input heat number
                    pyautogui.hotkey('shift','tab') # shift tab back
                pyautogui.hotkey('shift','tab')
                time.sleep(0.5)
            
            # this is all the same as above for the fist item
            else:
                item_change_1 = item_list[0]
                amount_change_1 = amount_list[0]
                heat_change_1 = heat_list[0]
                line_1 = int(item_change_1) -1
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
                stock_text_im = ImageGrab.grab(bbox =(539,arrow_top,599,arrow_bottom))
                stock_size = tuple(4*x for x in stock_text_im.size)
                stock_text_im = stock_text_im.resize(stock_size, Image.ANTIALIAS)
                stock_text_im_gs = stock_text_im.convert('LA')
                stock_text_string = pytesseract.image_to_string(stock_text_im_gs, config=custom_oem_psm_config)
                stock_text_list.append(stock_text_string)
                #wip
                wip_text_im = ImageGrab.grab(bbox =(252,arrow_top,311,arrow_bottom))
                wip_size = tuple(4*x for x in wip_text_im.size)
                wip_text_im = wip_text_im.resize(wip_size, Image.ANTIALIAS)
                wip_text_im_gs = wip_text_im.convert('LA')
                wip_text_string = pytesseract.image_to_string(wip_text_im_gs, config=custom_oem_psm_config)
                wip_text_list.append(wip_text_string)
                #released
                released_text_im = ImageGrab.grab(bbox =(600,arrow_top,657,arrow_bottom))
                released_size = tuple(4*x for x in released_text_im.size)
                released_text_im = released_text_im.resize(released_size, Image.ANTIALIAS)
                released_text_im_gs = released_text_im.convert('LA')
                released_text_string = pytesseract.image_to_string(released_text_im_gs, config=custom_oem_psm_config)
                released_text_list.append(released_text_string)
                #on_order
                on_order_text_im = ImageGrab.grab(bbox =(658,arrow_top,723,arrow_bottom))
                on_order_size = tuple(4*x for x in on_order_text_im.size)
                on_order_text_im = on_order_text_im.resize(on_order_size, Image.ANTIALIAS)
                on_order_text_im_gs = on_order_text_im.convert('LA')
                on_order_text_string = pytesseract.image_to_string(on_order_text_im_gs, config=custom_oem_psm_config)
                on_order_text_list.append(on_order_text_string)
                #stocking
                stocking_text_im = ImageGrab.grab(bbox =(817,arrow_top,886,arrow_bottom))
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

                pyautogui.typewrite(['tab'])
                if '+' not in str(amount_change_1):
                    pyautogui.typewrite(str(amount_change_1))
                if len(heat_change_1) > 0:
                    pyautogui.typewrite(['tab'])
                    pyautogui.typewrite(str(heat_change_1))
                    pyautogui.hotkey('shift','tab')
                pyautogui.hotkey('shift','tab')
                time.sleep(0.5)
                
                # then it remembers the last line item it when to so it can go down or up to the next line
                last_line = item_change_1
                
                # this is the loop to read and change the rest of the lines
                for change in range(1,max_range):
                    item_change = item_list[change]
                    amount_change = amount_list[change]
                    heat_change = heat_list[change]
                    # this is to move up or down to the next line
                    line = int(item_change) - last_line
                    # if the line number is bigger than the previos it know to go down
                    if line > 0:
                        for number_of_down in range(line):
                            pyautogui.typewrite(['down'])
                            time.sleep(0.1)
                    # if the line number is less it knows to go up
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
                    stock_text_im = ImageGrab.grab(bbox =(539,arrow_top,599,arrow_bottom))
                    stock_size = tuple(4*x for x in stock_text_im.size)
                    stock_text_im = stock_text_im.resize(stock_size, Image.ANTIALIAS)
                    stock_text_im_gs = stock_text_im.convert('LA')
                    stock_text_string = pytesseract.image_to_string(stock_text_im_gs, config=custom_oem_psm_config)
                    stock_text_list.append(stock_text_string)
                    #wip
                    wip_text_im = ImageGrab.grab(bbox =(252,arrow_top,311,arrow_bottom))
                    wip_size = tuple(4*x for x in wip_text_im.size)
                    wip_text_im = wip_text_im.resize(wip_size, Image.ANTIALIAS)
                    wip_text_im_gs = wip_text_im.convert('LA')
                    wip_text_string = pytesseract.image_to_string(wip_text_im_gs, config=custom_oem_psm_config)
                    wip_text_list.append(wip_text_string)
                    #released
                    released_text_im = ImageGrab.grab(bbox =(600,arrow_top,657,arrow_bottom))
                    released_size = tuple(4*x for x in released_text_im.size)
                    released_text_im = released_text_im.resize(released_size, Image.ANTIALIAS)
                    released_text_im_gs = released_text_im.convert('LA')
                    released_text_string = pytesseract.image_to_string(released_text_im_gs, config=custom_oem_psm_config)
                    released_text_list.append(released_text_string)
                    #on_order
                    on_order_text_im = ImageGrab.grab(bbox =(658,arrow_top,723,arrow_bottom))
                    on_order_size = tuple(4*x for x in on_order_text_im.size)
                    on_order_text_im = on_order_text_im.resize(on_order_size, Image.ANTIALIAS)
                    on_order_text_im_gs = on_order_text_im.convert('LA')
                    on_order_text_string = pytesseract.image_to_string(on_order_text_im_gs, config=custom_oem_psm_config)
                    on_order_text_list.append(on_order_text_string)
                    #stocking
                    stocking_text_im = ImageGrab.grab(bbox =(817,arrow_top,886,arrow_bottom))
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

                    pyautogui.typewrite(['tab'])
                    if '+' not in str(amount_change):
                        pyautogui.typewrite(str(amount_change))
                    if len(heat_change) > 0:
                        pyautogui.typewrite(['tab'])
                        pyautogui.typewrite(str(heat_change))
                        pyautogui.hotkey('shift','tab')
                    pyautogui.hotkey('shift','tab')
                    time.sleep(0.5)

                    # this is to remember the line its at so it can go down or up to the next one
                    last_line = item_change
            
            pyautogui.click(1423,15)
            end_change = time.time()-start_change
            print('\nChange value time: ' + str(round(end_change, 3)) + ' Seconds')

            # put all the lists into a list to make sure the floats have a decimal in them
            list_list = [stock_text_list,wip_text_list,released_text_list,on_order_text_list]
            # loop to cycle through the lists
            for j in range(len(list_list)):
                current_list = list_list[j]
                # loop to cycle through all the values in the list
                for i in range(len(current_list)):
                    string = current_list[i]
                    # if theres not a decimal 4 digits from the right it knows theres not one in there because all values are to 3 decimals
                    if "." not in string[:4]:
                        len_string = len(string)
                        front_string = len_string - 3
                        string = string[:front_string] + "." + string[:3]
                
            # click save
            pyautogui.click(75,65)
            # open transfer window
            pyautogui.click(550,60)
            time.sleep(1)
            # wait till the transfer widow opens
            okWindow = gw.getWindowsWithTitle('Start Manufacturing Order Detail')
            while len(gw.getWindowsWithTitle('Start Manufacturing Order Detail')) == 0:
                time.sleep(1)
            time.sleep(0.5)
            pyautogui.PAUSE = 0.1

            os.chdir(r"C:\Users\jlee.NTPV\Documents\GitHub\material_transfer")
            
            start_transfer = time.time()
            # material transfer loop
            red_x_list = []
            transfer_time_list = []
            wip_after_list = []
            transfers_list = []
            stock_list = []
            amount_transfer_list = []
            wip_list = []
            repeat_item_list = []
            repeat_amount_list = []
            pyautogui.PAUSE = 0.1
            # loop thorugh all the items
            for transfer in range(0,max_range):
                pyautogui.click(645,684)
                transfer_switch = False
                item_transfer = int(item_list[transfer])
                amount_transfer = str(amount_list[transfer])
                stock = str(stock_text_list[transfer])
                part = str(part_list[transfer])
                start_wip = float(wip_text_list[transfer])
                released = float(released_text_list[transfer])
                stocking = str(stocking_text_list[transfer])

                wip = '+'
                wip_switch = False
                
                # if + in the amount transfer remove it
                transfer_add = False
                if '+' in amount_transfer:
                    amount_transfer = amount_transfer.replace("+", "")
                    transfer_add = True

                amount_transfer = float(amount_transfer)

                # if the dtock amount is empty change it to 0
                if len(stock) == 0:
                    stock = '0.000'
                # remove the comma from a number that is over 999
                if ',' in str(stock) and len(stock) >= 9:
                    stock = stock.replace(",","",1)
                # replace commas with decimals
                if ',' in str(stock):
                    stock = stock.replace(",",".")
                # if a space in stock replace with decimal
                if ' ' in str(stock):
                    stock = stock.replace(" ",".")

                stock = float(stock)

                # if the part has already been tranfered the stock wont be correct so i need to subtract the amount tranfered
                if part in transfers_list:
                    for search in range(0,len(transfers_list)):
                        item = transfers_list[search]
                        required = amount_transfer_list[search]
                        stock = float(stock_list[search])
                        past_wip = wip_list[search]
                        #print(item)
                        #print(required)
                        #print(stock)
                        #print(past_wip)
                        if item == part and stock > 0:
                            if '+' in past_wip:
                                stock = float(stock) - float(required)
                            else:
                                stock = float(stock) + float(required)

                #if amount_transfer == 0: # transfer amount is 0
                    #wip = '0'
                    #print('\nTransfer amount is 0')
                    #red_x_list.append('Transfer amount is 0')
                    #transfer_switch = False
                #else:
                transfer_switch = True
                if amount_transfer < start_wip:
                    wip = '-'
                    wip_switch = False
                    pyautogui.click(831,682)
                    amount_transfer = start_wip - float(amount_list[transfer])
                    wip_after = start_wip - amount_transfer
                    red_x_list.append('Too much in WIP')
                else:
                    wip_switch = True
                    wip = '+'
                    
                    if start_wip == amount_transfer: # already in wip
                        wip = '0'
                        print('\nItem {} already in WIP'.format(item_transfer))
                        red_x_list.append('Already in WIP')
                        transfer_switch = False
                    else:
                        if stock == 0:
                            print("\nNone of item {} in stock".format(item_transfer))
                            wip = '0'
                            red_x_list.append('None in stock')
                            if not transfer_add:
                                repeat_item_list.append(item_transfer)
                                repeat_amount_list.append("0")
                            transfer_switch = False
                            amount_transfer = float(0)
                        else:
                            if amount_transfer > stock: #not enough in stock
                                wip = '+'
                                amount_transfer = stock
                                print('\nNot enough of item {} in stock'.format(item_transfer))
                                red_x_list.append('Not enough in stock')
                                if not transfer_add:
                                    repeat_item_list.append(item_transfer)
                                    repeat_amount_list.append(amount_transfer)
                                transfer_switch = True
                                wip_switch = True
                            else:
                                transfer_switch = True
                                if transfer_add:
                                    wip = '+'
                                    wip_switch = True
                                amount_transfer = amount_transfer - start_wip
                                """
                                else:
                                    if amount_transfer > start_wip:
                                        wip_switch = True
                                        wip = '+'
                                        amount_transfer = amount_transfer - start_wip
                                    else:
                                        wip = '-'
                                        wip_switch = False
                                        pyautogui.click(831,682)
                                        amount_transfer = start_wip - float(amount_list[transfer])
                                        red_x_list.append('Too much in WIP')
                                """
                    
                transfer_time = time.strftime('%Y-%m-%d, %H:%M:%S', time.localtime())
                transfer_time_list.append(transfer_time)

                if transfer_switch:
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
                    time.sleep(0.5)
                    if wip_switch:
                        time.sleep(0.25)
                        red_x = pyautogui.locateOnScreen('red_x.png', region=(650,450,100,100))
                        time.sleep(0.25)
                        if red_x is not None:
                            print('\nNot enough of item {} in stock'.format(item_transfer))
                            red_x_list.append('Not enough in stock')
                            wip_after = float(wip_text_list[transfer])
                        else:
                            print('\nEnough of item {} in stock'.format(item_transfer))
                            red_x_list.append('Enough in stock')
                            wip_after = start_wip + amount_transfer

                    pyautogui.click(1034,597)
                    pyautogui.click(868,568)
                    print('\nTransfering {} {} for item {}'.format(round(amount_transfer,3), stocking, item_transfer))
                
                else:
                    wip_after = float(wip_text_list[transfer])
                
                wip_after_list.append(float(wip_after))
                stock_list.append(float(stock))
                amount_transfer_list.append(float(amount_transfer))
                wip_list.append(wip)
                transfers_list.append(part)
            # close transfer window
            pyautogui.click(1004, 776)

            while len(gw.getWindowsWithTitle('Start Manufacturing Order Detail')) != 0:
                time.sleep(0.5)

            go_up()
            
            if len(repeat_item_list) > 0:
                pyautogui.PAUSE = 0.1
                repeat_max_range = len(repeat_item_list)
                pyautogui.doubleClick(219,243)
                if repeat_max_range == 1:
                    item_change = repeat_item_list[0]
                    amount_change = repeat_amount_list[0]
                    line = int(item_change) -1
                    for number_of_down in range(line):
                        pyautogui.typewrite(['down'])
                        time.sleep(0.05)
                    pyautogui.typewrite(str(amount_change))
                    
                else:
                    item_change_1 = repeat_item_list[0]
                    amount_change_1 = repeat_amount_list[0]
                    line_1 = int(item_change_1) -1
                    for number_of_down in range(line_1):
                        pyautogui.typewrite(['down'])
                    pyautogui.typewrite(str(amount_change_1))

                    last_line = item_change_1
                    
                    for change in range(1,repeat_max_range):
                        item_change = repeat_item_list[change]
                        amount_change = repeat_amount_list[change]
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
                            
                        last_line = item_change
                pyautogui.typewrite(['up'])
                pyautogui.typewrite(['down'])
            
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

            time.sleep(0.5)
            pyautogui.click(75,65)
            time.sleep(0.5)

            end_transfer = time.time() - start_transfer
            print('\nTransfer time: ' + str(round(end_transfer, 2)))
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