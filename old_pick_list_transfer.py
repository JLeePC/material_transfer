#! python3
# MISys semi-auto transfer

import pyautogui
import pygetwindow as gw
import time
import pyperclip
import csv
import os
from file_routing import file_routing
from copy_clipboard import copy_clipboard

class EmptyInput(Exception):
    pass

ask_for_next_job = False
pyautogui.PAUSE = 0.1
print('Ctrl_c to quit.')

try:
    while True:
        pyautogui.PAUSE = 0.1
        # Go to next job
        if ask_for_next_job is True:
            pyautogui.click(1423,15)
            next_job = input('\nEnter next job: ')
            if len(next_job) <= 4:
                # click window then type next job with zfill
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
        released_button = pyautogui.locateOnScreen('released.png')
        time.sleep(0.25)

        if released_button is not None:

            pyautogui.click(150,150)
            print('\nLooking for side bar.')
            button_location = pyautogui.locateOnScreen('up_arrow.png', region=(875,207,29,42))
            time.sleep(1)
            
            if button_location is None:
                print('Up arrow not found')
            else:
                print('Up arrow found')
                pyautogui.click(890,250, clicks=4, interval=0.05)
                pyautogui.click(890,227)
                
            if button_location is not None:
                print('\nLooking for LABOR.')

                labor1 = pyautogui.locateOnScreen('labor1.png', region=(110,232,134,737))
                time.sleep(0.1)
                labor2 = pyautogui.locateOnScreen('labor2.png', region=(110,232,134,737))
                time.sleep(0.1)

                if labor1 or labor2 is not None:
                    page = 1
                    labor_lines = True
                    print('There is LABOR in this job.')
                else:
                    page = 2
                    labor_loop = True
                    while labor_loop:
                        pyautogui.click(891,961, clicks=35, interval=0.05) # down arrow
                        time.sleep(0.25)
                        labor1 = pyautogui.locateOnScreen('labor1.png', region=(110,232,134,737))
                        labor2 = pyautogui.locateOnScreen('labor2.png', region=(110,232,134,737))
                        bottom_arrow = pyautogui.locateOnScreen('bottom_arrow.png', region=(880,940,30,40))
                        if labor1 or labor2 is not None:
                            labor_lines = True
                            labor_loop = False
                            print('There is LABOR in this job.')
                            break
                        if bottom_arrow is not None:
                            labor_lines = False
                            labor_loop = True
                            print('There is no LABOR in this job.')
                            break

            else:
                print('\nLooking for LABOR.')
                labor1 = pyautogui.locateOnScreen('labor1.png')
                labor2 = pyautogui.locateOnScreen('labor2.png')
                if labor1 or labor2 is not None:
                    labor_lines = True
                    print('There is LABOR in this job.')
                else:
                    labor_lines = False
                    print('There is no LABOR in this job.')
                
            pyautogui.PAUSE = 0.1

            labor_no = str('LABOR')
            
            if button_location is not None and labor_lines is True:
                pyautogui.click(890,948, clicks=4, interval=0.05)
                pyautogui.click(890,964)
                pyautogui.doubleClick(219,242)
                current_part_no2 = copy_clipboard()
                current = current_part_no2.split('-')
                if labor_no in current:
                    print('Labor at bottom')
                    labor_lines = False
                pyautogui.click(890,250, clicks=4, interval=0.05)
                pyautogui.click(890,227)
            elif button_location is None and labor_lines is True:
                pyautogui.doubleClick(219,242)
                down_loop = True
                current_part_no4 = copy_clipboard()
                current = current_part_no4.split('-')
                if labor_no in current:
                    print('Labor at top.')
                    down_loop = False
                    labor_lines = True
                else:
                    while down_loop:
                        pyautogui.PAUSE = 0.03
                        im1 = pyautogui.screenshot()
                        for i in range(10):
                            pyautogui.typewrite(['down'])
                        im2 = pyautogui.screenshot()
                        if im1 == im2:
                            down_loop = False
                            break
                        current_part_no3 = copy_clipboard()
                        current = current_part_no3.split('-')
                    if labor_no in current:
                        labor_lines = False
                        print('Labor at bottom')
                            
            if labor_lines:
                if button_location is not None:
                    pyautogui.doubleClick(892,250)
                    pyautogui.doubleClick(892,250)
                    pyautogui.click(890,227)
            if labor_lines:
                print('\nLooking for LABOR lines.')
                pyautogui.doubleClick(219,242)
                labor = False
                total_lines = 0
                while labor is False:
                    current_part_no1 = copy_clipboard()
                    current = current_part_no1.split('-')
                    if labor_no in current:
                        pyautogui.click(919,291)
                        labor = True
                        time.sleep(.5)
                        break
                    pyautogui.typewrite(['down'])
                    pyautogui.PAUSE = 0.1
                    total_lines = total_lines + 1
                first_labor = current_part_no1
                print('First LABOR line is: ' + str(first_labor))

                if button_location is not None:
                    pyautogui.doubleClick(892,250)
                    pyautogui.doubleClick(892,250)
                    pyautogui.click(890,227)
                    time.sleep(0.25)
                pyautogui.doubleClick(219,242)
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
                        if button_location is not None:
                            pyautogui.PAUSE = 0.1
                            pyautogui.doubleClick(892,250)
                            pyautogui.doubleClick(892,250)
                            pyautogui.click(890,227)
                        print('\nFound labor copy.')
                        break
                    if labor_no in current:
                        pyautogui.click(919,291)
                        time.sleep(.1)
                        if button_location is not None:
                            pyautogui.PAUSE = 0.1
                            pyautogui.doubleClick(892,250)
                            pyautogui.doubleClick(892,250)
                            pyautogui.click(890,227)
                        pyautogui.doubleClick(219,242)
                        print('Moved ' + str(current_part_no))
                        total_down = True
                        time.sleep(0.5)
                        continue
                    pyautogui.PAUSE = 0.05
                    print('Current part: ' +current_part_no)
                    pyautogui.typewrite(['down'])
                    total_down = False

            pyautogui.doubleClick(219,242)
            pwht = copy_clipboard()
            time.sleep(0.1)
            if 'PWHT' in pwht:
                pyautogui.click(919,291)

            if button_location is not None:
                pyautogui.PAUSE = 0.1
                pyautogui.doubleClick(892,250)
                pyautogui.doubleClick(892,250)
                pyautogui.click(890,227)
                pyautogui.PAUSE = 0.05

            pyautogui.PAUSE = 0.05
            pyautogui.click(1423,15)
            line_skip = int(input('How many lines need to be skipped?: '))
            print('\nPress Ctrl-C to quit.')
            # start loop to get the line and amount information
            # type stop to move on
            item_list = []
            amount_list = []
            heat_list = []
            last_heat_number = ''
            #job_range = []
            stop_loop = False
            while not stop_loop:
                item_input = input("Please enter the item number (enter + to quit): ")
                try:
                    if '-' in item_input:
                        length = len(item_list) - 1
                        item_list.pop(length)
                        amount_list.pop(length)
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

                heat_number = input("Please enter the heat number: ")

                if '+' in heat_number:
                    heat_number = last_heat_number
                if len(heat_number) == 0:
                    heat_number = 'None'

                item_list.append(int(item_input))
                heat_list.append(str(heat_number.upper()))
                last_heat_number = heat_number
                
            if button_location is not None:
                pyautogui.PAUSE = 0.1
                pyautogui.doubleClick(892,250)
                pyautogui.doubleClick(892,250)
                pyautogui.click(890,227)
                pyautogui.PAUSE = 0.05
            # ask if there are any lines to change
            # if yes then start the change loop
            pyautogui.PAUSE = 0.1
            max_range = len(item_list)
            pyautogui.doubleClick(219,242)
            part_list = []
            if max_range ==1:
                item_change = item_list[0]
                # use the current value to see how many to 'down'
                line = int(line_skip) + int(item_change) -1
                for number_of_down in range(line):
                    pyautogui.typewrite(['down'])
                    time.sleep(0.05)
                # material change
                part_no = copy_clipboard()
                time.sleep(0.25)
                print(part_no)
                part_list.append(str(part_no))
                
            else:
                item_change_1 = item_list[0]
                line_1 = int(line_skip) + int(item_change_1) -1
                for number_of_down in range(line_1):
                    pyautogui.typewrite(['down'])
                # material change
                time.sleep(0.2)
                part_no = copy_clipboard()
                time.sleep(0.2)
                print(part_no)
                part_list.append(str(part_no))

                last_line = item_change_1
                
                for change in range(1,max_range):
                    item_change = item_list[change]
                    # use the current value to see how many to 'down'
                    line = int(item_change) - last_line
                    if line > 0:
                        for number_of_down in range(line):
                            pyautogui.typewrite(['down'])
                    elif line < 0:
                        for number_of_up in range(abs(line)):
                            pyautogui.typewrite(['up'])
                    # material change
                    time.sleep(0.2)
                    part_no = copy_clipboard()
                    time.sleep(0.25)
                    print(part_no)
                    part_list.append(str(part_no))
                    last_line = item_change

            if button_location is not None:
                pyautogui.PAUSE = 0.1
                pyautogui.doubleClick(892,250)
                pyautogui.doubleClick(892,250)
                pyautogui.click(890,227)
                pyautogui.PAUSE = 0.05

            max_range = len(item_list)
            pyautogui.doubleClick(279,245)
            amount_list = []
            if max_range ==1:
                item_change = item_list[0]
                # use the current value to see how many to 'down'
                line = int(line_skip) + int(item_change) -1
                for number_of_down in range(line):
                    pyautogui.typewrite(['down'])
                    time.sleep(0.05)
                # material change
                part_no = copy_clipboard()
                time.sleep(0.25)
                print(part_no)
                amount_list.append(float(part_no))
                
            else:
                item_change_2 = item_list[0]
                line_2 = int(line_skip) + int(item_change_2) -1
                for number_of_down in range(line_2):
                    pyautogui.typewrite(['down'])
                # material change
                time.sleep(0.2)
                part_no = copy_clipboard()
                time.sleep(0.2)
                print(part_no)
                amount_list.append(float(part_no))

                last_line = item_change_2
                
                for change in range(1,max_range):
                    item_change = item_list[change]
                    # use the current value to see how many to 'down'
                    line = int(item_change) - last_line
                    if line > 0:
                        for number_of_down in range(line):
                            pyautogui.typewrite(['down'])
                    elif line < 0:
                        for number_of_up in range(abs(line)):
                            pyautogui.typewrite(['up'])
                    # material change
                    time.sleep(0.2)
                    part_no = copy_clipboard()
                    time.sleep(0.25)
                    print(part_no)
                    amount_list.append(float(part_no))
                    last_line = item_change
                            
            # if amount to skip is >0 then go down past labor
            labor_range = item_list[0]
            labor_range = int(labor_range)
            if len(item_list) == 0:
                for number_of_down_labor in range(line_skip):
                    pyautogui.typewrite(['down'])
            wip = '+'
            pyautogui.PAUSE = 0.1
        # material transfer loop
            red_x_list = []
            transfer_time_list = []
            start_transfer = time.time()
            for transfer in range(0,max_range):
                item_transfer = item_list[transfer] + line_skip
                amount_transfer = amount_list[transfer]
                
                transfer_time = time.strftime('%Y-%m-%d, %H:%M:%S', time.localtime())
                
                red_x = None

                time.sleep(0.25)

                transfer_time_list.append(transfer_time)
                
                if red_x is not None:
                    red_x_list.append('None in stock')

                else:
                    red_x_list.append('In stock')

            while len(gw.getWindowsWithTitle('Start Manufacturing Order Detail')) != 0:
                time.sleep(0.5)

            pyautogui.doubleClick(210,101)
            job_no = copy_clipboard()
            file_routing()

            current_date = time.strftime('%Y-%m-%d, %Hh %Mm %Ss', time.localtime())

            with open('{}.csv'.format(current_date), 'w', newline='') as csv_file:
                fieldnames = ['Time','Job number','+/-','Item','Part Number','Heat number','Amount']
                csv_writer = csv.DictWriter(csv_file, fieldnames=fieldnames)

                csv_writer.writeheader()

                for write in range(0,max_range):

                    item_writer = item_list[write]
                    amount_writer = amount_list[write]
                    part_writer = part_list[write]
                    job_writer = job_no
                    transfer_time_writer = transfer_time_list[write]
                    heat_number_writer = heat_list[write]
                    
                    csv_writer.writerow({'Time': transfer_time_writer, 'Job number': job_writer, '+/-': wip,
                                         'Item': item_writer, 'Part Number': part_writer,
                                         'Heat number': heat_number_writer, 'Amount': amount_writer})


            os.chdir(r"D:\MIsys Data")

            with open('Old_Master.csv', 'a', newline='') as csv_file:
                fieldnames = ['Time','Job number','+/-','Item','Part Number','Heat number','Amount']
                csv_writer = csv.DictWriter(csv_file, fieldnames=fieldnames)

                #csv_writer.writeheader()

                for write in range(0,max_range):

                    item_writer = item_list[write]
                    amount_writer = amount_list[write]
                    part_writer = part_list[write]
                    job_writer = job_no
                    transfer_time_writer = transfer_time_list[write]
                    heat_number_writer = heat_list[write]
                    
                    csv_writer.writerow({'Time': transfer_time_writer, 'Job number': job_writer, '+/-': wip,
                                         'Item': item_writer, 'Part Number': part_writer,
                                         'Heat number': heat_number_writer, 'Amount': amount_writer})

            os.chdir(r"C:\Users\jlee.NTPV\Documents\GitHub\material_transfer")

            end_transfer = time.time() - start_transfer
            print('\nTransfer time: ' + str(round(end_transfer, 2)))
            total_time = time.time() - total_start
            minutes = 0
            while total_time >= 60:
                total_time = total_time - 60
                minutes = minutes + 1
            
            print('\nTotal Time: ' + str(minutes) + ' Minutes ' + str(round(total_time, 2)) + ' Seconds')
            print('\nCompleted Loop.')
            print('----------------------------------------')

        else:
            print('The job is closed.')
        ask_for_next_job = True

# Have a kill switch
except KeyboardInterrupt:
    print('\nDone')

