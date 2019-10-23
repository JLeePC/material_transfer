#! python3
# MISys semi-auto transfer

# coordinates:
# window   (x=1896, y=484)
# line no. (x=590, y=453)
# quantity (x=615, y=611)
# transfer (x=621, y=779)
# ok       (x=1034, y=597)
# IDLE     (x=1890, y=1007)
# required  (280, 240)
# save      (75, 65)
# transfer  (550, 60)
# swap line (82, 223)
# error     (863, 569)
# wip->stock(831, 682)

import pyautogui
import pygetwindow as gw
import time
import pyperclip

ask_for_next_job = False
total_start = time.time()
pyautogui.PAUSE = 0.1
pyautogui.click(150,150)

def copy_clipboard():
    pyautogui.hotkey('ctrl', 'c')
    time.sleep(.01)
    return pyperclip.paste()

try:
    while True:
        # Go to next job
        if ask_for_next_job is True:
            pyautogui.click(1890,1007)
            print('\nCtrl_c to quit.')
            next_job = input('Enter next job: ')
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
                
        pyautogui.click(1423,15)
        time.sleep(0.5)
        change_part = input('Are there part numbers that need to change? ')
        
        if '1' in str(change_part):
            pyautogui.PAUSE = 0.1
            change_part = True
            while change_part:
                pyautogui.click(1423,15)
                change_line = input('Which line needs to change?: ').zfill(2)
                new_part = input('Whats the new part number?: ')
                if '01' in change_line:
                    pyautogui.doubleClick(219,242)
                    pyautogui.click(919,289) # bottom
                    time.sleep(0.25)
                    pyautogui.typewrite(['tab'])
                    pyautogui.typewrite('0')
                    time.sleep(0.25)
                    pyautogui.click(920,312) # new line
                    pyautogui.typewrite(str(new_part))
                    pyautogui.click(920,227) # top
                else:
                    pyautogui.doubleClick(219,242)
                    change_line_range = int(change_line) - 1
                    for number_of_down in range(change_line_range):
                        pyautogui.PAUSE = 0.05
                        pyautogui.typewrite(['down'])
                    pyautogui.click(919,289) # bottom
                    time.sleep(0.25)
                    pyautogui.typewrite(['tab'])
                    pyautogui.typewrite('0')
                    time.sleep(0.25)
                    pyautogui.click(920,312) # new line
                    pyautogui.typewrite(str(new_part))
                    pyautogui.click(920,227) # top
                    for number_of_down in range(change_line_range):
                        pyautogui.click(920,269)
                pyautogui.click(1423,15)
                change_more = input('Are there more lines to change?: ')
                if '0' in change_more:
                    change_part = False
                    break
        
        pyautogui.click(1423,15)
        pyautogui.PAUSE = 0.1
        labor_lines = input('Are there labor lines? ')
        if not '0' in str(labor_lines):
            right_bar = input('Is there a scroll bar? ')
            if '1' in right_bar:
                pyautogui.doubleClick(892,250)
                pyautogui.doubleClick(892,250)
                pyautogui.click(890,227)
        
        if not '0' in str(labor_lines):
            print('Looking for LABOR lines.')
            pyautogui.doubleClick(219,242)
            # go till first labor then move that to the bottom
            labor_no = str('LABOR')
            labor = False
            while labor is False:
                current_part_no = copy_clipboard()
                current = current_part_no.split('-')
                if labor_no in current:
                    pyautogui.click(919,291)
                    labor = True
                    time.sleep(.5)
                    break
                pyautogui.typewrite(['down'])
            first_labor = current_part_no
            print('First LABOR line is: ' + str(first_labor))
            # go down till you meet that same labor line to see total amount of lines
            if '1' in right_bar:
                pyautogui.doubleClick(892,250)
                pyautogui.doubleClick(892,250)
                pyautogui.click(890,227)
            pyautogui.doubleClick(219,242)
            labor_copy = False
            total_lines = 0
            while labor_copy is False:
                current_part_no = copy_clipboard()
                current = current_part_no.split('-')
                if current_part_no == first_labor:
                    labor_copy = True
                    if '1' in right_bar:
                        pyautogui.PAUSE = 0.1
                        pyautogui.doubleClick(892,250)
                        pyautogui.doubleClick(892,250)
                        pyautogui.click(890,227)
                    print('Found labor copy after ' + str(total_lines) + ' lines.')
                    break
                if labor_no in current:
                    pyautogui.click(919,291)
                    time.sleep(.5)
                    if '1' in right_bar:
                        pyautogui.PAUSE = 0.1
                        pyautogui.doubleClick(892,250)
                        pyautogui.doubleClick(892,250)
                        pyautogui.click(890,227)
                    pyautogui.doubleClick(219,242)
                    total_lines = 0
                    print('Moved ' + str(current_part_no))
                    time.sleep(0.25)
                    continue
                pyautogui.PAUSE = 0.05
                total_lines = total_lines + 1
                pyautogui.typewrite(['down'])
        pyautogui.PAUSE = 0.03
        print('\nPress Ctrl-C to quit.')
        pyautogui.click(1890,1007)
        labor = int(input('How many lines need to be skipped?: '))
        # start loop to get the line and amount information
        # type stop to move on
        item_list = []
        amount_list = []
        #job_range = []
        stop_loop = False
        while not stop_loop:
            item_input = input("Please enter the item number (enter + to quit): ")
            try:
                if '+' in str(item_input):
                    stop_loop = True
                    break
            except ValueError:
                continue
            amount_input = input("Please enter the amount: ")    
            try:
                item_list.append(int(item_input))
                amount_list.append(str(amount_input))
            except ValueError:
                print("Please enter a valid number or + to quit")
                continue
        # ask if there are any lines to change
        # if yes then start the change loop
        max_range = len(item_list)
        pyautogui.click(1890,1007)
        skip_me = input("Do you have numbers to change? (1/0): ")
        pyautogui.doubleClick(280, 240)
        reset = item_list[0]
        for up_reset in range(reset):
            pyautogui.typewrite(['up'])
        if '1' in str(skip_me):
            start_change = time.time()
            if max_range ==1:
                item_change = item_list[0]
                amount_change = amount_list[0]
                # use the current value to see how many to 'down'
                line = int(labor) + int(item_change) -1
                for number_of_down in range(line):
                    pyautogui.typewrite(['down'])
                    time.sleep(0.05)
                # material change
                pyautogui.typewrite(str(amount_change))        
                # 'up' to top
                for number_of_up in range(line):
                    pyautogui.typewrite(['up'])
            else:
                for change in range(0,max_range):
                    item_change = item_list[change]
                    amount_change = amount_list[change]
                    # use the current value to see how many to 'down'
                    line = int(labor) + int(item_change) -1
                    for number_of_down in range(line):
                        pyautogui.typewrite(['down'])
                    # material change
                    pyautogui.typewrite(str(amount_change))
                    # 'up' to top
                    for number_of_up in range(line):
                        pyautogui.typewrite(['up'])
        # if amount to skip is >0 then go down past labor
        labor_range = item_list[0]
        labor_range = int(labor_range)
        if labor >0:
            for number_of_down_labor in range(labor):
                pyautogui.typewrite(['down'])
        else:
            for number_out_of_labor in range(labor_range):
                pyautogui.typewrite(['down'])
            pyautogui.typewrite(['up'])
        pyautogui.click(1890,1007)
        if '1' in str(skip_me):
            end_change = time.time()-start_change
            print('\nChange value time: ' + str(round(end_change, 3)) + ' Seconds')
        
        #placeholder = []
        #num_to_skip = []
        #stop_loop = False
        # skip_line = str(input("Do you have numbers to skip? (1/0): "))
        #skip_line = '0'
        #if '1' in skip_line:
            # stop_loop is a secondary measure to prevent infinite loops, not required, but precautionary
            #while not stop_loop:
                #user_input = input("Please enter the number you would like to skip (enter + to quit): ")
                #try:
                    #if '+' in str(user_input):
                        #stop_loop = True
                        #break
                #except ValueError:
                    #continue
                    
                #try:
                    #placeholder.append(int(user_input))
                #except ValueError:
                    #print("Please enter a valid number or + to quit")
                    #continue
            # We need to remove possible duplicates
            #for num in placeholder:
                #if num not in num_to_skip:
                    #num_to_skip.append(num)
        # click save
        pyautogui.click(75,65)
        time.sleep(0.5)
        # open transfer window
        pyautogui.click(550,60)
        okWindow = gw.getWindowsWithTitle('Start Manufacturing Order Detail')
        while len(gw.getWindowsWithTitle('Start Manufacturing Order Detail')) == 0:
            time.sleep(0.25)
        time.sleep(0.25)
        pyautogui.click(1800,1007)
        print('\nCtrl-C to quit')
        wip = input('Which direction do you want to transfer? (+ = to WIP/- = to Stock): ')
        
        if '-' in wip:
            pyautogui.click(831,682)
        start_transfer = time.time()
    # material transfer loop
        for transfer in range(0,max_range):
            item_transfer = item_list[transfer] + labor
            amount_transfer = amount_list[transfer]
            
            #if item_transfer in num_to_skip:
                #continue
            if float(amount_transfer) == 0:
                continue
            
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
            # ok window
            okWindow = gw.getWindowsWithTitle('Start Manufacturing Order Detail')
            while len(gw.getWindowsWithTitle('Start Manufacturing Order Detail')) == 1:
                time.sleep(0.25)
                # print("Current value of getwindows: {}".format(len(gw.getWindowsWithTitle('Start Manufacturing Order Detail'))))
            time.sleep(0.25)
            pyautogui.click(1034,597)
            pyautogui.click(868,568)
            print('Item ' + str(item_transfer) + ' Done.')
        # close transfer window
        pyautogui.click(1004, 776)
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
        while len(gw.getWindowsWithTitle('Start Manufacturing Order Detail')) != 0:
                time.sleep(0.5)
        print('\nCompleted Loop.')
        print('----------------------------------------')
        ask_for_next_job = True
# Have a kill switch
except KeyboardInterrupt:
    print('\nDone')

