#! python3
# MISys semi-auto transfer

# coordinates:
# window   (x=1896, y=484)
# line no. (x=590, y=453)
# quantity (x=615, y=611)
# transfer (x=621, y=779)
# ok       (x=1034, y=597)
# IDLE     (x=1890, y=1007)

import pyautogui
import pygetwindow as gw
import time

pyautogui.PAUSE = 0
print('Press Ctrl-C to quit.')

labor = int(input('How many lines need to be skipped?: '))
# start loop to get the line and amount information
# type stop to move on

item_list = []
amount_list = []
job_range = []
stop_loop = False
while not stop_loop:
    item_input = input("Please enter the item number (enter STOP to quit): ")
    
    try:
        if 'STOP' in str(item_input) or 'stop' in str(item_input):
            stop_loop = True
            break
    except ValueError:
        continue
    amount_input = input("Please enter the amount: ")    
    try:
        item_list.append(int(item_input))
        amount_list.append(str(amount_input))
    except ValueError:
        print("Please enter a valid number or STOP to quit")
        continue

# required  (280, 240)
# save      (75, 65)
# transfer  (550, 60)
# swap line (82, 223)
# error     (863, 569)

# ask if there are any lines to change
# if yes then start the change loop
max_range = len(item_list)
pyautogui.click(1890,1007)
skip_me = str(input("Do you have numbers to change? (Y/N): "))
pyautogui.doubleClick(280, 240)
try:
    if 'Y' in skip_me or 'y' in skip_me:
        if max_range ==1:
            item_change = item_list[0]
            amount_change = amount_list[0]
            
            # use the current value to see how many to 'down'
            line = int(labor) + int(item_change) -1
            for number_of_down in range(line):
                pyautogui.typewrite(['down'])

            # material change
            pyautogui.typewrite(str(amount_change))

                    
            # 'up' to top
            for number_of_up in range(line):
                pyautogui.typewrite(['up'])
        
        
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
    labor_range = int(labor_range)-1
    if labor >0:
        for number_of_down_labor in range(labor):
            pyautogui.typewrite(['down'])
    else:
        for number_out_of_labor in range(labor_range):
            pyautogui.typewrite(['down'])
            
    pyautogui.click(1890,1007)
    input("Press ENTER to continue: ")      
    # click save
    pyautogui.click(75,65)
    time.sleep(2)

    # open transfer window
    pyautogui.click(550,60)
    okWindow = gw.getWindowsWithTitle('Start Manufacturing Order Detail')
    while len(gw.getWindowsWithTitle('Start Manufacturing Order Detail')) == 0:
        time.sleep(0.5)
        print("Current value of getwindows: {}".format(len(gw.getWindowsWithTitle('Start Manufacturing Order Detail'))))

# material transfer loop
    for transfer in range(0,max_range):

        item_transfer = item_list[transfer] + labor
        amount_transfer = amount_list[transfer]
        
        # click in IDLE window for inputs
        #pyautogui.click(1890,1007)
        
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
            print("Current value of getwindows: {}".format(len(gw.getWindowsWithTitle('Start Manufacturing Order Detail'))))

        print('OK done.')
        time.sleep(0.5)
        pyautogui.click(1034,597)
        
# Have a kill switch
except KeyboardInterrupt:
    print('\nDone')
# close transfer window
pyautogui.click(1004, 776)
print('\nComplete.')
