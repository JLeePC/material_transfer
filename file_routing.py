import os
import time
import pyautogui
import pyperclip
from copy_clipboard import copy_clipboard

def file_routing():

    job = copy_clipboard()

    job = job.upper()

    os.chdir(r"D:\MIsys Data\Jobs")

    os_list = os.listdir('.')

    current_date = time.strftime('%Y-%m-%d, %Hh %Mm %Ss', time.localtime())

    current_year = time.strftime('%Y', time.localtime())

    current_month = time.strftime('%m', time.localtime())

    os_list = os.listdir('.')

    if job not in os_list:
        os.mkdir(r"D:\MIsys Data\Jobs\{}".format(job))

    os.chdir(job)

    os_list = os.listdir('.')

    if current_year not in os_list:
        os.mkdir(r"D:\MIsys Data\Jobs\{}\{}".format(job,current_year))

    os.chdir(current_year)

    os_list = os.listdir('.')

    if current_month not in os_list:
        os.mkdir(r"D:\MIsys Data\Jobs\{}\{}\{}".format(job,current_year,current_month))

    os.chdir(current_month)

    print("\n" + os.getcwd())
