import pyautogui
import requests, re
from time import sleep
from datetime import *
from selenium import webdriver
from selenium.webdriver.common.by import By
from pprint import pprint
import numpy as np
from bs4 import BeautifulSoup
import pyodbc
import pandas as pd
import os
from os import listdir
import psutil
from psutil import disk_usage
from os.path import isfile, join
from wakepy import set_keepawake, unset_keepawake
import ctypes

# vars
free_space_limit = 2
# U is Archive01 Z is Archive02 (overflow as needed more space)
# storage_location = "U:\\Dicom\\"
storage_location = "Z:\\Dicom\\"

# SQL connection
conn = pyodbc.connect('Driver={SQL Server};'
                      r'Server=UHLSQLPRIME01\UHLBRICCSDB;'
                      'Database=i2b2_app03_b1_data;'
                      'Trusted_Connection=yes;')
# options
pd.set_option("display.max_columns", None)
pd.set_option("expand_frame_repr", True)
pd.set_option("display.width", 220)
pyautogui.FAILSAFE = False


def storage_check():
    global free_space
    free_space = round(psutil.disk_usage(storage_location).free / 1000000000, 1)
    print('free space on storage drive:', free_space, ' GB')
    if free_space < free_space_limit:
        raise Exception('insufficient storage')


def build_lists_of_to_do():
    global list_to_dicom
    sql_list_to_dicom = pd.read_sql_query('SELECT * FROM i2b2_app03_b1_data.dbo.DICOM_List', conn)
    df = pd.read_csv(storage_location + 'results.csv',
                     parse_dates=['date_time_finished', 'date_time_opened'],
                     dtype={'RWES': str,
                            'BptNumber': str,
                            'number_of_Dicoms_on_right_Date': int
                            })
    sql_list_to_dicom.shape
    # list of all that have been done:
    list_plus_completed_detail = pd.merge(sql_list_to_dicom, df, how="left", on=["BptNumber"])
    list_plus_completed_detail.shape
    # drop where none on the right date is found
    list_plus_completed_detail = list_plus_completed_detail.drop(
        list_plus_completed_detail[list_plus_completed_detail.number_of_Dicoms_on_right_Date == 0].index)
    list_plus_completed_detail.shape
    # drop the completed participants
    list_plus_completed_detail = list_plus_completed_detail[list_plus_completed_detail['date_time_finished'].isnull()]
    list_plus_completed_detail.shape
    # final list
    list_to_dicom = list_plus_completed_detail
    # Resetting the index as some have now been removed which would mess with the iteration loop
    list_to_dicom.reset_index(inplace=True, drop=True)
    print(list_to_dicom.columns)
    print(list_to_dicom.head(7))
    print('How may outstanting?')
    print(list_to_dicom.index.max())


def log_me_in():
    global driver, f
    driver = webdriver.Ie()
    driver.implicitly_wait(3)
    driver.get("https://uvweb.pacs.uhl-tr.nhs.uk/login.jsp")
    driver.maximize_window()
    f = open("G:\dicom\.env", "r")
    myid = f.readline().splitlines()
    mypass = f.readline()
    f.close()
    sleep(1)
    username = driver.find_element(By.ID, "userName")  # find_element_by_id
    username.send_keys(myid)
    sleep(1)
    password = driver.find_element(By.ID, "password")  # find_element_by_id
    password.send_keys(mypass)
    login = driver.find_element(By.NAME, "login")  # find_element_by_name
    login.click()


def close_study():
    sleep(2)
    pyautogui.keyDown('alt')
    pyautogui.press('f4')
    pyautogui.keyUp('alt')
    sleep(3)
    pyautogui.press('return')
    sleep(1)


storage_check()
build_lists_of_to_do()
log_me_in()

pacsWindow = driver.window_handles
print(pacsWindow)

################ start iterations

finish_line = list_to_dicom.index.max()
print(finish_line)
i = 0

while i < finish_line and free_space >= free_space_limit:
    print("starting loop " + str(i))
    NextInList = ""
    NextInList_bpt = ""
    date_to_find = ""
    print("variables made")
    NextInList = list_to_dicom.at[i, 'MRN']
    print("NextInList made")
    NextInList_bpt = list_to_dicom.at[i, 'BptNumber']
    print("NextInList_bpt made")
    date_to_find = datetime.utcfromtimestamp(
        list_to_dicom['ct_date_time_start'].values[i].astype(datetime) / 1_000_000_000).strftime('%m-%d-%Y')
    print("date_to_find made")
    print(NextInList)
    print(NextInList_bpt)
    print(date_to_find)
    # open advanced search and find its handle and switch to window
    pacsWindow = driver.window_handles
    print(pacsWindow)
    driver.current_url
    driver.switch_to.window(pacsWindow[0])
    driver.current_url
    driver.execute_script("openSearch(false);")
    sleep(1)
    advsearchwindow = list(set(driver.window_handles) - set(pacsWindow))[0]
    driver.window_handles
    print(advsearchwindow)
    print(pacsWindow)
    driver.switch_to.window(
        advsearchwindow)  # driver.switch_to.window(PacsWindow) #driver.switch_to.window(advsearchwindow)
    driver.switch_to.frame("LIST")
    ########## Enter details for search and submit
    sleep(1)
    driver.find_element(By.NAME, "searchPatientId").send_keys(NextInList)  # find_element_by_name x 3 for next few lines
    driver.find_element(By.NAME,
                        "searchFirstName").click()  # click away from searchPatientId to enable searchStudyDescr
    # driver.execute_script("dialogForm.searchOrderStatus[8].click();") # sets to complete, however we mhy also want read offline...
    driver.execute_script("dialogForm.searchImgCnt.click();")  # only studys with images
    driver.find_element(By.NAME, "searchStudyDescr").send_keys("CT Cardiac angiogram")
    driver.switch_to.window(advsearchwindow)
    driver.switch_to.frame("TOOLBAR")
    butts = driver.find_elements(By.CLASS_NAME, "search_button")  # find_elements_by_class_name
    butts[0].click()
    sleep(3)
    # driver.window_handles by monitoring the window_handles it seems to close the connection!!!!!!!!!!
    print(pacsWindow)
    print(driver.window_handles)
    print("switching...PacsWindow")
    driver.switch_to.window(pacsWindow[0])
    driver.switch_to.frame("tableFrame")
    soup = BeautifulSoup(driver.page_source, 'lxml')
    print('looking for date match')
    print(date_to_find)
    list_to_dicom.head(2)
    pprint(soup.find('td', string=re.compile(date_to_find)))
    print('looking for date match end')
    number_of_Dicoms_on_right_Date = len(soup.find_all('td', string=re.compile(date_to_find)))
    print('number of Dicoms on date of interest')
    print(number_of_Dicoms_on_right_Date)
    # Now to record the number of in range (should only be 1 for yes or 0 for nune however it's possible there are
    # more then one.
    # if there is only valid one in the list, create new folder and Select
    if number_of_Dicoms_on_right_Date == 1:
        path = os.path.join(storage_location, NextInList_bpt)
        print("folder to save is")
        print(path)
        if not os.path.exists(path):
            os.makedirs(path)
        listTableForm = driver.find_element(By.NAME, "listTableForm")
        d_found_at = soup.find('td', string=re.compile(date_to_find))
        # images_toProcess is reliant on the img column being two to the right of the date column
        images_to_process = int(d_found_at.find_next_sibling().find_next_sibling().string)
        sleep(1)
        listTableForm.click()
        continue_to_extract = 1
    else:
        continue_to_extract = 0
        to_log = np.array(
            [NextInList + ',' + NextInList_bpt + ',' + str(number_of_Dicoms_on_right_Date) + ',' + str(
                datetime.now()) + ',,'])
        print(to_log)
        with open(storage_location + "results.csv", "ab") as f:
            np.savetxt(f, (to_log), fmt='%s', delimiter=' ')
    ########## Export the image
    if continue_to_extract == 1:
        sleep(images_to_process / 75)  # page loading, for participants with loads of images this can take some time.
        sleep(10)  # when there is VERY few the above line will not be enough
        pyautogui.press('f12')  # opens the save images windows, focused on save button by default
        sleep(2)
        pyautogui.typewrite('d')  # set file type to DICOM(*.dcm)
        sleep(1)
        pyautogui.press('tab')  # tab over to folder
        sleep(.1)
        pyautogui.press('tab')
        sleep(.1)
        pyautogui.press('tab')
        sleep(.1)
        pyautogui.press('tab')
        sleep(.1)
        pyautogui.typewrite(path)  # set to predefined path
        sleep(1)
        pyautogui.press('tab')  # tab apply to images
        sleep(.1)
        pyautogui.press('down')
        sleep(.1)  # select the non-default save 'Entire Study' (two down from default)
        pyautogui.press('down')
        sleep(.1)
        pyautogui.press('tab')
        sleep(.1)
        pyautogui.press('tab')  # should save you're last settings so shouldn't need to check not create subfolder
        # sleep(.1)
        # pyautogui.press('space')           # click not create subfolder      # select File Name Header and type bpt number in
        sleep(1)
        pyautogui.typewrite(NextInList_bpt)
        # option - anonymise should always be defult and checked
        pyautogui.press('tab')  # tab to the save button
        starting_download = datetime.now()
        pyautogui.press('return')  # time to save the files!
        # take about ten minutes to save the stuff to C drive so wait at least 5 mins before starting to check
        sleep(120)
        images_to_do = images_to_process
        # keep checking for finished!
        while images_to_do > 1:
            number_of_dicoms_downloaded = len([f for f in listdir(path) if isfile(join(path, f))])
            # occasionally an extra file is found and downloaded hence the need to prevent a negative number.
            images_to_do = max(images_to_process - number_of_dicoms_downloaded, 0)
            message = str(number_of_dicoms_downloaded) + ' of ' + str(images_to_process) + ' downloaded, ' + str(
                images_to_do) + ' to do.'
            print(message)
            print('Time now:', str(datetime.now()))
            # for setting sleep, it takes about .7 seconds per file, this means we've check just before the extract is
            # due to finish, this prevents output going overboard. + 2 to ensure near completion it's not logging many
            # near the end.
            timer = round(images_to_do * .7, 0)

            print('Timer set to:', str(timer))
            # The next while loop should stop the screen from locking (which messes up the program) by
            # keyboard interaction every minute, until it's next time to check .
            while timer > 60:
                # sleep(1)
                # The below wasn't preventing the screen lock kicking in so has been commented out.
                # print('Timer set to:', str(timer), ' ...sleeping for a min')
                sleep(60)
                # pyautogui.press('volumedown')
                # sleep(.5)
                # pyautogui.press('volumeup')
                # sleep(.5)
                ES_CONTINUOUS = 0x80000000
                ES_SYSTEM_REQUIRED = 0x00000001
                ES_DISPLAY_REQUIRED = 0x00000002
                # ES_AWAYMODE_REQUIRED = 0x00000040
                # #pyautogui.typewrite('')
                ##pyautogui.moveRel(xOffset, yOffset, duration=num_seconds)
                # pyautogui.moveRel(20, 30, duration=.5)
                # pyautogui.moveRel(-20, -30, duration=.5)
                timer = max(timer - 60, 1)
            sleep(timer)
            print('Timer set to:', str(timer), ' ...finished sleep.')
        finished_downloading = datetime.now()
        to_log = np.array(
            [NextInList + ',' + NextInList_bpt + ',' + str(number_of_Dicoms_on_right_Date) + ',' +
             str(starting_download) + ',' + str(datetime.now()) + ','])
        print(to_log)
        with open(storage_location + "results.csv", "ab") as f:
            np.savetxt(f, (to_log), fmt='%s', delimiter=' ')
        download_took = finished_downloading - starting_download
        close_study()
        print(NextInList_bpt + " finished! Time taken(h:mm:ss.ms):", download_took)
    i = i + 1

    free_space = round(psutil.disk_usage(storage_location).free / 1000000000, 1)
    print('free space on storage drive:', free_space, ' GB')
    if free_space < free_space_limit:
        raise Exception('insufficient storage')

print('Finished First Pass of all of them')
############################## reiterate now


# pyautogui.click('exit.jpg') # need to exit now, not sure how to....


# Root folder to be bptnumber
# sleep(600) # 10 expect a ten minute download time.
# find icon (done).Click
##########


# END Start again with next in set!


###################################################
#################     Apendix     #################
###################################################
# print(soup.prettify())
# driver.find_element_by_id("id0").click() #open Study results
# Ref. Phys <> DO NOT USE - Referrer must be specified

# qsearchstring = \
# driver.find_element_by_name("qsearchstring").send_keys(NextInList)
# qsearchbutton = driver.find_element_by_name("qsearch")
# qsearchstring.clear()
# qsearchstring
# qsearchbutton.click()
# note at this point qsearchstring and qsearchbutton will now be out of date.


# advsearch = driver.find_element_by_id("search")
# search_button_map = driver.find_element_by_name("search_button_map")

# select = driver.find_element_by_name("select")
# search.click()
# search.is_displayed()
# search.submit()
# search.send_keys("openSearch(True);")
# QuickSearch = driver.find_element_by_xpath("//img[@title='Quick Search']")
# search = driver.find_element_by_xpath("//img[@id='search']")
# search = driver.find_element_by_xpath("//img[@name='search']")
# search = driver.find_element_by_xpath("//img[@alt='Search']")
# driver.switch_to.default_content()

# driver.quit()
## AdvSearhcPage = driver.page_source

#   title = re.findall("<TITLE>.*</TITLE>",AdvSearhcPage)
#   pprint(title)

# FAILS!!!
# driver.execute_script("goSearch();") #sumits the search, unfortunatly this then hangs!!!
# driver.execute_script("_search()") unfortunatly this then hangs too!!!
# driver.execute_script("parent.window.close();")
# driver.execute_script("parent.TOOLBAR._submit();")
# driver.execute_script("parent.LIST.goSearch()")
# print("executeing...parent.TOOLBAR.goSearch()")
# driver.execute_script("parent.TOOLBAR.goSearch()")

# td_tags[23][break-word]
# soup = BeautifulSoup(driver.page_source,'lxml')
# print(soup.prettify())

# find_the_first_date = soup.find('td', style = "WORD-WRAP: break-word", align = 'right')
# find_the_dates = soup.find_all('td', style = "WORD-WRAP: break-word", align = 'right')
# pprint(find_the_first_date)
# pprint(find_the_dates)
# cmd = 'WMIC PROCESS get Caption,Commandline,Processid'
# proc = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE)
# for line in proc.stdout:
#    print(line)
