import pyautogui
import requests, re
# import httplib2
# import urllib
# import scrapy
from time import sleep
from datetime import *
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from pprint import pprint
import unittest
import numpy as np
from bs4 import BeautifulSoup
import pyodbc
import pandas as pd
import os
from os import listdir
from os.path import isfile, join
import subprocess
from pywinauto import Desktop,Application
from pathlib import Path

conn = pyodbc.connect('Driver={SQL Server};'
                      r'Server=UHLSQLPRIME01\UHLBRICCSDB;'
                      'Database=i2b2_app03_b1_data;'
                      'Trusted_Connection=yes;')

########################### Build list of (['UhlSystemNumber', 'MRN', 'BptNumber', 'RecruitingSite', 'ct_count','echo_count', 'DICOM_Images_Pseudonymised', 'ct_date_time_start'])
ListToDicom = pd.read_sql_query('SELECT * FROM i2b2_app03_b1_data.dbo.DICOM_List', conn)
#Need to remove from the table above all that have ALREADY BEEN DONE!!!

#df2 = pd.DataFrame(
#   ...:     {
#   ...:         "A": 1.0,
#   ...:         "B": pd.Timestamp("20130102"),
#   ...:         "C": pd.Series(1, index=list(range(4)), dtype="float32"),
#   ...:         "D": np.array([3] * 4, dtype="int32"),
#   ...:         "E": pd.Categorical(["test", "train", "test", "train"]),
#   ...:         "F": "foo",
#   ...:     }
#   ...: )

df = pd.read_csv('C:\\briccs_ct\\results.csv', parse_dates = ['date_time_finished','date_time_opened'])
completed = df[df.date_time_finished.notnull()]
list_minus_completed = pd.merge(ListToDicom, completed,how="left", on=["BptNumber"])

#result = pd.merge(left, right, on="key")

with open("C:\\briccs_ct\\results.csv", "ab") as f:
    np.savetxt(f, (to_log), fmt='%s', delimiter=' ')

print(ListToDicom.columns)
print(ListToDicom.head(1))
print('How may outstanting?')
#list_minus_completed.date_time_finished.isnull()
print(ListToDicom.index.max())

# select the search button
driver = webdriver.Ie()
driver.implicitly_wait(3)
driver.get("https://uvweb.pacs.uhl-tr.nhs.uk/login.jsp")
driver.maximize_window()

# LogMeIn
#  .env

f = open("G:\dicom\.env", "r")
myid = f.readline().splitlines()
mypass = f.readline()
f.close()
myid
mypass

sleep(1)
username = driver.find_element_by_id("userName")
username.send_keys(myid)
sleep(1)
password = driver.find_element_by_id("password")
password.send_keys(mypass)

login = driver.find_element_by_name("login")
login.click()
# end of LogMeIn

PacsWindow = driver.window_handles
print(PacsWindow)

################ start iterations
i = 9
NextInList = ListToDicom.at[i, 'MRN']
# NextInList = 'RWES0112807'   ############################HARD CODED FOR TESTING, REMOVE TO GO LIVE
NextInList_bpt = ListToDicom.at[i, 'BptNumber']

date_to_find = datetime.utcfromtimestamp(
    ListToDicom['ct_date_time_start'].values[i].astype(datetime) / 1_000_000_000).strftime('%m-%d-%Y')
print(NextInList)
print(date_to_find)

# open advanced search and find its handle and switch to window

PacsWindow = driver.window_handles
print(PacsWindow)
driver.current_url
driver.switch_to.window(PacsWindow[0])
driver.current_url
driver.execute_script("openSearch(false);")
sleep(1)
advsearchwindow = list(set(driver.window_handles) - set(PacsWindow))[0]

driver.window_handles
print(advsearchwindow)
print(PacsWindow)
driver.switch_to.window(
    advsearchwindow)  # driver.switch_to.window(PacsWindow) #driver.switch_to.window(advsearchwindow)
driver.switch_to.frame("LIST")

########## Enter details for search and submit
sleep(1)
driver.find_element_by_name("searchPatientId").send_keys(NextInList)
driver.find_element_by_name("searchFirstName").click()  # click away from searchPatientId to enable searchStudyDescr
# driver.execute_script("dialogForm.searchOrderStatus[8].click();") # sets to complete, however we mhy also want read offline...
driver.execute_script("dialogForm.searchImgCnt.click();")  # only studys with images
driver.find_element_by_name("searchStudyDescr").send_keys("CT Cardiac angiogram")

driver.switch_to.window(advsearchwindow)
driver.switch_to.frame("TOOLBAR")
butts = driver.find_elements_by_class_name("search_button")
butts[0].click()
sleep(3)
# driver.window_handles by monitoring the window_handles it seems to close the connection!!!!!!!!!!
print(PacsWindow)
print(driver.window_handles)

print("switching...PacsWindow")
driver.switch_to.window(PacsWindow[0])
driver.switch_to.frame("tableFrame")

soup = BeautifulSoup(driver.page_source, 'lxml')
pprint(soup.find('td', string=re.compile(date_to_find)))
number_of_Dicoms_on_right_Date = len(soup.find_all('td', string=re.compile(date_to_find)))
print('number of Dicoms on date of intrest')
print(number_of_Dicoms_on_right_Date)
# Now to record the number of in range (should only be 1 for yes or 0 for nune however it's possible there are
# more then one.

# if there is only valid one in the list, create new folder and Select
if number_of_Dicoms_on_right_Date == 1:
    path = os.path.join('C:\\briccs_ct\\', NextInList_bpt)
    print("folder to save is")
    print(path)
    if not os.path.exists(path):
        os.makedirs(path)
    listTableForm = driver.find_element_by_name("listTableForm")
    d_found_at = soup.find('td', string=re.compile(date_to_find))
    # images_toProcess is reliant on the img column being two to the right of the date column, if it's not use settings
    images_to_process = int(d_found_at.find_next_sibling().find_next_sibling().string)
    sleep(1)
    listTableForm.click()
to_log = np.array(
    [NextInList + ',' + NextInList_bpt + ',' + str(number_of_Dicoms_on_right_Date) + ',' + str(datetime.now())])
print(to_log)
with open("C:\\briccs_ct\\results.csv", "ab") as f:
    np.savetxt(f, (to_log), fmt='%s', delimiter=' ')

# ----- pyautogui.click() works however the cursor / mouse pointer will need to be moved to position
# ----- pyautogui.rightClick() and the pointer can be just about anywhere, sending F12 would be better
# ----- pyautogui.press('f12') works

cmd = 'WMIC PROCESS get Caption,Commandline,Processid'
proc = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE)
for line in proc.stdout:
    print(line)


########## Export the image
sleep(5)  # page loading
pyautogui.press('f12')      # opens the save images windows, focused on save button by default
pyautogui.typewrite('d')    # set file type to DICOM(*.dcm)
pyautogui.press('tab')      # tab over to folder
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.typewrite(path)   # set to predefined path
pyautogui.press('tab')      # tab apply to images
pyautogui.press('tab')
pyautogui.press('down') # select the non-default save 'Entire Study' (two down from default)
pyautogui.press('down')
pyautogui.press('tab')      # tab to do not create subfolder
pyautogui.press('space')           # click not create subfolder
pyautogui.press('tab')      # select File Name Header and type bpt number in
pyautogui.typewrite(NextInList_bpt)
                            # option - anonymise should always be defult and checked
pyautogui.press('tab')      # tab to the save button
starting_download = datetime.now()
pyautogui.press('space')    # time to save the files!

# take about ten minutes to save the stuff to C drive so wait at least 5 mins before starting to check
sleep(300)

images_to_do = images_to_process
while images_to_do>1:
    sleep(1)
    number_of_dicoms_downloaded = len([f for f in listdir(path) if isfile(join(path, f))])
    images_to_do = images_to_process - number_of_dicoms_downloaded
    message = str(number_of_dicoms_downloaded)+' of '+str(images_to_process)+' downloaded, '+str(images_to_do)+' to do.'
    print(message)
    print('Time now:', str(datetime.now()))
    sleep(5)
finished_downloading = datetime.now()
download_took = finished_downloading - starting_download

print(NextInList_bpt + " finished! Time taken(h:mm:ss.ms):", download_took)




#pyautogui.click('exit.jpg') # need to exit now, not sure how to....


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
