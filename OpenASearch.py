import win32com.client as win32
import requests, re
#import httplib2
#import urllib
#import scrapy
from time import sleep
from datetime import datetime
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
conn = pyodbc.connect('Driver={SQL Server};'
                      r'Server=UHLSQLPRIME01\UHLBRICCSDB;'
                      'Database=i2b2_app03_b1_data;'
                      'Trusted_Connection=yes;')

########################### Build list of (['UhlSystemNumber', 'MRN', 'BptNumber', 'RecruitingSite', 'ct_count','echo_count', 'DICOM_Images_Pseudonymised', 'ct_date_time_start'])
ListToDicom = pd.read_sql_query('SELECT * FROM i2b2_app03_b1_data.dbo.DICOM_List',conn)
print(ListToDicom.columns)
print(ListToDicom.head(1))
print('How may outstanting?')
print(ListToDicom.index.max())


#select the search button
driver = webdriver.Ie()
driver.implicitly_wait(3)
driver.maximize_window()
driver.get("https://uvweb.pacs.uhl-tr.nhs.uk/login.jsp")

#LogMeIn
#  .env

f = open("G:\dicom\.env", "r")
myid =f.readline().splitlines()
mypass =f.readline()
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
#end of LogMeIn

PacsWindow = driver.window_handles
print(PacsWindow)

################ start iterations
i=4 #needs to be set to 0 and recover from a zero images error when going live
NextInList = ListToDicom.at[i,'MRN']
NextInList = 'RWES3137509'
date_to_find = datetime.utcfromtimestamp(ListToDicom['ct_date_time_start'].values[i].astype(datetime)/1_000_000_000).strftime('%m-%d-%Y')
print(NextInList)
print(date_to_find)

#open advanced search and find its handle and switch to window
driver.execute_script("openSearch(false);")
sleep(1)
advsearchwindow = list(set(driver.window_handles) - set(PacsWindow))[0]

driver.window_handles
print(advsearchwindow)
print(PacsWindow)
driver.switch_to.window(advsearchwindow) #driver.switch_to.window(PacsWindow) #driver.switch_to.window(advsearchwindow)
driver.switch_to.frame("LIST")

########## Enter details for search and submit
sleep(1)
driver.find_element_by_name("searchPatientId").send_keys(NextInList)
driver.find_element_by_name("searchFirstName").click() # click away from searchPatientId to enable searchStudyDescr
#driver.execute_script("dialogForm.searchOrderStatus[8].click();") # sets to complete, however we mhy also want read offline...
driver.execute_script("dialogForm.searchImgCnt.click();") #only studys with images
driver.find_element_by_name("searchStudyDescr").send_keys("CT Cardiac angiogram")

driver.switch_to.window(advsearchwindow)
driver.switch_to.frame("TOOLBAR")
butts = driver.find_elements_by_class_name("search_button")
print("executeing...the submit...")
print(butts)
butts[0].click()
print("still not locked...")
sleep(3)
#driver.window_handles by monitoring the window_handles it seems to close the connection!!!!!!!!!!
print(PacsWindow)
print(driver.window_handles)


print("switching...PacsWindow")
driver.switch_to.window(PacsWindow[0])
driver.switch_to.frame("tableFrame")
driver.find_element_by_name("listTableForm").click()

soup = BeautifulSoup(driver.page_source,'lxml')
pprint(soup.find('td', string = re.compile(date_to_find)))
number_of_Dicoms_on_right_Date = len(soup.find_all('td', string = re.compile(date_to_find)))
print('number of Dicoms on date of intrest')
print(number_of_Dicoms_on_right_Date)
# Now to record the number of in range (should only be 1 for yes or 0 for nune however it's possible there are
# more then one.

# Select the first in the list if there is only one
if number_of_Dicoms_on_right_Date==1:
    driver.find_element_by_name("listTableForm").click()




########## Export the image
# sleep(600) # 10 expect a ten minute download time.
#select using win32 the other program window
#left click
#entire study
#file type DICOM(*.dcm)
#Root folder to be bptnumber
#option - Anonymise checked
##########



###################################################
#################     Apendix     #################
###################################################
#print(soup.prettify())
#driver.find_element_by_id("id0").click() #open Study results
# Ref. Phys <> DO NOT USE - Referrer must be specified

#qsearchstring = \
#driver.find_element_by_name("qsearchstring").send_keys(NextInList)
#qsearchbutton = driver.find_element_by_name("qsearch")
#qsearchstring.clear()
#qsearchstring
#qsearchbutton.click()
#note at this point qsearchstring and qsearchbutton will now be out of date.


#advsearch = driver.find_element_by_id("search")
#search_button_map = driver.find_element_by_name("search_button_map")

#select = driver.find_element_by_name("select")
#search.click()
#search.is_displayed()
#search.submit()
#search.send_keys("openSearch(True);")
#QuickSearch = driver.find_element_by_xpath("//img[@title='Quick Search']")
#search = driver.find_element_by_xpath("//img[@id='search']")
#search = driver.find_element_by_xpath("//img[@name='search']")
#search = driver.find_element_by_xpath("//img[@alt='Search']")
#driver.switch_to.default_content()

#driver.quit()
## AdvSearhcPage = driver.page_source

#   title = re.findall("<TITLE>.*</TITLE>",AdvSearhcPage)
#   pprint(title)

#FAILS!!!
#driver.execute_script("goSearch();") #sumits the search, unfortunatly this then hangs!!!
#driver.execute_script("_search()") unfortunatly this then hangs too!!!
#driver.execute_script("parent.window.close();")
#driver.execute_script("parent.TOOLBAR._submit();")
#driver.execute_script("parent.LIST.goSearch()")
#print("executeing...parent.TOOLBAR.goSearch()")
#driver.execute_script("parent.TOOLBAR.goSearch()")

#td_tags[23][break-word]
#soup = BeautifulSoup(driver.page_source,'lxml')
#print(soup.prettify())

#find_the_first_date = soup.find('td', style = "WORD-WRAP: break-word", align = 'right')
#find_the_dates = soup.find_all('td', style = "WORD-WRAP: break-word", align = 'right')
#pprint(find_the_first_date)
#pprint(find_the_dates)
