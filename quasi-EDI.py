"""
Created on Fri Jan 15 10:43:45 2021

@author: ToToChow
UpDates:
Wed Feb 03, 2021 - Adding in Logs

To Do: 
add unit of measure    
error check on qty to make sure numbers only
add comments as shipto address

    

"""

## Import Packages ##
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
import time
import pandas as pd
import numpy as np
from datetime import datetime, date, timedelta
import os
import re
import shutil


## Define File Path ##
ChromeDriverPATH = 'M:\\TC\\Freshline\\Control Files\\chromedriver.exe'
FreshlinePath = 'M:\\TC\\Freshline\\New Orders\\'
ArchivePath = 'M:\\TC\\Freshline\\Archived Orders\\' 
ErrorPath = 'M:\\TC\\Freshline\\Error Orders\\' 
LogPath = 'M:\\TC\\Freshline\\Orders Log\\Logs.txt'
driver = webdriver.Chrome(ChromeDriverPATH)
driver.get(** Removed for privacy reason **)

FISH1 =** Removed for privacy reason **
TFISH =** Removed for privacy reason **

#!! SELECT SERVER HERE !!#
##########################
##########################
SelectServer = TFISH
##########################
##########################
##########################

##########################
##########################
### Define Sleep Times ###
##########################
#Long Pause
lp= 1.5
#Medium Pause
mp = 1.25
#Short Pause
sp = .75
##########################
##########################


## Functions to be called ##
## Log in ##
def LogIn():
    himitsu = open("M:\\TC\\Freshline\\Control Files\\himitsu.txt", "r")
    UserID = re.search('(?<=L:)\S+',re.findall(r'L:.*', himitsu.read())[0])[0]
    himitsu.seek(0)
    Password = re.search('(?<=P:)\S+',re.findall(r'P:.*', himitsu.read())[0])[0]
    himitsu.close()
    time.sleep(sp)
    ## Try logging in, if log in successful,it will append a "success" message to the log file. "failed" if otherwise ##
    try:
        driver.find_element_by_id("mLogin-inputEl").send_keys(UserID)
        driver.find_element_by_id("mPassword-inputEl").send_keys(Password)
        driver.find_element_by_id("mLoginBtn-btnEl").click()
        time.sleep(mp)
        driver.find_element_by_id("combo-1029-trigger-picker").click()
        driver.find_element_by_id("combo-1029-inputEl").send_keys(SelectServer)
        driver.find_element_by_id("combo-1029-inputEl").send_keys(Keys.TAB)
        driver.find_element_by_id("button-1032-btnInnerEl").click()
        WriteLog("Logged in on: " + SelectServer + ' as: ' + UserID + ' at: ' + str(datetime.now()))
    except:
        WriteLog("Failed to Log on: " + SelectServer + ' as: ' + UserID + ' at: ' + str(datetime.now()) + ' ' + driver.find_element_by_id("toast-1027-innerCt").text)


## function to write log file #
def WriteLog(msg):
    with open(LogPath, "a+") as file_object:
        # Move read cursor to the start of file.
        file_object.seek(0)
        # If file is not empty then append '\n'
        data = file_object.read(100)
        if len(data) > 0 :
            file_object.write("\n")
        # Append text at the end of file
        file_object.write(msg)
        file_object.close()


## Opening Sales Data Entry Link ##
def OpenSales():
    time.sleep(mp)
    driver.get("https://7seas.moniroo.com/V5/sales/create")



## Adding a new item line and enter the itemNum and Qty ##
def EnterOrderItems(vItemNum, vQuantity):
    time.sleep(mp)
    driver.find_element_by_id('btnAdd').click()
    time.sleep(sp)
    driver.find_element_by_id('itemCombo-inputEl').send_keys(vItemNum)
    time.sleep(sp)
    driver.find_element_by_id('itemCombo-inputEl').send_keys(Keys.DOWN + Keys.TAB*4)
    time.sleep(sp)
    driver.find_element_by_id('quantityField-inputEl').send_keys(str(vQuantity) + Keys.ENTER)


## Try to Submitting Orders and Return Errors, if any ##
## If the data entry was successful, then the file will be moved to the archive folder ##
## if not, then it will be moved to the error folder ##

def Submit(file, Cust, CustOrd):
    driver.find_element_by_id('button-1065-btnWrap').click()
    driver.find_element_by_id('button-1006-btnEl').click()   
    time.sleep(lp)
    try:
        WriteLog(CustOrd + ' ' + driver.find_element_by_css_selector("[id$=displayfield-inputEl]").text + '\n')
        driver.find_element_by_id('button-1005-btnIconEl').click()
        shutil.move(file, ArchivePath)

    except:
        driver.find_element_by_id('btnClear').click()
        WriteLog(Cust + '\t' + CustOrd + '\t' + driver.find_element_by_css_selector("[id^=uxNotification-]" + '\n').text)
        driver.find_element_by_xpath('/html/body/div[24]/div[3]/div/div/a[2]/span/span/span[2]')
        shutil.move(file, ErrorPath)
        
#    try:
#         WriteLog('\n\t' + Cust + '\t' + CustOrd + '\t' + driver.find_element_by_css_selector("[id^=uxNotification-]" + '\n').text)
#    except:
#        pass

## Entering Orders ##
def EnterOrderHeader(vSalesType,vCustomer,vShipDate,vCustomerPO):
    OpenSales()
    time.sleep(lp)
    frame = driver.find_element_by_id('box-1036')
    driver.switch_to.frame(frame)
    driver.find_element_by_name('salesTypeCombo-inputEl').click()
    driver.find_element_by_id('salesTypeCombo-inputEl').send_keys(vSalesType)
    time.sleep(mp)
    driver.find_element_by_name('customerCombo-inputEl').click()
    driver.find_element_by_id('customerCombo-inputEl').send_keys(vCustomer)
    time.sleep(mp)
    driver.find_element_by_name('shipDate-inputEl').click()
    driver.find_element_by_id('shipDate-inputEl').send_keys(vShipDate)
    time.sleep(mp)
    driver.find_element_by_name('custPoNum-inputEl').click()
    driver.find_element_by_id('custPoNum-inputEl').send_keys(vCustomerPO)


## Access Freshline Path folder and return the amount of files ##
def ReadLoopFiles(path):
    folders_dir = os.listdir(path)
    files_csv = [path + s for s in folders_dir]
    if len(files_csv) == 0:
        WriteLog('\t **No New Files**\n')
    for f in files_csv:
        data = pd.read_csv(f)
        SalesType = data.iloc[0][0]
        Customer = data.iloc[0][1]
        ShipDate = data.iloc[0][4]
        CustomerPO = data.iloc[0][5]
        if len(data['vCustomer'].unique()) == 1:
            WriteLog('The following order is being entered ' + Customer + ' - ' + CustomerPO + ' :')
            EnterOrderHeader(SalesType, Customer, ShipDate, CustomerPO)
            for i in data.index:
                itemNum = data.iloc[i][2]
                Quantity = data.iloc[i][3]
                EnterOrderItems(itemNum, Quantity)
                WriteLog('Item: ' + itemNum + ' Qty Ordered: ' + str(Quantity))
        else:
            WriteLog('The following order is being entered ' + Customer + ' - ' + CustomerPO + ' : Error' )
            shutil.move(f, ErrorPath)
            continue
        Submit(f, Customer, CustomerPO)



## do these things ##


LogIn()
ReadLoopFiles(FreshlinePath)
driver.close()
