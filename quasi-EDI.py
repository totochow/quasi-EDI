"""
Created on Fri Jan 15 10:43:45 2021

@author: ToToChow
UpDates:
Wed Feb 03, 2021    - Added Logging files
Mon Feb 08, 2021    - added ability to add shipto address to overall comments
                    


To Do: 
add unit of measure    
error check on qty to make sure numbers only
add comments as shipto address

    

"""

## Import Packages ##
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
#from selenium.webdriver.support.select import Select

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


import time
import pandas as pd
#import numpy as np
from datetime import datetime, date#, timedelta
import os
#import re
import shutil
import sys

import my_config as my_c

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

import imaplib
import requests


username = my_c.usernameo365
password = my_c.passwordo365
mail_from = my_c.usernameo365
mail_to = my_c.spam_to_group

FL_login = my_c.FL_login
FL_password = my_c.FL_password

## Define File Path ##

ChromeDriverPATH = my_c.ChromeDriverPATH
FreshlinePath = my_c.FreshlinePath
PendingPath = my_c.PendingPath
ArchivePath = my_c.ArchivePath
#ErrorPath = my_c.ErrorPath
LogPath = my_c.LogPath
SalesEntryPath = my_c.SalesEntryPath
Site = my_c.Site


driver = webdriver.Chrome(executable_path=ChromeDriverPATH)#, options=options)


FISH1 = my_c.FISH1
TFISH = my_c.TFISH

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
lp= 5	
#Medium Pause
mp = 3
#Short Pause
sp = 2
##########################
##########################
#### Batching Amount #####
batch = 25
##########################
    
    
    

## Functions to be called ##
## Log in ##
def LogIn():
    open(LogPath, "w")
    
    UserID = my_c.L
    Password = my_c.P

    ## Try logging in, if log in successful,it will append a "success" message to the log file. "failed" if otherwise ##
    try:
        time.sleep(lp)
        driver.find_element_by_id("mLogin-inputEl").send_keys(UserID)
        driver.find_element_by_id("mPassword-inputEl").send_keys(Password)
        driver.find_element_by_id("mLoginBtn-btnEl").click()
        time.sleep(mp)
        driver.find_element_by_id("combo-1029-trigger-picker").click()
        driver.find_element_by_id("combo-1029-inputEl").send_keys(SelectServer)
        driver.find_element_by_id("combo-1029-inputEl").send_keys(Keys.TAB)
        driver.find_element_by_id("button-1032-btnInnerEl").click()
        WriteLog("\n\n\nLogged in on: " + SelectServer + ' as: ' + UserID + ' at: ' + str(datetime.now()))
    except:
        Err = "Failed to Log on: " + SelectServer + ' as: ' + UserID + ' at: ' + str(datetime.now()) + ' ' + driver.find_element_by_id("toast-1027-innerCt").text
        WriteLog(Err)
        Email(Err)

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


## Email Function
def Email(x):
    mail_subject = 'EDI Error'
    mail_body = 'The following error message was returned: \n\n\n' + x  + '\n\n\n' 
    mimemsg = MIMEMultipart()
    mimemsg['From']=mail_from
    mimemsg['To']=mail_to
    mimemsg['Subject']=mail_subject
    mimemsg.attach(MIMEText(mail_body, 'plain'))
    mimemsg['X-Priority'] = '2' #.Importance = 2
    connection = smtplib.SMTP(host='smtp.office365.com', port=587)
    connection.starttls()
    connection.login(username,password)
    connection.send_message(mimemsg)
    connection.quit()

def finance_email(excel_file, file_name):
    d = date.today().strftime("%Y-%m-%d")
    mail_subject = "FishDelish Orders Download " + d
    mail_body = "Please do not reply, this is an automated email. \nSee attached Excel File\n" +  "FishDelish Orders Download on " + d
    # Create a multipart message and set headers
    mimemsg = MIMEMultipart()
    mimemsg['From']=mail_from
    mimemsg['To']=mail_to
    mimemsg['Subject']=mail_subject
    mimemsg.attach(MIMEText(mail_body, 'plain'))
    part = MIMEBase('application', "octet-stream")
    part.set_payload(open(excel_file, "rb").read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment', filename=file_name)
    #mimemsg.attach(part)
    mimemsg.attach(part)

    connection = smtplib.SMTP(host='smtp.office365.com', port=587)
    connection.starttls()
    connection.login(username,password)
    connection.send_message(mimemsg)
    connection.quit()

## Opening Sales Data Entry Link ##
def OpenSales():
    time.sleep(mp)
    driver.get(SalesEntryPath)


## Adding a new item line and enter the itemNum and Qty and Price ##
def EnterOrderItems(vItemNum, vQuantity, vPrice):
    time.sleep(mp) #pause
    #Click Add button
    driver.find_element_by_id('btnAdd').click()
    time.sleep(sp)
    #Enter Item Number
    driver.find_element_by_id('itemCombo-inputEl').send_keys(vItemNum + Keys.SPACE)
    time.sleep(sp)
    #Tabbing tru to Qty field
    driver.find_element_by_id('itemCombo-inputEl').send_keys(Keys.DOWN + Keys.TAB*4)
    time.sleep(sp)
    #Entering Qty
    driver.find_element_by_id('quantityField-inputEl').send_keys(str(vQuantity) + Keys.TAB)
    time.sleep(sp)
    #Entering Price
    driver.find_element_by_xpath('/html/body/div[2]/div/div[3]/div[3]/div/div[6]/table/tbody/tr/td[2]/table/tbody/tr/td[1]/input').send_keys(Keys.BACKSPACE*20)
    time.sleep(sp)
    driver.find_element_by_xpath('/html/body/div[2]/div/div[3]/div[3]/div/div[6]/table/tbody/tr/td[2]/table/tbody/tr/td[1]/input').send_keys(str(vPrice))
    time.sleep(sp)
    driver.find_element_by_xpath('/html/body/div[2]/div/div[3]/div[3]/div/div[6]/table/tbody/tr/td[2]/table/tbody/tr/td[1]/input').send_keys(Keys.ENTER)                                                           
    time.sleep(sp)
  
    #If all successful, log as success
    if len(vItemNum) >= 16:
        Tab = '\t'
    else:
        Tab = '\t\t'
    WriteLog(vItemNum + Tab + str(vQuantity) +'\t$ '+ str(vPrice) + '\t Entered!')


## Try to Submitting Orders and Return Errors, if any ##
## If the data entry was successful, then the file will be moved to the archive folder ##
## if not, then it will be moved to the error folder ##

def Submit(file, Cust, CustOrd):
    driver.find_element_by_id('button-1065-btnWrap').click()

    driver.find_element_by_id('button-1006-btnEl').click()   

    time.sleep(mp)
    try:
        WriteLog(str(CustOrd) + ' ' + driver.find_element_by_css_selector("[id$=displayfield-inputEl]").text + '\n')
        driver.find_element_by_id('button-1005-btnIconEl').click()


    except:
        driver.find_element_by_id('btnClear').click()
        ErrorMsg = Cust + '\t' + str(CustOrd) + '\t *****' + driver.find_element_by_css_selector("[id^=uxNotification-]").text + '*****\n'+ '\n'
        
        WriteLog(ErrorMsg)
        Email(ErrorMsg)
        
#    try:
#         WriteLog('\n\t' + Cust + '\t' + CustOrd + '\t' + driver.find_element_by_css_selector("[id^=uxNotification-]" + '\n').text)
#    except:
#        pass

## Entering Orders Header##
def EnterOrderHeader(vSalesType, vCustomer, vDZone, vCustomerPO, vShipDate, OverallComments):
    #Call OpenSales to refresh order entry page
    OpenSales()
    time.sleep(lp) #pause
    #tell ChromeDriver to find and focus to the right frame with box1036 (order entry frame)
    frame = driver.find_element_by_id('box-1036')
    driver.switch_to.frame(frame)
    #Data entry "Sales Type"
    driver.find_element_by_name('salesTypeCombo-inputEl').click()
    driver.find_element_by_id('salesTypeCombo-inputEl').send_keys(vSalesType)
    time.sleep(mp)
    #Data entry "Customer"
    driver.find_element_by_name('customerCombo-inputEl').click()
    driver.find_element_by_id('customerCombo-inputEl').send_keys(vCustomer)
    time.sleep(mp)
    #Entering Shipping Methods
    driver.find_element_by_name('shippingMethodCombo-inputEl').click()
    time.sleep(mp)
    driver.find_element_by_name('shippingMethodCombo-inputEl').clear()
    driver.find_element_by_id('shippingMethodCombo-inputEl').send_keys(vDZone)
    #Entering Site
    driver.find_element_by_name('siteCombo-inputEl').click()
    time.sleep(mp)
    driver.find_element_by_name('siteCombo-inputEl').clear()
    driver.find_element_by_id('siteCombo-inputEl').send_keys(Site)    
    time.sleep(mp)
    #Entering Customer PO Number
    driver.find_element_by_name('custPoNum-inputEl').click()
    driver.find_element_by_id('custPoNum-inputEl').send_keys(str(vCustomerPO))
    time.sleep(mp)
    #Entering the requested Ship Date
    driver.find_element_by_name('shipDate-inputEl').click()
    driver.find_element_by_id('shipDate-inputEl').send_keys(vShipDate)
    time.sleep(mp)
    #Entering overall comments
    driver.find_element_by_name('textarea-1016-inputEl').click()
    driver.find_element_by_id('textarea-1016-inputEl').send_keys(OverallComments)

## Access Freshline Path folder and return the amount of files ##
def ReadLoopFiles(files_csv):
    # If there is no files, log 'no new files'
    if len(files_csv) == 0:
        WriteLog('\t **No New Files**\n')
    # Read files one by one
    for f in files_csv:
        #logging file name and invoice numbers within the file
        WriteLog('** Reading file ' + f + ' **')
        data = pd.read_excel(f)
        WriteLog('The following orders are in this file: \n' + ', '.join(str(e) for e in data['Invoice Number'].unique()))
        
        #looping thru each invoice number
        for inv in data['Invoice Number'].unique():
            df = data[data['Invoice Number'] == inv].reset_index()  
            #First lets define all the HEADING variables
            SalesType = 'ORD'
            Customer = my_c.FLCustomer
            #'Delivery Zone Name'
            DZone = df['Delivery Zone Name'][0]
            #'Invoice Number' = inv
            #'Customer Name'
            OrderCustomer = df['Customer Name'][0]
            #'Full Address'
            Address = df['Full Address'][0]
            'Delivery Type'
            #'Customer Phone'
            Phone = df['Customer Phone'][0]
            #'Order Notes'
            #'Customer Email'
            Email = df['Customer Email'][0]
            #'Delivery Date'
            ShipDate = df['Delivery Date'][0]
            #'Order Notes'
            ONotes = '' if pd.isna(df['Order Notes'][0]) else ('*** Special Instruction: ' + str(df['Order Notes'][0]) + '*** \n')
            
            #Constructing updates logs ro text file
            OverallComments = ONotes + 'Order Number: ' +str(inv)+ '\nShip To: ' + str(OrderCustomer) +', \n'+ str(Address) +'\n' + 'Phone: ' + str(Phone) +'\nEmail: ' + str(Email)
            WriteLog(OverallComments)
            WriteLog('Ship Date: \t'+ str(ShipDate))
            
            #Enter Header Information for the order in the for loop
            EnterOrderHeader(SalesType,Customer,DZone,inv,ShipDate,OverallComments)
            
            #Looping through each item number:
            for itm, line_num in zip(df['Inventory Code'].tolist(), range(len(df['Inventory Code'].tolist()))):
                try: 
                    EnterOrderItems(itm, df['Quantity'].iloc[line_num], df['Price'].iloc[line_num] )
                except:
                    WriteLog(itm + '\t *** Failed ***' )
                    Email(OrderCustomer + '\t' + inv + '\t' + itm + '\t *** Failed ***' )
                    #shutil.move(f, ErrorPath)
                    #print('moved')
                    continue
            time.sleep(lp)
            #pressing submit, if error, use the variables to log errors
            Submit(f, OrderCustomer, inv)
        shutil.move(f, ArchivePath)

def BatchingEntry(path02):
    k = min(len(os.listdir(path02)),batch) #determine if what is the minimum batch size for reading files
    folders_dir = os.listdir(path02)[:k] #list out all files within the folder upto k files
    WriteLog('Processing the following files\n' + '\t'.join(str(e) + '\n' for e in folders_dir))
    files_csv = [path02 + s for s in folders_dir]
    for file in files_csv:
        shutil.move(file, PendingPath)
    new_files_csv =  [PendingPath + s for s in folders_dir]
    return new_files_csv

def FL_Email_SO():
    driver.get(my_c.FL_IO)
    time.sleep(mp)
    driver.find_element_by_xpath("/html/body/div/div/div[1]/div/div/div[1]/div/div/input").send_keys(FL_login)
    #Enter Password
    driver.find_element_by_xpath("/html/body/div/div/div[1]/div/div/div[2]/div/div/input").send_keys(FL_password)
    #Click Login Button
    driver.find_element_by_xpath("/html/body/div/div/div[1]/div/div/div[3]/button/span[1]").click()
    time.sleep(lp)
    driver.get(my_c.FL_Export)
    time.sleep(mp)
    driver.find_element_by_xpath('/html/body/div/main/header/div/div/a[1]').click() #/html/body/div[1]/main/header/div[2]/div/a[1]
    driver.find_element_by_xpath('/html/body/div[1]/main/header/div/div/button[2]').click() # /html/body/div[1]/main/header/div[2]/div/button[2]
    time.sleep(mp)
    driver.find_element_by_xpath('/html/body/div[5]/div[3]/div/div[3]/button[2]').click()
    time.sleep(30) #wait for the email to apperat in CST's mail box
    conn = imaplib.IMAP4_SSL("outlook.office365.com")
    conn.login(username,password)
    conn.select("Freshline")
    resp, items = conn.uid("search",None, 'All')
    emailid = items[0].split()
    emailid
    resp, data = conn.uid("fetch",emailid[0], "(RFC822)")
    email_body = data[0][1].decode('utf-8')
    csv = 'http://url' + str(email_body).replace('=\r\n', '').replace('=\\r\\n', '').replace('=3D','=').partition('http://url')[2].partition('"')[0] 
    #URL of the image to be downloaded is defined as image_url
    r = requests.get(csv) # create HTTP response object
    file_name = 'M:\\TC\\Freshline\\New Orders\\' + str(my_c.tomo) + '.xlsx'
    #Organize mail box by moving the now read email to archive box (copy and delete)
    if r:
        with open(file_name, "wb") as f:
            f.write(r.content)
           
    result = conn.uid('COPY', emailid[0], 'Archive')
    if result[0] == 'OK':
        mov, data = conn.uid('STORE', emailid[0] , '+FLAGS', '(\Deleted)')
    conn.expunge()
    finance_email(file_name, str(my_c.tomo) + '.xlsx') #email finance with the excel file attached
    conn.close()
    conn.logout()
    #driver.quit()
    #print('close browser')


## do these things ##
FL_Email_SO()
driver.get(my_c.Moniroo)
time.sleep(lp)
try:
    element = WebDriverWait(driver, 0).until(EC.presence_of_element_located((By.ID, "mLogin-inputEl")))
except:
    WriteLog('\t **Log in Timed Out**\n')
    driver.quit()
LogIn()
#new_files_csv = []
#Batching12(FreshlinePath)

ReadLoopFiles(BatchingEntry(FreshlinePath))


driver.quit()
sys.exit(0)
