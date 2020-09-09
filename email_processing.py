# This program will extract information from emails within Microsoft Outlook. The program can then aggregate data into 1 file and then send that data to the necessary party.
# Some elements of this program have been removed. Certain URL paths are missing. Specific data analysis has also been removed.
#!/usr/bin/env python
# coding: utf-8

from openpyxl import load_workbook
import os
import pandas as pd
from selenium import webdriver
import time
import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
root_folder = outlook.Folders.Item("Folder")
email = win32com.client.Dispatch("Outlook.Application")
xlapp = win32com.client.DispatchEx("Excel.Application")

mail = email.CreateItem(0)
mail.To = ''
mail.SentOnBehalfOfName = ''
mail.Subject = 'Daily Alert Report'


## This is the outlook folder with the Alerts - if there are NO Alerts - this program will terminate
alerts = root_folder.Folders['Alerts']
if len(alerts.Items) == 0:
    mail.Body = 'No Alerts to Submit.'
    mail.Send()
# else:

excel_alerts = excel_path
wb = xlapp.workbooks.open(excel_alerts)

# Optional, e.g. if you want to debug
xlapp.Visible = False

# Refresh all data connections.
wb.RefreshAll()
time.sleep(5)
wb.Save()

## fn clean_data

# Quit
xlapp.Quit()


driver = webdriver.Chrome()
## log in to SFTP Website
driver.get(loginpath)
driver.find_element_by_id('username').send_keys('user')
driver.find_element_by_id ('password').send_keys('pw')
driver.find_element_by_name('Submit').click()
driver.find_element_by_name('files[]').send_keys(file)
time.sleep(5)
driver.find_element_by_xpath("//button[@type='submit']").click()
time.sleep(5)
driver.close()

mail.Body = str(len(ed_alerts.Items)) + 'Alerts Submitted.'
mail.Send()


processed_alerts = root_folder.Folders['Alerts - Processed']

while len(alerts.Items) > 0:
    for message in alerts.Items:
        
        message.Move(processed_alerts)

