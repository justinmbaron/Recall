import os
#import openpyxl
#from openpyxl import Workbook
import csv
import time
import sys
from datetime import date
from datetime import datetime
import glob
from selenium import webdriver
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from tkinter import *
import pymsgbox

def loginWriteupp():
    #Login to writeUpp
    loginDriver = driver
    loginDriver.get(loginURL)
    driver.maximize_window()
    time.sleep(2)
    userNameField = loginDriver.find_element_by_id('EmailAddress')
    userNameField.send_keys(userName)
    passwordField = driver.find_element_by_id('Password')
    passwordField.send_keys(password)
    time.sleep(1)
    submitButton = driver.find_element_by_xpath('/html/body/div[2]/main/div/div[2]/div/form/div[3]/div/div/button')
    submitButton.click()
    time.sleep(3)
    return

#Code starts here
writeUppURL = 'https://dr-andrew-iles.writeupp.com/'
driverPath = 'C:/Billing/geckodriver'

loginURL = 'https://portal.writeupp.com/login'

userName = 'consultant@drandrewiles.co.uk'
password = 'V0qvlF9KZ$Ur'

version_no = "V1 31 December 2022"

wd = 'C:\\Billing\\Recall'
downloadDirectory = wd

document_url = 'https://dr-andrew-iles.writeupp.com/patientsbydocuments.aspx'
docs_url = 'https://dr-andrew-iles.writeupp.com/admin/documents.aspx'



profile = webdriver.FirefoxProfile()
profile.set_preference('browser.download.folderList', 2)
profile.set_preference('browser.download.manager.showWhenStarting',False)
profile.set_preference('browser.download.dir', downloadDirectory)
profile.set_preference('browser.helperApps.neverAsk.saveToDisk','text/csv')
driver = webdriver.Firefox(executable_path = driverPath,firefox_profile=profile)


def get_documents():
    document_list = []
    driver.get(docs_url)
    all_documents = driver.find_elements(By.CSS_SELECTOR, 'td')
    document_names = all_documents[::3]
    for document_name in document_names:
        print(document_name.text)
        document_list.append(document_name.text)
    return document_list


loginWriteupp()
# driver.get(taqg_url)
# all_documentqs = driver.find_element(By.ID,'ctl00_ctl00_Content_ContentPlaceHolder1_cmbRef').text
# taqg_list = all_taqgs.split('\n')
# for taqg in taqg_list:
#     print(taqg)

all_documents = get_documents()

def fillout_documents(e):
    # update entry box with whatever is clicked
    document_entry.delete(0, END)
    document_entry.insert(0,document_list_box.get(ANCHOR))

def update_document_list(document_names):
    # Clear the list box
    document_list_box.delete(0, END)
    # Add documents to listbox
    for document in document_names:
        document_list_box.insert(END, document)

def clear_window():
    # destroy all widgets from frame
    for widget in root.winfo_children():
       widget.destroy()

def check_documents(e):
    # Check entry vs listbox
    typed = document_entry.get()
    if type == "":
        data = all_documents
    else:
        data = []
        for item in all_documents:
            if typed.lower() in item.lower():
                data.append(item)
    update_document_list(data)


# UI
root = Tk()
root.geometry("500x800")
document_list_lbl = Label(root, text="Select a document")
document_entry = Entry(root,width=40 )
document_list_box = Listbox(root, width=40)
submit_button = Button(root, text="Push Me", command=clear_window)

document_list_lbl.grid(row=1, column=1, sticky="W",pady=10, padx=10)
document_entry.grid(row=1, column=2, sticky="W", pady=10, padx=10)
document_list_box.grid(row=2, column=2, sticky="W", pady=10, padx=10)
submit_button.grid(row=4,column=3,sticky="E", pady=10, padx=10)

# Populate list box
all_documents = get_documents()
update_document_list(all_documents)


document_list_box.bind("<<ListboxSelect>>", fillout_documents)

# Bind keystroke to function
document_entry.bind("<KeyRelease>", check_documents)

root.mainloop()

