import sys
import webbrowser
from tkinter import *
from tkcalendar import *
import customtkinter
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
import time
import os
import csv
from datetime import datetime
import glob
import pymsgbox
from configparser import ConfigParser
import ast

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

def display_button(button_text):
    submit_button = customtkinter.CTkButton(root,
                                            text=button_text,
                                            border_width=2,
                                            corner_radius=8,
                                            width=120,
                                            height=32,
                                            font=("Arial", 20),
                                            command=main_process)
    submit_button.pack(pady=20, padx=10)

def write_line(row_text):
    writer.writerow(row_text)


def clear_window():
    # destroy all widgets from frame
    for widget in root.winfo_children():
        widget.destroy()

def get_documents():
    # get all the names of document templates
    document_list = []
    driver.get(docs_url)
    all_documents = driver.find_elements(By.CSS_SELECTOR, 'td')
    document_names = all_documents[::3]
    for document_name in document_names:
        # print(document_name.text)
        document_list.append(document_name.text)
    #document_list = ['Doc1', 'Doc2', 'Doc3', 'Doc4', 'Doc5', 'Doc6', 'Doc7', 'Doc8', 'Doc9', 'Doc10', 'Doc11', ]
    return document_list

def get_emails():
    # Get all the names of email templates
    # select test patient
    email_list = []
    driver.get(email_url)
    time.sleep(1)
    all_emails = driver.find_elements(By.CSS_SELECTOR, 'td')
    email_names = all_emails[::4]
    for email_name in email_names:
        # print(email_name.text)
        email_list.append(email_name.text)
    #document_list = ['Doc1', 'Doc2', 'Doc3', 'Doc4', 'Doc5', 'Doc6', 'Doc7', 'Doc8', 'Doc9', 'Doc10', 'Doc11', ]
    return email_list


def get_direct_messages():
    # Get all the names of direct message templates
    direct_message_list =[]
    driver.get(direct_message_url)
    time.sleep(1)
    driver.find_element_by_xpath('/html/body/form/div[5]/div/div/div/p[2]/button').click() # send direct message
    time.sleep(2)
    # if there is more than one name - choose the top one
    try:
        top_name = driver.find_element(By.CSS_SELECTOR,"a[class='list-view__inner list-view__inner']")
        top_name.click() #select first name on list
        time.sleep(2)
    except:
        pass
    #driver.find_element(By.CSS_SELECTOR,"a[class='re-button re-directMessageTemplatesManage']").click() #manage
    driver.find_element(By.PARTIAL_LINK_TEXT,'Templates').click() #Templates
    time.sleep(2)
    # all_direct_messages = driver.find_elements(By.CSS_SELECTOR,"li")
    all_direct_messages = driver.find_elements(By.CSS_SELECTOR, "div[rel='directMessageTemplates']")
    direct_message_list = all_direct_messages[0].text.split('\n')
    return direct_message_list


def get_tags():
    # Get all the tags
    tag_list = []
    driver.get(tag_url)
    all_tags = driver.find_element(By.ID,'ctl00_ctl00_Content_ContentPlaceHolder1_cmbRef').text
    all_tag_list = all_tags.split('\n')
    for tag in all_tag_list:
        tag_list.append(tag)
    #tag_list = ['Tag1', 'Tag2', 'Tag3', 'Tag4', 'Tag5', 'Tag6', 'Tag7', 'Tag8', 'Tag9', 'Tag10']
    return tag_list

def get_cover_emails():
    cover_email_list=[]
    driver.get(email_cover_url)
    time.sleep(2)
    all_cover_emails = driver.find_elements(By.CSS_SELECTOR, 'td')
    email_names = all_cover_emails[::4]
    for email_name in email_names:
        # print(email_name.text)
        cover_email_list.append(email_name.text)
    return cover_email_list

def get_sms():
    sms_list=[]
    driver.get(sms_url)
    time.sleep(2)
    #driver.find_element_by_id('ctl00_ctl00_ctl00_Content_ContentPlaceHolder1_Content_cmbTemplate').click()
    all_sms_templates = driver.find_elements(By.CSS_SELECTOR,'option')
    for sms_template in all_sms_templates:
        sms_template_name = sms_template.text
        if sms_template_name == 'Blank':
            pass
        elif sms_template_name =='Variables':
            break # you have reached the end
        else:
            sms_list.append(sms_template.text)
    return sms_list


def clean_up_files():
    # Delete all the downloaded tagged patient files
    files = glob.glob(wd+'//*.csv')
    for f in files:
        os.remove(f)

def download_tagged_patients():
    driver.get(tag_url)
    tag_selector_field = Select(driver.find_element_by_id('ctl00_ctl00_Content_ContentPlaceHolder1_cmbRef'))
    tag_selector_field.select_by_visible_text(chosen_tag)
    time.sleep(1)
    open_tick = driver.find_element(By.CSS_SELECTOR,"input[name='ctl00$ctl00$Content$ContentPlaceHolder1$cbOpenOnly']")
    open_tick.click()
    driver.find_element_by_id('ctl00_ctl00_Content_ContentPlaceHolder1_btnExportCsv').click() # export button

def download_activity():
    driver.get(activity_url)
    time.sleep(1)
    from_field = driver.find_element_by_id('ctl00_ctl00_Content_ContentPlaceHolder1_dfDateFrom')
    from_field.clear()
    from_field.send_keys(start_date)
    driver.find_element_by_id('ctl00_ctl00_Content_ContentPlaceHolder1_btnFilter').click()
    time.sleep(2)
    driver.find_element_by_id('ctl00_ctl00_Content_ContentPlaceHolder1_btnExportCsv').click()
    time.sleep(2)
    return

def find_patients_to_recall():
    # pymsgbox.alert('Finding Patients - this takes a few seconds', timeout=200)
    os.chdir(wd)
    d = {}
    download_activity()
    with open('Activity by date.csv', mode='r') as f:
        appointments = csv.reader(f)
        for appointment in appointments:
            wuid = appointment[0]
            this_appointment = [[appointment[6], appointment[7], appointment[1]]]
            if wuid in d:
                d[wuid] = d[wuid] + this_appointment
            else:
                d[wuid] = this_appointment
    f.close()

    recall_list = []
    with open(tag_filename, newline='') as tagged_patient_file:
        all_patients = csv.reader(tagged_patient_file)
        next(all_patients)
        for patient in all_patients:
            wuid = patient[0]
            patient_name = patient[1]
            if wuid in d:
                # Appointments in the period but check they weren't all cancelled etc
                possible_appointments = d[wuid]
                keep_flag = True
                for appointment in possible_appointments:
                    if appointment[1] in include_status_list:
                        keep_flag = False

                # print(wuid, keep_flag)
                if keep_flag == True:
                    # no valid appointments so mark them for comms
                    recall_list.append([wuid,patient_name])

            else:
                # no appointments in the time period so mark them for comms
                recall_list.append([wuid,patient_name])
    tagged_patient_file.close()
    return recall_list

def process_direct_message():
    global direct_message_object
    direct_message_names = get_direct_messages()
    direct_message_text = "Select a Direct Message Template"
    direct_message_object = RecallUi(root, direct_message_names) \
        .set_label_text(direct_message_text) \
        .render()
    display_button("Continue to send")

def process_letter_email():
    global doc_object
    doc_list = get_documents()
    doc_text = "Select a Document"
    doc_object = RecallUi(root, doc_list) \
        .set_label_text(doc_text) \
        .render()

    global cover_email_object
    cover_email_list = get_cover_emails()
    cover_email_text = "Select a cover email template"
    cover_email_object = RecallUi(root, cover_email_list) \
        .set_label_text(cover_email_text) \
        .render()
    display_button("Continue to send")


def process_letter_direct():
    global doc_object
    doc_list = get_documents()
    doc_text = "Select a Document"
    doc_object = RecallUi(root, doc_list) \
        .set_label_text(doc_text) \
        .render()

    global direct_message_object
    direct_message_names = get_direct_messages()
    direct_message_text = "Select a Direct Message Template"
    direct_message_object = RecallUi(root, direct_message_names) \
        .set_label_text(direct_message_text) \
        .render()
    display_button("Continue to send")


def process_email():
    global email_object
    email_name_list = get_emails()
    email_message_text = "Select an email template"
    email_object = RecallUi(root, email_name_list) \
        .set_label_text(email_message_text) \
        .render()
    display_button("Continue to send")


def process_sms():
    global sms_object
    sms_name_list = get_sms()
    sms_message_text = 'Select an SMS template'
    sms_object = RecallUi(root, sms_name_list) \
        .set_label_text(sms_message_text) \
        .render()
    display_button("Continue to send")

def find_patient(wuid):
    search_boxs =driver.find_elements(By.CSS_SELECTOR,"input[type='search']")
    # search_box = driver.find_element_by_id('ctl00_ctl00_Content_siteHead_dfSearchWidget')
    search_boxs[0].send_keys(wuid)
    search_boxs[0].send_keys(Keys.RETURN)
    time.sleep(1)
    return


def process_choice(choice):
    print('you chose ', choice)
    print('Date is: '+cal.get_date())
    print(tag_object.entry.get())

    global start_date, chosen_tag, comms_method
    start_date = cal.get_date()
    chosen_tag = tag_object.entry.get()
    comms_method = choice

    # Check they have actually chosen
    if comms_method == 'Make your selection':
        pymsgbox.alert("You didn't select a way to communicate, try again")
        return

    confirmation_text = "You have chosen to send a  "+comms_method+" to patients tagged with "+chosen_tag+" who have not been seen since "+start_date
    response = pymsgbox.confirm(text = confirmation_text, buttons=['OK','Try again'] )
    if response == 'Try again':
        # go back and get another selection
        return
    # carry on anad get the next selections
    clear_window()

    # Opem Run report
    time_stamp  = datetime.now().strftime('%d.%m.%Y %H%M%S')
    report_name = run_report+' ' + time_stamp+'.csv'
    os.chdir(reports_folder)
    global writer,rf
    rf=open(report_name,'a',newline='')
    writer = csv.writer(rf)
    write_line([confirmation_text])
    write_line(report_header)
    os.chdir(wd)


    if comms_method == "Email":
        process_email()
    elif comms_method == "Direct Message":
        process_direct_message()
    elif comms_method == "Letter attached to email":
        process_letter_email()
    elif comms_method == "Letter attached to Direct message":
        process_letter_direct()
    elif comms_method == "SMS":
        process_sms()
    else:
        # Shouldn't really get here
        pymsgbox.alert('This should never happen line 326 - call Justin')
        return

def send_emails(patients):
    chosen_email_template = email_object.entry.get()
    for patient in patients:
        find_patient(patient[0])
        driver.get(send_email_url)
        time.sleep(1)
        driver.find_element_by_xpath('/html/body/form/div[5]/div/div/div/div[3]/div/div[1]/a').click() # open To: panel
        time.sleep(1)
        #driver.find_element_by_xpath('/html/body/form/div[5]/div/div/div/div[3]/div/div[2]/div/div/table[1]/tr[2]/td[4]/input').click() # tick-box
        all_rows = driver.find_elements(By.CSS_SELECTOR,"tr")

        for row in all_rows:
            row_text = row.text
            if 'Patient' in row_text:
                error_flag = False
                print('Found it')
                all_tds = row.find_elements_by_css_selector("*")
                all_tds[3].click()
                break
            else:
                error_flag= True

        if not error_flag: # found a patient email address
            email_template_selector = Select(
                driver.find_element_by_id('ctl00_ctl00_ctl00_Content_ContentPlaceHolder1_Content_cmbTemplate'))
            email_template_selector.select_by_visible_text(chosen_email_template)
            time.sleep(1)
            driver.find_element_by_id('ctl00_ctl00_ctl00_Content_ContentPlaceHolder1_Content_pbSend').click() # send button
            print('Sent!')
            # something here for run report
            report_line = [patient[0], 'Email Sent']
            write_line(report_line)
        elif error_flag: # Didn't find a patient email address
            print('no patient email for ' + patient[0])
            # error stuff goes here
            report_line = [patient[0], "***Errror", "No email address"]
            write_line(report_line)



def send_direct_message(patients):
    first_time_flag = True
    chosen_direct_template = direct_message_object.entry.get()
    for patient in patients:
        find_patient(patient[0])
        driver.get(direct_message_url)
        time.sleep(1)
        driver.find_element_by_xpath('/html/body/form/div[5]/div/div/div/p[2]/button').click()  # send direct message
        time.sleep(1)

        # Choose the patient name from the list of peple to send to
        possible_links = driver.find_elements(By.CSS_SELECTOR, "a[class='list-view__inner list-view__inner']")
        for link in possible_links:
            if link.text == patient[1]:
                link.click()
                break
        time.sleep(2)

        try:
            # Check if the no email/mobile by clicking on back button in pop-up
            #driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[1]/div/div[2]/div/form/div[3]/button[2]').click() #back button
            driver.find_element(By.CSS_SELECTOR,"input[id='email']").click()
            time.sleep(1)
            print('No email or Phone')
            # write something to the run report
            report_line = [patient,'***Error','No email or Phone']
            write_line(report_line)

            #driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[1]/div/div[2]/div/form/div[3]/button[2]').click() #back button
            close_x = driver.find_element(By.CSS_SELECTOR,"button[title='Close']") # close x
            close_x.click()
            time.sleep(1)
        except:
            # all good
            driver.find_element(By.PARTIAL_LINK_TEXT, 'Templates').click()  # Templates
            time.sleep(1)
            driver.find_element_by_link_text(chosen_direct_template).click()
            send_button = driver.find_element(By.CSS_SELECTOR,"button[class='flat-button flat-button--primary flat-button--small']")
            print('Found the Send button')
            print('stop')
            send_button.click() # send button uncomment when live
            report_line = [patient[0],'DM sent']
            write_line(report_line)
            if first_time_flag:
                tick_box = WebDriverWait(driver, 15).until(
                    EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/div[1]/p[2]/label/input")))
                tick_box.click()

                close_button = WebDriverWait(driver, 15).until(
                    EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div[2]/header/button")))
                close_button.click()
                first_time_flag = False
            time.sleep(1)

def send_letter_email(patients):
    error_flag = False
    chosen_letter = doc_object.entry.get()
    chosen_cover_email = cover_email_object.entry.get()
    for patient in patients:
        find_patient(patient[0])
        driver.get(docs_to_send_url)
        time.sleep(1)
        #scroll into view befor clicking
        document_link = driver.find_element_by_link_text(chosen_letter)
        driver.execute_script("arguments[0].scrollIntoView(true);", document_link)
        time.sleep(0.5)
        document_link.click()
        time.sleep(1)
        driver.find_element_by_id('ctl00_ctl00_Content_ContentPlaceHolder1_emailDoc').click() #email icon
        time.sleep(1)
        driver.find_element_by_xpath('/html/body/form/div[5]/div/div/div/div[3]/div/div[1]/a').click()  # open To: panel
        time.sleep(1)
        # driver.find_element_by_xpath('/html/body/form/div[5]/div/div/div/div[3]/div/div[2]/div/div/table[1]/tr[2]/td[4]/input').click() # tick-box
        all_rows = driver.find_elements(By.CSS_SELECTOR, "tr")
        for row in all_rows:
            row_text = row.text
            if 'Patient' in row_text:
                error_flag = False
                print('Found it')
                all_tds = row.find_elements_by_css_selector("*")
                all_tds[3].click()
                break
            else:
                error_flag= True
        if not error_flag: #found email address for patient
            email_template_selector = Select(
                driver.find_element_by_id('ctl00_ctl00_ctl00_Content_ContentPlaceHolder1_Content_cmbTemplate'))
            email_template_selector.select_by_visible_text(chosen_cover_email)
            driver.find_element_by_id('ctl00_ctl00_ctl00_Content_ContentPlaceHolder1_Content_pbSend').click() # send button
            print('Sent!')
            # something here for run report
            report_line = [patient[0],'Document sent']
            write_line(report_line)
        elif error_flag: #no email address for patient
            print('no patient email for ' + patient[0])
            # error stuff goes here
            report_line = [patient[0], "***Errror", "No email address"]
            write_line(report_line)


def send_letter_direct(patients):
    first_time_flag = True
    chosen_letter = doc_object.entry.get()
    chosen_cover_direct_message = direct_message_object.entry.get()
    for patient in patients:
        find_patient(patient[0])
        driver.get(docs_to_send_url)
        time.sleep(1)

        document_link = driver.find_element_by_link_text(chosen_letter)
        driver.execute_script("arguments[0].scrollIntoView(true);", document_link)
        time.sleep(0.5)
        document_link.click()

        time.sleep(1)
        driver.find_element_by_id('ctl00_ctl00_Content_ContentPlaceHolder1_pbSaveAndContinue').click()
        time.sleep(1)
        driver.find_element_by_id('ctl00_ctl00_Content_ContentPlaceHolder1_dmButton').click()
        time.sleep(1)

        # Choose the patient name from the list of people to send to
        possible_links = driver.find_elements(By.CSS_SELECTOR, "a[class='list-view__inner list-view__inner']")
        for link in possible_links:
            if link.text == patient[1]:
                link.click()
                break
        time.sleep(2)

        try:
            # Check if the no email/mobile by clicking on back button in pop-up
            # driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[1]/div/div[2]/div/form/div[3]/button[2]').click() #back button
            driver.find_element(By.CSS_SELECTOR, "input[id='email']").click()
            time.sleep(1)
            print('No email or Phone')
            # write something to the run report
            report_line = [patient[0], '***Error', 'No email or Phone']
            write_line(report_line)

            # driver.find_element_by_xpath('/html/body/div[2]/div[2]/div[1]/div/div[2]/div/form/div[3]/button[2]').click() #back button
            close_x = driver.find_element(By.CSS_SELECTOR, "button[title='Close']")  # close x
            close_x.click()
            time.sleep(1)
        except:
            # all good
            driver.find_element(By.PARTIAL_LINK_TEXT, 'Templates').click()  # Templates
            time.sleep(1)
            driver.find_element_by_link_text(chosen_cover_direct_message).click()
            send_button = driver.find_element(By.CSS_SELECTOR,
                                              "button[class='flat-button flat-button--primary flat-button--small']")
            print('Found the Send button')
            print('stop')
            send_button.click()  # send button uncomment when live
            report_line = [patient[0], 'Document sent by DM']
            write_line(report_line)
            if first_time_flag:
                tick_box = WebDriverWait(driver, 15).until(
                    EC.element_to_be_clickable((By.XPATH, "/html/body/div[3]/div[2]/div[1]/p[2]/label/input")))
                tick_box.click()

                close_button = WebDriverWait(driver, 15).until(
                    EC.element_to_be_clickable((By.XPATH, "/html/body/div[3]/div[2]/header/button")))
                close_button.click()
                first_time_flag = False

            time.sleep(1)


def send_sms(patients):
    chosen_sms_template = sms_object.entry.get()
    for patient in patients:
        find_patient(patient[0])
        driver.get(send_sms_url)
        time.sleep(1)
        driver.find_element_by_xpath('/html/body/form/div[5]/div/div/div/div[3]/div/div[1]/a').click() # open To: panel
        time.sleep(1)
        all_rows = driver.find_elements(By.CSS_SELECTOR, "tr")

        for row in all_rows:
            row_text = row.text
            if 'Patient' in row_text:
                error_flag = False
                print('Found it')
                all_tds = row.find_elements_by_css_selector("*")
                all_tds[4].click()
                break
            else:
                error_flag = True

        if not error_flag:  # found a patient email address
            sms_template_selector = Select(driver.find_element_by_id('ctl00_ctl00_ctl00_Content_ContentPlaceHolder1_Content_cmbTemplate'))
            sms_template_selector.select_by_visible_text(chosen_sms_template)
            time.sleep(1)
            driver.find_element_by_id(
                'ctl00_ctl00_ctl00_Content_ContentPlaceHolder1_Content_pbSend').click()  # send button
            print('Sent!')
            # something here for run report
            report_line = [patient[0], 'SMS Sent']
            write_line(report_line)

        elif error_flag:  # Didn't find a patient mobile phone
            print('no mobile phone for ' + patient[0])
            # error stuff goes here
            report_line = [patient[0], "***Errror", "No mobile phone"]
            write_line(report_line)

def main_process():
    # Find all the patients with the selected tags
    root.withdraw()
    download_tagged_patients()
    list_of_patients = find_patients_to_recall()
    if len(list_of_patients) == 0:
        print('No Patients to contact')
        pymsgbox.alert('No Patients to contact')
        return


    # Now do whatever has been requested to each of the patients
    if comms_method == 'Email':
        send_emails(list_of_patients)
    elif comms_method == 'Direct Message':
        send_direct_message(list_of_patients)
    elif comms_method == 'Letter attached to email':
        send_letter_email(list_of_patients)
    elif comms_method == 'Letter attached to Direct message':
        send_letter_direct(list_of_patients)
    elif comms_method == 'SMS':
        send_sms(list_of_patients)
    else:
        pymsgbox.alert('This should never happen line 301 - call Justin')
    pymsgbox.alert('All done!')
    rf.close()
    driver.quit()
    sys.exit()

class RecallUi:

    def __init__(self, master, selection_list):

        self.label_text = "List of Options"
        self.master = master
        self.selection_list = selection_list

    def render(self):
        self.set_widgets()
        self.place_widgets()
        self.update_display_list(self.selection_list)
        self.bind_inputs()
        return self

    def fillout_entry_box(self, e):
        # update entry box with whatever is clicked
        self.entry.delete(0, END)
        self.entry.insert(0, self.list_box.get(ANCHOR))

    def check_list(self, e):
        # Check entry vs listbox
        typed = self.entry.get()
        if typed == "":
            data = self.selection_list
        else:
            data = []
            for item in self.selection_list:
                if typed.lower() in item.lower():
                    data.append(item)
        self.update_display_list(data)

    def update_display_list(self, list_content):
        # Clear the list box
        self.list_box.delete(0, END)
        # Add documents to listbox
        for list_item in list_content:
            self.list_box.insert(END, list_item)
        return self

    def get_label_text(self):
        # getter
        return self.label_text

    def set_label_text(self, label_text):
        # setter
        self.label_text = label_text
        return self

    def set_widgets(self):
        self.list_lbl = customtkinter.CTkLabel(
            self.master,
            text=self.get_label_text(),
            font=("Arial", 20),
            fg_color=("#2A719D"),
            text_color=("white"),
            corner_radius=8
        )

        self.entry = customtkinter.CTkEntry(
            self.master,
            width=350,
            corner_radius=8,
            font=("Arial", 20)
        )

        self.list_box = Listbox(self.master, width=40)

    def place_widgets(self):
        self.list_lbl.pack(pady=10, padx=10)
        self.entry.pack(pady=10, padx=10)
        self.list_box.pack(pady=10, padx=10)


    def bind_inputs(self):
        self.list_box.bind("<<ListboxSelect>>", self.fillout_entry_box)
        # Bind keystroke to function
        self.entry.bind("<KeyRelease>", self.check_list)


def get_first_choice():
    tag_list = get_tags()
    tag_text = "Select a Tag"
    global tag_object
    tag_object = RecallUi(root, tag_list) \
        .set_label_text(tag_text) \
        .render()

    # Get date
    date_select_lbl = customtkinter.CTkLabel(
        root,
        text="What is the earliest appointment date?",
        font=("Arial", 20),
        fg_color=("#2A719D"),
        text_color=("white"),
        corner_radius=8
    )
    date_select_lbl.pack(pady=10)
    global cal
    cal = Calendar(root, selectmode="day", year=t_year, month=t_month, day=t_day, date_pattern="dd/mm/yyyy")
    cal.pack(pady=20)

    # Decide what is required
    title_lbl = customtkinter.CTkLabel(
        root,
        text="How would you like to communicate to tagged patients?",
        font=("Arial", 20),
        fg_color=("#2A719D"),
        text_color=("white"),
        corner_radius=8
    )
    title_lbl.pack()

    choose_list = customtkinter.CTkOptionMenu(root,
                                              values=selections,
                                              corner_radius=8,
                                              #=("Arial", 14),
                                              dropdown_font=("Arial", 14),
                                              command=process_choice)
    choose_list.pack(padx=20)

    version_label = Label(root, text=version_no)
    version_label.pack(pady=60, padx=10, anchor="e")

    root.mainloop()


# Code starts here
version_no = "V1.1 PC 01/02/23"
wd = 'C:\\Billing\\Recall'
downloadDirectory = wd
config_file = 'recall.ini'
os.chdir(wd)
config = ConfigParser()
config.read(config_file)

writeUppURL = config.get('url', 'ini_writeupp_url')
driverPath = config.get('files', 'ini_driver_path')
loginURL = config.get('login', 'ini_login_URL')
userName = config.get('login', 'ini_user')
password = config.get('login', 'ini_password')

reports_folder = config.get('files', 'ini_reports_folder')
tag_filename = config.get('files', 'ini_tag_report')

document_url = config.get('url', 'ini_document_url')
docs_url = config.get('url', 'ini_docs_url')
docs_to_send_url = config.get('url', 'ini_docs_to_send_url')
tag_url = config.get('url', 'ini_tag_url')
email_url = config.get('url', 'ini_email_url')
sms_url = config.get('url', 'ini_sms_url')
appointments_url = config.get('url', 'ini_appointments_url')
direct_message_url = config.get('url', 'ini_direct_message_url')
email_cover_url = config.get('url', 'ini_email_cover_url')
activity_url = config.get('url', 'ini_activity_url')
send_email_url = config.get('url', 'ini_send_email_url')
send_sms_url = config.get('url', 'ini_send_sms_url')


include_status_list = ast.literal_eval(config.get('writeupp','ini_include_status_list'))

test_patient_wuid = config.get('writeupp', 'ini_test_patient_wuid')
gui_size = config.get('writeupp', 'ini_gui_size')
gui_colour = config.get('writeupp', 'ini_gui_colour')

selections = ['Make your selection','Email', 'Direct Message','Letter attached to email', 'Letter attached to Direct message','SMS']
error_text='A WriteUpp DM uses an access code'
run_report ='Tag Recall Run Report '
report_header = ['WUID','Action','Comments']

today_obj = datetime.today()
today_string = datetime.strftime(today_obj,'%d%b%Y')
t_day = int(datetime.strftime(today_obj, '%d'))
t_month = int(datetime.strftime(today_obj, '%m'))
t_year = int(datetime.strftime(today_obj, '%Y'))

firefox_options = Options()
firefox_options.binary_location = r"C:\Program Files\Mozilla Firefox\firefox.exe"
firefox_options.add_argument("--disable-infobars")
firefox_options.add_argument("--disable-extensions")
firefox_options.add_argument("--disable-popup-blocking")

profile = webdriver.FirefoxProfile()
profile.set_preference('browser.download.folderList', 2)
profile.set_preference('browser.download.manager.showWhenStarting', False)
profile.set_preference('browser.download.dir', downloadDirectory)
profile.set_preference('browser.helperApps.neverAsk.saveToDisk', 'text/csv')
profile.set_preference('browser.download.alwaysOpenPanel', False)

driver = webdriver.Firefox(executable_path = driverPath,firefox_profile=profile, options=firefox_options)

# UI
# customtkinter.set_appearance_mode("dark")
# customtkinter.set_default_color_theme("blue")

root = customtkinter.CTk()
root.geometry(gui_size)
root.configure(fg_color=gui_colour)

clean_up_files()
loginWriteupp()
time.sleep(1)
# select test patient
patient_wuid = test_patient_wuid
search_box = driver.find_element_by_id('ctl00_ctl00_Content_siteHead_dfSearchWidget')
search_box.send_keys(patient_wuid)
time.sleep(1)
search_box.send_keys(Keys.RETURN)
time.sleep(3)
get_first_choice()


submit_button = Button(root, text="Push Me", command= main_process)
submit_button.pack(pady=20, padx=10)



