from selenium.webdriver.firefox.options import Options as FirefoxOptions
import re
import json
from mysql.connector import Error
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
import logging
import os
import configparser
import traceback
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.common import exceptions
import shutil
from datetime import datetime
from datetime import timedelta
from dateutil import tz
import time
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
from openpyxl import Workbook
import openpyxl
import mysql.connector
from dotenv import load_dotenv
load_dotenv()

for handler in logging.root.handlers[:]:
    logging.root.removeHandler(handler)

def data_recon(file_data,current_date):
    def return_value(data):
        return data if data else 0
    logging.info("INFO :: Entered data_recon() function")
    Udate = set()
    mail_body = f"Data Recon of data uploaded on {current_date.strftime('%d-%m-%Y')}\n"
    try:
        connection = mysql.connector.connect(host=config["DB Config"]["host"],
                                                port=config["DB Config"]["port"],
                                                database=config["DB Config"]["database"],
                                                user=config["DB Config"]["user_name"],
                                                password=config["DB Config"]["paswd"])
        if connection.is_connected():
            logging.info("INFO :: Connection to Database is established")
            cursor = connection.cursor()
            for data in file_data:
                if data.get('make') != "suzlon_weekly":
                    Udate.add(data.get('gendate'))
            for x in Udate:
                check_value_daily = 0
                for data in file_data:
                    if data.get('gendate') == x and data.get('make') != "suzlon_weekly":
                        check_value_daily += data.get('mckwhday')
                cursor.execute(f"SELECT sum({'mckwhday'}) FROM spi_windmill_gen_daily_report where gendate='{x}';")
                db_value = cursor.fetchall()
                db_value = return_value(db_value[0][0])
                mail_body = mail_body+f"\nData Recon for {x} daily data\n"
                mail_body = mail_body+f"\t\t\tData from file : {check_value_daily}\n"
                mail_body = mail_body+f"\t\t\tData from database : {db_value}\n"
                mail_body = mail_body+f"\t\t\tThe Variance is : {check_value_daily - int(db_value)}\n"
            logging.info("INFO :: Daily mail have been analysed")
            check_value_weekly = []
            # check_value_weekly = [(data.get('ebkwhday'),datetime.strptime(data.get('gendate'),'%Y-%m-%d')) for data in file_data if data.get('make') == "suzlon_weekly"]
            for data in file_data:
                if data.get('make') == "suzlon_weekly":
                    print(data.get('gendate'))
                    check_value_weekly.append((data.get('ebkwhday'),datetime.strptime(data.get('gendate'),'%Y-%m-%d')))
            if check_value_weekly:
                max_val = max([cw[1] for cw in check_value_weekly]).strftime('%Y-%m-%d')
                min_val = min([cw[1] for cw in check_value_weekly]).strftime('%Y-%m-%d')
                sum_val = [return_value(cw[0]) for cw in check_value_weekly]
                sum_val = sum(sum_val)
                cursor.execute(f"SELECT sum(ebkwhday) FROM spi_windmill_gen_daily_report where gendate between '{min_val}' and '{max_val}';")
                weekly_db_value = cursor.fetchall()
                weekly_db_value = return_value(weekly_db_value[0][0])
                mail_body = mail_body+f"\n\nData Recon for weekly data from {min_val} to {max_val}\n"
                mail_body = mail_body+f"\t\t\tData from file : {sum_val}\n"
                mail_body = mail_body+f"\t\t\tData from database : {weekly_db_value}\n"
                mail_body = mail_body+f"\t\t\tThe Variance is : {sum_val- int(weekly_db_value)}\n"
                # min([datetime.strptime(x,'%d-%b-%y') for x in content.split('\n')])
            logging.info("INFO :: Weekly mail have been analysed")
        cursor.close()
        connection.close()
        sending_mail(f"RAPBot Recon for {', '.join(Udate)} data",mail_body,"ADMIN")
    except Exception as e:
        logging.error(f"ERROR :: Occured in data_recon function Error is : {e}")
        sending_mail("RAPBot Error notification",f"Error occured in data_recon() fucntion \nerror is : {e}\ndetailed error : {traceback.print_exc()}","Admin")

def convert_time_zone(date_obj):
    if datetime.utcnow().hour == datetime.now().hour:
        from_tz = tz.gettz('UTC')
        to_tz = tz.gettz('Asia/Kolkata')
        date_obj = date_obj.replace(tzinfo=from_tz)
        date_obj = date_obj.astimezone(to_tz)
        return date_obj
    else:
        return date_obj.replace(tzinfo=to_tz)


def insert_time_zone(date_obj):
    return date_obj.replace(tzinfo=tz.gettz('Asia/Kolkata'))


# def last_mod_time(fname):
#     folder_time = os.path.getmtime(fname)
#     return os.path.getmtime(fname)


# def move_zip_file(browser, customer_type):
#     SECONDS_IN_DAY = 400
#     # now= datetime.now().time().second
#     now = time.time()
#     before = now - SECONDS_IN_DAY
#     copy_path = download_file_path
#     #  logging.info(f"RAPBot has started moving the file to {output_path}")
#     for file_name in os.listdir(copy_path):
#         target_path = os.path.join(copy_path, file_name)
#         if last_mod_time(target_path) > before:
#             return file_name


def sending_mail(subject, body_mes, mail_type):
    msg = MIMEMultipart()
    msg['From'] = config["Login Details"]["user_name"]  # from address
    # msg['To'] = "jaikishore1997@gmail.com" # to address
    if mail_type.lower() == 'bussiness':
        msg['To'] = ", ".join(config['Report']['mail'].split(","))
    elif mail_type.lower() == 'admin':
        msg['To'] = config["Admin"]["mail"]

    # msg['Subject'] = "Choice Reports RAPBot notification"
    msg['Subject'] = subject
    body = f'{body_mes}'
    msg.attach(MIMEText(body, 'plain'))
    # put your relevant SMTP here
    server = smtplib.SMTP('smtp.gmail.com', '587')
    server.ehlo()
    server.starttls()
    server.ehlo()
    # server.login(your mail id, your password)  ### if applicable
    server.login(config["Login Details"]["user_name"],
                 config["Login Details"]["password"])
    server.send_message(msg)
    server.quit()


def login_gmail(browser):
    try:
        WebDriverWait(browser, 45).until(EC.element_to_be_clickable((By.XPATH, '//input[@type="email" and @aria-label="Email or phone"]'))).send_keys(config['Login Details']['user_name'])
        # WebDriverWait(browser, 45).until(EC.element_to_be_clickable((By.XPATH,'//input[@id="identifierId"]'))).send_keys(config['Login Details']['user_name'])
        logging.info("INFO :: Email Entered")
        browser.find_element_by_xpath('//button[span="Next"]').click()
        WebDriverWait(browser, 45).until(EC.element_to_be_clickable(
            (By.XPATH, '//input[@type="password" and @aria-label="Enter your password"]'))).send_keys(config['Login Details']['password'])
        logging.info("INFO :: Password Entered")
        browser.implicitly_wait(10)
        WebDriverWait(browser, 30).until(EC.element_to_be_clickable(
            (By.XPATH, '//*[@id="passwordNext"]/div/button'))).click()
        browser.implicitly_wait(20)
        logging.info("INFO :: Logged in successfully")
        validate_mail(browser)
    except exceptions.StaleElementReferenceException as se:
        print(f"Stale element exception occured : {se}")
        logging.error(f"ERROR :: Stale element exception occured : {se}")
        # browser.quit()
        start_program(browser)
        # browser.get(config["Website"]["url"])
        # login_gmail(browser)
    except Exception as e:
        print(f"Error occured in login_mail : {e}")
        logging.error(f"ERROR :: Error Occured in login_mail function --------------------> {e}")
        browser.save_screenshot('Screeshot.png')
        # sending_mail("RAP Bot Error Notification",
        #              f'Error Occured in login_mail function --------------------> {e}', 'Admin')


def download_button_click(browser, mail_val, exception_case_file=False):
    customer_type = mail_val[1]
    logging.info(
        f"INFO :: It came to download_button_click fucntion. The mail type : {customer_type} and the subject : {mail_val[0]}")
    browser.implicitly_wait(30)
    if exception_case_file:
        logging.info(
            f"INFO :: The mail {mail_val[0]} sent at {mail_val[2]} is an exception run")
    ele_len = len(browser.find_elements_by_xpath(
        f'//div[@class="aQH"]/span[@download_url]'))
    logging.info(
        f"INFO :: No. of files attached to {mail_val[0]} {mail_val[2]} mail is : {ele_len}")
    print("No of files in email : ", ele_len)
    if ele_len:
        for index in range(1, ele_len+1):
            filename = browser.find_element_by_xpath(
                f'//div[@class="aQH"]/span[@download_url][{index}]/div[1]/div[@data-tooltip="Download"]').get_attribute('aria-label')
            filename = filename.replace("Download attachment", "").strip()
            print(filename)
            # destination_path = ''
            if '.xls' in filename or '.xlsx' in filename:
                if 'tenksasi ' not in filename.lower() and 'vestas_daily_ss' not in filename.lower():
                    action = ActionChains(browser)
                    file_element = browser.find_element_by_xpath(
                        f'//div[@class="aQH"]/span[@download_url][{index}]/a')
                    action.move_to_element(file_element).perform()
                    browser.find_element_by_xpath(
                        f'//div[@class="aQH"]/span[@download_url][{index}]/div[1]/div[@data-tooltip="Download"]').click()
                    time.sleep(15)
                    move_downloaded_file(browser, customer_type, filename) if not exception_case_file else move_downloaded_file(
                        browser, customer_type, filename, True)
    else:
        # send mail for no document attached in mail
        logging.error(
            f"ERROR :: No files are attached in {mail_val[0]} sent at {mail_val[2]}")
        sending_mail('RAP Bot Error Notification',
                     f'No attachments are found in {mail_val[0]} sent on {mail_val[2]} Please do check to it', 'Admin')        
        sending_mail('RAP Bot Error Notification',
                     f'No attachments are found in {mail_val[0]} sent on {mail_val[2]} Please do check to it', 'Bussiness')


def email_back_button_click(browser):
    try:
        email_back = ActionChains(browser)
        back_button = browser.find_element_by_xpath(
            '//*[@id=":4"]/div[2]/div[1]/div/div[1]/div')
        email_back.move_to_element(back_button)
        WebDriverWait(browser, 20).until(EC.element_to_be_clickable(
            (By.XPATH, '//div[@title="Back to Inbox" and @role="button"]'))).click()
    except Exception as e:
        logging.error(
            "ERROR :: Exception occured in email_back_button_click function message : {}".format(e))


def validate_mail(browser):
    logging.info("INFO :: Entered Validate_mail function")
    count = 0
    WebDriverWait(browser, 30).until(EC.element_to_be_clickable(
        (By.XPATH, '//tr[contains(@class,"zA")]')))
    element_len = browser.find_elements_by_xpath(
        f'//tr[contains(@class,"zA")]')
    subject_val = {}
    file_path = []
    mail_tracker = {}
    # suzlonCheckFilePath = os.path.dirname(__file__)+"/suzlonCheck.json"
    current_date = convert_time_zone(datetime.now())
    for x in config['Subject']:
        for y in config['Subject'][x].split(','):
            subject_val[y] = x
    exception_flag = [True for x in config['Exception']
                      if str(config['Exception'][x]).upper() == "ON"]
    if exception_flag:
        logging.info(f"INFO :: Exception flag status : True")
    else:
        logging.info(f"INFO :: Exception flag status : False")
    try:
        if not exception_flag:
            for ele in range(1, len(element_len)+1):
                print("Mailtrack : ",mail_tracker)
                # print("\nFile data : ",file_data)
                content_page = browser.find_elements_by_xpath(
                    '//*[@id=":1"]/div/div[2]/div/table/tr/td[1]/div[2]')
                check_count = 0
                while len(content_page) != 0:
                    print("Waiting in while loop.........")
                    check_count += 1
                    if check_count > 20:
                        email_back_button_click(browser)
                        print("Again back button clicked")
                    logging.info(
                        "INFO :: -----------------> Waiting in loop while loop in valide_mail()")
                    content_page = browser.find_elements_by_xpath(
                        '//*[@id=":1"]/div/div[2]/div/table/tr/td[1]/div[2]')
                print("\nEnded loop")
                logging.info("INFO :: -----------> While loop ended")
                element = browser.find_element_by_xpath(
                    f'//tr[contains(@class,"zA")][{ele}]/td[5]/div')
                mail_check_elemt = browser.find_element_by_xpath(
                    f'//tr[contains(@class,"zA")][{ele}]/td[4]/div[1]/span/span').get_attribute('email')
                time_check = browser.find_element_by_xpath(
                    f'//tr[contains(@class,"zA")][{ele}]/td[8]/span').get_attribute('title')
                mail_id = [config['Mail'][x] for x in config['Mail']]
                is_customer_mail = False

                print("The Current date is : ",current_date)
                logging.info(f"INFO :: Current date is {str(current_date)}")
                mail_recived_time = convert_time_zone(
                    datetime.strptime(time_check, '%a, %b %d, %Y, %I:%M %p'))
                subject_check = browser.find_element_by_xpath(
                    f'//tr[contains(@class,"zA")][{ele}]/td[5]/div[1]/div[1]/div[1]/span/span').text
                print(mail_check_elemt, " : ", time_check, " : ", subject_check)
                customer_type = ''
                for x in subject_val:
                    if x in subject_check:
                        is_customer_mail = True
                        customer_type = subject_val[x]
                        break
                print(customer_type)
                try:
                    if is_customer_mail and current_date.date() == mail_recived_time.date():
                        element.click()
                        logging.info("INFO :: Mail Element clicked")
                        mail_element = browser.find_elements_by_xpath(
                            f'//div[@class="aQH"]/span[@download_url]')
                        vestas_limit_end_time = datetime.strptime(
                            config["Mail Time"]['vestas_end_time'], '%I:%M %p')
                        vestas_limit_end_time = insert_time_zone(vestas_limit_end_time.replace(
                            day=current_date.day, month=current_date.month, year=current_date.year))
                        vestas_limit_start_time = datetime.strptime(
                            config["Mail Time"]['vestas_start_time'], '%I:%M %p')
                        vestas_limit_start_time = insert_time_zone(vestas_limit_start_time.replace(
                            day=current_date.day, month=current_date.month, year=current_date.year))

                        suzlon_limit_end_time = datetime.strptime(
                            config["Mail Time"]['suzlon_end_time'], '%I:%M %p')
                        suzlon_limit_end_time = insert_time_zone(suzlon_limit_end_time.replace(
                            day=current_date.day, month=current_date.month, year=current_date.year))
                        suzlon_limit_start_time = datetime.strptime(
                            config["Mail Time"]['suzlon_start_time'], '%I:%M %p')
                        suzlon_limit_start_time = insert_time_zone(suzlon_limit_start_time.replace(
                            day=current_date.day, month=current_date.month, year=current_date.year))

                        # suzlon daily download
                        if "suzlon" in customer_type and "daily" in subject_check.lower():
                            print(mail_check_elemt in mail_id)
                            print(suzlon_limit_start_time < mail_recived_time)
                            print(suzlon_limit_end_time > mail_recived_time)
                            mail_val = [subject_check, customer_type, time_check]
                            if mail_check_elemt and mail_check_elemt in mail_id and suzlon_limit_end_time > mail_recived_time and suzlon_limit_start_time < mail_recived_time:
                                logging.info(
                                    f"INFO :: Mail is {customer_type} type and the subject is '{subject_check}'")
                                print("Suzlon_daily is selected")
                                if 'suzlon_daily' in mail_tracker:
                                    mail_tracker['suzlon_daily'].append(subject_check)
                                else:
                                    mail_tracker['suzlon_daily'] = [subject_check]
                                download_button_click(browser, mail_val)
                            if int(config["Bot"]["morning_mail_process"]) and suzlon_limit_start_time > mail_recived_time and not mail_tracker.get('suzlon_daily'):
                                #  and mail_tracker.get('suzlon_daily') < config['No. of Mails']['suzlon_daily'] 
                                if 'suzlon_daily_morning' in mail_tracker:
                                    mail_tracker['suzlon_daily_morning'].append(subject_check)
                                else:
                                    mail_tracker['suzlon_daily_morning'] = [subject_check]
                                download_button_click(browser, mail_val)
                            

                        # suzlon weekly download
                        elif "suzlon" in customer_type and "weekly" in subject_check.lower():
                            logging.info(
                                f"INFO :: Mail is {customer_type} type and the subject is '{subject_check}'")
                            mail_val = [subject_check,
                                        customer_type, time_check]
                            print("Suzlon_weekly is selected")
                            if 'suzlon_weekly' in mail_tracker:
                                mail_tracker['suzlon_weekly'].append(subject_check)
                            else:
                                mail_tracker['suzlon_weekly'] = [subject_check]
                            download_button_click(browser, mail_val)

                        # vestas daily download
                        elif "vestas" in customer_type:
                            if mail_recived_time < vestas_limit_end_time and mail_recived_time > vestas_limit_start_time:
                                logging.info(f"INFO :: Mail is {customer_type} type and the subject is '{subject_check}'")
                                mail_val = [subject_check,
                                            customer_type, time_check]
                                print("Vestas_daily is selected")
                                if 'vestas_daily' in mail_tracker:
                                    mail_tracker['vestas_daily'] += 1
                                else:
                                    mail_tracker['vestas_daily'] = 1
                                download_button_click(browser, mail_val)

                        # email click back button
                        email_back_button_click(browser)
                    # END of normal flow program
                    elif current_date.date() > mail_recived_time.date():
                        for swf in suzlon_weekly_file:
                            read_excel_file(browser, swf, 'suzlon_weekly')
                        suzlon_weekly_file.clear()
                        logging.info("INFO :: Daily normal flow program came to and END")
                        break
                except Exception as e:
                    logging.error(
                        f"ERROR :: Mail : {mail_check_elemt} ,Time : {time_check} ,Subject:{subject_check}     Error occured in validate_mail function -------> {e}")
                    sending_mail(str(config['Mail Content']['error_notify']),
                                 f"Mail : {mail_check_elemt} ,Time : {time_check} ,Subject:{subject_check} \nError Occured in validate_mail function ----------------------------> {e}", "Admin")
            # check all the mail are sent properly and notify admin in case of error
            data_recon(file_data,current_date)
            send_error_mail(browser,mail_tracker,current_date)
        else:
            logging.info("INFO :: Exception is enabled")
            exception_case(browser)
    except Exception as e:
        logging.error(
            f"INFO :: Error has occured in validate_mail function----------------> {e}")
        sending_mail("Error RAP Bot Notification",f"Error occured in validate_mail function. \nError is {e} \n detailed error : {traceback.print_exc()}","Admin")


def send_error_mail(browser, mail_tracker, today, suzlonCheckFilePath=None):
    print("Entered send_error_mail function")
    print("Mail Tracker is : ",mail_tracker)
    logging.info("INFO :: Entered send_error_mail() function")
    for x in mail_tracker:
        logging.info(f"INFO :: {x} count : {len(mail_tracker[x]) if 'suzlon' in x else mail_tracker[x]}")
    if mail_tracker:
        # for x in mail_tracker:
        # Check for suzlon morning mail
        if mail_tracker.get('suzlon_daily'):
            if len(mail_tracker.get('suzlon_daily')) < int(config['No. of Mails']['suzlon_daily']):
                # send mail for less no of emails sent for suzlon daily
                spi_daily_count = 0
                kr_daily_count = 0
                for y in mail_tracker.get('suzlon_daily'):
                    if 'spi' in y.lower() or 'skr' in y.lower():
                        spi_daily_count += 1
                    elif 'k r' in y.lower():
                        kr_daily_count += 1
                sending_mail("ERROR RAP Bot Notification",f"On {today.strftime('%d/%m/%Y')} Suzlon daily have less no. mail of  recieved :\n SPI Power : {spi_daily_count}\n KR Wind Energy : {kr_daily_count}","ADMIN")
        elif mail_tracker.get('suzlon_daily_morning'):
            if len(mail_tracker.get('suzlon_daily_morning')) < int(config['No. of Mails']['suzlon_daily']):
                # send mail for less no of emails sent for suzlon daily
                spi_daily_count = 0
                kr_daily_count = 0
                for y in mail_tracker.get('suzlon_daily_morning'):
                    if 'spi' in y.lower() or 'skr' in y.lower():
                        spi_daily_count += 1
                    elif 'k r' in y.lower():
                        kr_daily_count += 1
                sending_mail("ERROR RAP Bot Notification",f"On {today.strftime('%d/%m/%Y')} Suzlon daily Evenging mail have not sent so morning mails have been processed in morning mail no. mail of  recieved :\n SPI Power : {spi_daily_count}\n KR Wind Energy : {kr_daily_count}.","ADMIN")
            else:
                sending_mail("ERROR RAP Bot Notification",f"On {today.strftime('%d/%m/%Y')} Suzlon daily Evenging mail have not sent so morning mails have been processed without any deficiency.","ADMIN")
        else:
            sending_mail("ERROR RAP BOT Notification",f"On {today.strftime('%d/%m/%Y')} Suzlon daily mails are not recieved at time","ADMIN")

        if mail_tracker.get('vestas_daily'):
            if mail_tracker.get('vestas_daily') < int(config['No. of Mails']['vestas_daily']):
                print("Vestas daily mail is not yet sent")
                sending_mail("ERROR RAP Bot notification",f"On {today.strftime('%d/%m/%Y')} Vestas daily have less no. mail of  recieved :\n Vestas : {config['No. of Mails']['vestas_daily']}","ADMIN")
                # send mail for less no of emails sent for vestas daily
        elif today.strftime('%a') != "Sun":
            sending_mail("ERROR RAP BOT Notification",f"On {today.strftime('%d/%m/%Y')} Vestas daily mails are not recieved","ADMIN")
    else:
        sending_mail("ERROR RAP BOT Notification",f"On {today.strftime('%d/%m/%Y')} No mails have been recived.","ADMIN")

    if today.strftime('%a') == 'Fri' and suzlonCheckFilePath:
        if os.path.exists(suzlonCheckFilePath):
            with open(suzlonCheckFilePath, "r") as suzlonVal:
                suzlonWeekDataRead = json.loads(suzlonVal.read())
            if int(suzlonWeekDataRead['Download']) < 2:
                # send mail notif. for suzlon weekly not been sent
                sending_mail("RAP Bot notification",
                             "Suzlon weekly mail is not yet sent")


def exception_case(browser, customer_type=None):
    print("entered exception case")
    logging.info("INFO :: Entered exception_case fucntion")
    if customer_type:
        exception_customer = customer_type
    else:
        exception_customer = [x for x in config['Exception']
                              if config['Exception'][x].upper() == "ON"]

    # element_len = browser.find_elements_by_xpath(f'//tr[contains(@class,"zA")]')
    WebDriverWait(browser, 30).until(EC.element_to_be_clickable(
        (By.XPATH, '//tr[contains(@class,"zA")]')))
    element_len = browser.find_elements_by_xpath(
        f'//div[@gh]/div[2]/div[1]/table/tbody/tr[contains(@class,"zA")]')
    search_over = {}
    for type in exception_customer:
        search_over[type] = True
    try:
        print("entered try block")
        for email_index in range(1, len(element_len)+1):
            content_page = browser.find_elements_by_xpath(
                '//*[@id=":1"]/div/div[2]/div/table/tr/td[1]/div[2]')
            check_count = 0
            while len(content_page) != 0:
                print("Waiting in while loop.........")
                logging.info("INFO :: Waiting in while loop............")
                check_count += 1
                if check_count > 20:
                    email_back_button_click(browser)
                    print("Again back button clicked")
                logging.info(
                    "INFO :: -----------------> Waiting in loop while loop in exception_case()")
                content_page = browser.find_elements_by_xpath(
                    '//*[@id=":1"]/div/div[2]/div/table/tr/td[1]/div[2]')
            email_element = browser.find_element_by_xpath(
                f'//div[@gh]/div[2]/div[1]/table/tbody/tr[contains(@class,"zA")][{email_index}]/td[5]/div')
            print("Endedd while loop")
            logging.info("INFO :: --------Loop wait ended----------")
            for company in exception_customer:
                mail_subject = config['Subject'][company]
                subject_val = ''
                mail_subject = mail_subject.split(',')
                # mail_time = datetime.strptime(config['Exception Config'][company],'%d/%m/%Y %I:%M %p')
                mail_time = config['Exception Config'][company].split('-')
                mail_time = [insert_time_zone(datetime.strptime(
                    x.strip(), '%d/%m/%Y %I:%M %p')) for x in mail_time]
                mail_check_elemt = browser.find_element_by_xpath(
                    f'//div[@gh]/div[2]/div[1]/table/tbody/tr[contains(@class,"zA")][{email_index}]/td[4]/div[1]/span/span').get_attribute('email')
                time_check = browser.find_element_by_xpath(
                    f'//div[@gh]/div[2]/div[1]/table/tbody/tr[contains(@class,"zA")][{email_index}]/td[8]/span').get_attribute('title')
                print(time_check)
                subject_check = browser.find_element_by_xpath(
                    f'//div[@gh]/div[2]/div[1]/table/tbody/tr[contains(@class,"zA")][{email_index}]/td[5]/div[1]/div[1]/div[1]/span/span').text
                time_check = datetime.strptime(
                    time_check, '%a, %b %d, %Y, %I:%M %p')
                time_check = convert_time_zone(time_check)
                print("Converted time zone : ", time_check)
                # if mail_subject in subject_check and time_check == mail_time:
                print(email_index, " ", subject_check,
                      " : ", time_check, " : ", company)
                logging.info(
                    f"INFO :: ------------------------> In Index {email_index} This subject '{subject_check}' on time '{time_check}' belongs to {company} was in exception")
                is_check_mail = False
                for x in mail_subject:
                    if x in subject_check:
                        is_check_mail = True
                        break
                if is_check_mail:
                    logging.info(f"INFO :: Exception time : {str(mail_time[0])}")
                    logging.info(f"INFO :: Mail time : {time_check}")
                    print('-------->', mail_time[0])
                    print(mail_time[1])
                    print(time_check)
                    if mail_time[0] <= time_check and time_check <= mail_time[1]:
                        logging.info(
                            "INFO :: ########## Time matched and condition of exception case matched for this mail ##########")
                        email_element.click()
                        logging.info("INFO :: Mail element clicked")
                        print("Email element CLicked")
                        mail_val = [subject_check, company, time_check]
                        download_button_click(browser, mail_val, True)
                        email_back_button_click(browser)
                        break
                print(mail_time[0] > time_check)
                print(search_over)
                print(mail_time[0])
                print(time_check)
                if mail_time[0] > time_check:
                    search_over[company] = False
                print('^^^^^^^^^^', any([search_over[x] for x in search_over]))
            # End of exception run
            if not any([search_over[x] for x in search_over]):
                logging.info("INFO :: Started to process Suzlon Weekly")
                for swf in suzlon_weekly_file:
                    read_excel_file(browser, swf, 'suzlon_weekly')
                suzlon_weekly_file.clear()
                logging.info(
                    "INFO :: --------------- End of exception run ---------------")
                break
        if any([search_over[x] for x in search_over]):
            logging.info(
                "INFO :: Enter neext page of gmail in exception_run() function")
            next_page = browser.find_element_by_xpath(
                f'//*[contains(@id,":i") and @data-tooltip="Older"]')
            hoverAction = ActionChains(browser)
            hoverAction.move_to_element(next_page).perform()
            next_page.click()
            browser.implicitly_wait(20)
            exception_case(browser, [x for x in search_over if search_over[x]])
    except Exception as e:
        print(f"The exceptin occured in exception function is : {e}")
        sending_mail("RAP Bot error notification","Error Occured at exception_case() function : {}".format(e),'Admin')
        logging.info(f"ERROR :: Error has occured in exception case function : {e}")


def move_downloaded_file(browser, customer_type, file_name, exception=None):
    logging.info("INFO :: Bot came to move_downloaded_file() function")
    if exception:
        logging.info(f"INFO :: {file_name} is downloaded as a exception run")
    # dow_path = os.path.join(download_file_path,file_name)
    dow_path = download_file_path+'/'+file_name
    file_date = datetime.now().strftime("%d")
    file_month = datetime.now().strftime("%m")
    print("Download path : {}".format(dow_path))
    if not exception:
        des_path = os.path.join(
            copy_file_path, "SPI/{}/{}".format(file_month, file_date))
        # des_path = copy_file_path+"/"+"SPI/{}/{}".format(file_month,file_date)
    else:
        des_path = os.path.join(
            copy_file_path, "SPI/{}/{}/{}".format(file_month, file_date, "Exception_run"))
        # des_path = copy_file_path+"/"+"SPI/{}/{}/{}".format(file_month,file_date,"Exception_run")
    print(des_path)
    os.makedirs(des_path+"/{}".format(customer_type), exist_ok=True)
    try:
        shutil.move(dow_path, des_path+"/{}".format(customer_type))
    except Exception as e:
        print("Move file exception")
        if 'already exists' in e.args[0].lower():
            os.remove(dow_path)
            logging.error(f"ERROR :: Error occured in move_downloaded_file() function Already file exist error --------------> {e}")
        else:
            sending_mail("RAP Bot Error notification","Error occured at move_downloaded_file() function ERROR : {}".format(e),"ADMIN")
            logging.error(f"ERROR :: Error occured in move_downloaded_file() function --------------> {e}")

    print("---> File moved")
    logging.info("INFO :: File moved to {}/{}".format(des_path,customer_type))
    excel_file_path = des_path+"/{}/{}".format(customer_type, file_name)
    print("Destination Path: ", excel_file_path, "   ", customer_type)
    logging.info(f"INFO :: File is moved in path {excel_file_path}")
    print(excel_file_path)
    if 'suzlon_weekly' in customer_type.lower():
        suzlon_weekly_file.append(excel_file_path)
    else:
        read_excel_file(browser, excel_file_path, customer_type)
    # excel_file_path = os.path.join(des_path,customer_type,file_name)


def read_excel_file(browser, file_path, customer_type):
    logging.info("INFO :: Entered read_excel_file() function")

    def check_valuein_reporting_layer(cursor, query_val):
        check_command = f"select * from spi_windmill_gen_daily_report where gendate='{query_val[0]}' and companyname='{query_val[1]}' and locno='{query_val[2]}';"
        cursor.execute(check_command)
        fetched_data = cursor.fetchall()
        return False if fetched_data else True

    def read_location_master(cursor):
        cursor.execute("SELECT locno,make,section,site,weghtno FROM spi_group_windmill_data.location_master;")
        location_data = cursor.fetchall()
        location_dic = {}
        for x in location_data:
            location_dic[x[0]] = [x[1], x[2], x[3], x[4]]
        logging.info("INFO :: Loaded data from location_master")
        return location_dic

    def check_float_val(data):
        try:
            float(data)
            return data if float(data) else 0.0
        except:
            return 0.0
    try:
        connection = mysql.connector.connect(host=config["DB Config"]["host"],
                                             port=config["DB Config"]["port"],
                                             database=config["DB Config"]["database"],
                                             user=config["DB Config"]["user_name"],
                                             password=config["DB Config"]["paswd"])
        if connection.is_connected():
            logging.info("INFO :: Database connection is successfull established")
            cursor = connection.cursor()
            location = read_location_master(cursor)
            try:
                if "suzlon_daily" in customer_type:
                    sheet_val = pd.read_excel(file_path, sheet_name=None)
                    file_name = file_path.split('/')[-1]
                    duplicate_record_sd = []
                    recordInserted = []
                    for sheet_name in sheet_val:
                        if "generation" in sheet_name.lower():
                            logging.info(f"INFO :: INSERT GENERATION Data: {file_name}")
                            doc_val = sheet_val[sheet_name].fillna('')
                            for y in doc_val.columns:
                                if "date" in y.lower():
                                    doc_val.rename(
                                        columns={y: 'genDate'}, inplace=True)
                                if "customer" in y.lower() or 'company' in y.lower():
                                    doc_val.rename(
                                        columns={y: 'customerName'}, inplace=True)
                                if "state" in y.lower() or "site" in y.lower() or "section" in y.lower() or y.lower() == "mw" or y.lower() == "gf" or y.lower() == "fm" or y.lower() == "s" or y.lower() == "u" or y.lower() == "nor" or y.lower() == 'rna':
                                    doc_val.rename(
                                        columns={y: y.lower()}, inplace=True)
                                if 'htsc' in y.lower():
                                    doc_val.rename(
                                        columns={y: 'htscNo'}, inplace=True)
                                if 'loc' in y.lower():
                                    doc_val.rename(
                                        columns={y: 'locNo'}, inplace=True)
                                if 'gen' in y.lower() and 'day' in y.lower():
                                    doc_val.rename(
                                        columns={y: 'genkwhDay'}, inplace=True)
                                if 'gen' in y.lower() and 'mtd' in y.lower():
                                    doc_val.rename(
                                        columns={y: 'genkwhMtd'}, inplace=True)
                                if 'gen' in y.lower() and 'ytd' in y.lower():
                                    doc_val.rename(
                                        columns={y: 'genkwhYtd'}, inplace=True)
                                if 'plf' in y.lower() and 'day' in y.lower():
                                    doc_val.rename(
                                        columns={y: 'plfDay'}, inplace=True)
                                if 'plf' in y.lower() and 'mtd' in y.lower():
                                    doc_val.rename(
                                        columns={y: 'plfMtd'}, inplace=True)
                                if 'plf' in y.lower() and 'ytd' in y.lower():
                                    doc_val.rename(
                                        columns={y: 'plfYtd'}, inplace=True)
                                if 'avail' in y.lower():
                                    doc_val.rename(
                                        columns={y: 'mcAvail'}, inplace=True)
                                if 'hrs' in y.lower():
                                    if 'gen' in y.lower():
                                        doc_val.rename(
                                            columns={y: 'genHrs'}, inplace=True)
                                    else:
                                        doc_val.rename(
                                            columns={y: 'oprHrs'}, inplace=True)
                            try:
                                logging.info(
                                    "INFO :: -------- > Table used : suzlon_xl_daily_hist and spi_windmill_gen_daily_report")
                                for column_val in doc_val.iterrows():
                                    x = column_val[1]
                                    if re.match(r"\d{4}-\d{2}-\d{2}\s\d{2}:\d{2}:\d{2}", str(x.get('genDate'))) or re.match(r"\d{2}-[A-z]{3}-\d{4}", str(x.get('genDate'))):
                                        genDate = str(x.get('genDate')).split(' ')[0] if re.match(r"\d{4}-\d{2}-\d{2}\s\d{2}:\d{2}:\d{2}", str(
                                            x[0])) else datetime.strptime(x.get('genDate'), "%d-%b-%Y").strftime("%Y-%m-%d")
                                        db_command = f"insert into suzlon_xl_daily_hist(gendate,customername,state,site,section,mw,locno,genkwhday,genkwhmtd,genkwhytd,plfday,plfmtd,plfytd,mcavail,gf,fm,s,u,nor,genhrs,oprhrs) values('{genDate}','{x.get('customerName')}','{x.get('state')}','{x.get('site')}','{x.get('section')}',{float(check_float_val(x.get('mw')))},'{x.get('locNo')}',{float(check_float_val(x.get('genkwhDay')))},{float(check_float_val(x.get('genkwhMtd')))},{float(check_float_val(x.get('genkwhYtd')))},{float(check_float_val(x.get('plfDay')))},{float(check_float_val(x.get('plfMtd')))},{float(check_float_val(x.get('plfYtd')))},{float(check_float_val(x.get('mcAvail')))},{float(check_float_val(x.get('gf')))},{float(check_float_val(x.get('fm')))},{float(check_float_val(x.get('s')))},{float(check_float_val(x.get('u')))},{float(check_float_val(x.get('nor',x.get('rna'))))},{float(check_float_val(x.get('genHrs')))},{float(check_float_val(x.get('oprHrs')))});"
                                        customerName = "SPI Power" if "spi" in re.sub(r"\s+", '', x.get('customerName')).lower() or "skr" in re.sub(r"\s+", '', x.get(
                                            'customerName')).lower() else "KR Wind Energy" if "kr" in re.sub(r"\s+", '', x.get('customerName')).lower() else ''
                                        locNoVal = re.sub(
                                            r"\s+", '', x.get('locNo')) if "TP06" not in x.get('locNo') else "TP6"
                                        location_values = location.get(
                                            locNoVal)
                                        db_command2 = f"insert into spi_windmill_gen_daily_report(gendate,companyname,locno,mckwhday,gf,fm,sch,unsch,genhrs,oprhrs,mw,section,site,make,htno) values('{genDate}','{customerName}','{locNoVal}',{float(check_float_val(x.get('genkwhDay')))},{float(check_float_val(x.get('gf')))},{float(check_float_val(x.get('fm')))},{float(check_float_val(x.get('s')))},{float(check_float_val(x.get('u')))},{float(check_float_val(x.get('genHrs')))},{float(check_float_val(x.get('oprHrs')))},{float(check_float_val(x.get('mw')))},'{x.get('section')}','{x.get('site')}','{location_values[0]}','{location_values[3]}');"
                                        cursor.execute(db_command)
                                        suzlon_daily_check = check_valuein_reporting_layer(
                                            cursor, [str(x.get('genDate')).split(' ')[0], customerName, x.get('locNo')])
                                        if suzlon_daily_check:
                                            cursor.execute(db_command2)
                                            file_data.append({'gendate':genDate,'mckwhday':check_float_val(x.get('genkwhDay')),'make':location_values[0]})
                                            recordInserted.append(db_command2)
                                        else:
                                            duplicate_record_sd.append(
                                                "{} - {} - {}".format(x.get('genDate'), customerName, x.get('locNo')))
                                print(
                                    "\n\nSuccessfully Inserted in suzlon daily\n\n")
                                logging.info(
                                    f"INFO :: Data from {file_name} is Successfully Inserted into suzlon_xl_daily_hist Database")
                                if recordInserted:
                                    sending_mail(
                                        f"RAP Bot Successfull data uploaded notification for {customer_type}", f"Data from {file_name} is Successfully Inserted into  Database", "Admin")
                                    logging.info(f"INFO :: Data from {file_name} is Successfully Inserted into spi_windmill_gen_daily_report Database")
                            except Exception as e:
                                dataBaseError.append(e)
                                logging.error(
                                    f"ERROR :: An error occured while inserting GENERAL data from {file_name} into suzlon_xl_daily_hist and spi_windmill_gen_daily_report Database ----------------------> {e}")
                                sending_mail(f"RAP Bot notification for error in Database insert",
                                             f"GENERATION DATA (general sheet) from {file_name} or {customer_type} type is not Inserted into  Database Error is : {e}", "Admin")
                        elif 'break' in sheet_name.lower():
                            logging.info(f"INFO :: INSERT BREAKDOWN Data: {file_name}")
                            df = sheet_val[sheet_name].fillna("")
                            for y in df.columns:
                                if 'gen' in y.lower() and 'date'in y.lower():
                                    df.rename(columns={y:'genDate'},inplace=True)
                                if "customer" in y.lower() or 'company' in y.lower():
                                    df.rename(columns={y:'customerName'},inplace=True)
                                if "state" in y.lower() or "site" in y.lower() or "section" in y.lower() or y.lower() == "mw":
                                    df.rename(columns={y:y.lower()},inplace=True)
                                if 'loc' in y.lower():
                                    df.rename(columns={y:'locNo'},inplace=True)
                                if 'breakdown' in y.lower() and 'remark' in y.lower():
                                    df.rename(columns={y:'breakDownRemark'},inplace=True)
                                if 'formula' in y.lower() and 'parameter' in y.lower():
                                    df.rename(columns={y:'formulaParameter'},inplace=True)
                                if 'breakdown' in y.lower() and 'hrs' in y.lower():
                                    df.rename(columns={y:'breakDownHr'},inplace=True)
                            try:
                                for y_i,y in df.iterrows():
                                    if re.match(r'\d{2}-[A-z]{3}-\d{4}',str(y.get('genDate'))) or re.match(r'\d{4}-\d{2}-\d{2}\s\d{2}:\d{2}:\d{2}',str(y.get('genDate'))):
                                        date_time_val = datetime.strptime(y.get('genDate'),'%d-%b-%Y').strftime('%Y-%m-%d') if re.match(r'\d{2}-[A-z]{3}-\d{4}',str(y.get('genDate'))) else str(y.get('genDate')).split(' ')[0]
                                        db_query = f"insert into suzlon_breakdown_data(genDate,customername,state,site,section,mw,locno,remarks,breakdownhrs,parameter) values('{date_time_val}','{y.get('customerName')}','{y.get('state')}','{y.get('site')}','{y.get('section')}',{y.get('mw')},'{y.get('locNo')}','{y.get('breakDownRemark')}',{y.get('breakDownHr')},'{y.get('formulaParameter')}')"
                                        # print(db_query)
                                        cursor.execute(db_query)
                                    # else:
                                    #     print(y)
                            except Exception as be:
                                logging.error(f"ERROR :: Error occured while inserting {file_name} BREAK DOWN data : {be}")
                                sending_mail(f"RAP Bot notification for error in Database insert",f"BREAKDOWN Data(break down sheet) from {file_name} or {customer_type} type is not Inserted into  Database Error occured {e}", "Admin")

                    if duplicate_record_sd:
                        sending_mail(f"RAP Bot notification for duplicate records",
                                     f"Data from {file_name} file contains {str(duplicate_record_sd)} duplicate records", "Admin")

                if "vestas_daily" in customer_type:
                    df_dic = pd.read_excel(
                        file_path, sheet_name=None, header=None)
                    df_header = ""
                    file_name = file_path.split('/')[-1]
                    logging.info(
                        "INFO :: -------- > Table used : vestas_xl_daily_hist and spi_windmill_gen_daily_report")
                    try:
                        for sheet in df_dic:
                            df = df_dic[sheet].fillna('')
                            for x_i, x in df.iterrows():
                                if "date" in str(x[0]).lower():
                                    df_header = df.iloc[x_i-1:x_i+1].fillna('')
                                    break
                            df = df.iloc[x_i+1:]
                            head_val = ''
                            header = []
                            for x_i, x in df_header.iteritems():
                                if x.iloc[0]:
                                    head_val = 'cml_' if 'cumulative' in x.iloc[0].lower(
                                    ) else 'daily_' if 'daily' in x.iloc[0].lower() else ''
                                if 'date' in x.iloc[1].lower():
                                    x.iloc[1] = 'genDate'
                                if x.iloc[1].lower() == 'mw' or x.iloc[1].lower() == 'site':
                                    x.iloc[1] = x.iloc[1].lower()
                                if 'customer' in x.iloc[1].lower():
                                    x.iloc[1] = 'companyName'
                                if 'htno' in x.iloc[1].lower():
                                    x.iloc[1] = 'htno'
                                if 'loc' in x.iloc[1].lower():
                                    x.iloc[1] = 'locNo'
                                if 'reading' in x.iloc[1].lower() and 'taken' in x.iloc[1].lower():
                                    x.iloc[1] = 'reading_taken_time'
                                if 'hrs' in x.iloc[1].lower():
                                    if 'run' in x.iloc[1].lower():
                                        x.iloc[1] = head_val+"run_hr"
                                    if 'gen' in x.iloc[1].lower():
                                        x.iloc[1] = head_val+"gen_hr"
                                if x.iloc[1] == "g-0":
                                    x.iloc[1] = head_val+"g_0"
                                if x.iloc[1] == "GEN":
                                    x.iloc[1] = head_val+'gen'
                                if 'total' in x.iloc[1].lower() and 'prod' in x.iloc[1].lower():
                                    x.iloc[1] = head_val+"total_prod"
                                if 'total' in x.iloc[1].lower() and 'import' in x.iloc[1].lower():
                                    x.iloc[1] = head_val+"total_import"
                                if 'total' in x.iloc[1].lower() and 'export' in x.iloc[1].lower():
                                    x.iloc[1] = head_val+"total_export"
                                if x.iloc[1] == '06-09 am':
                                    if head_val+"06_09_am_1" in df_header.iloc[1].values:
                                        x.iloc[1] = head_val+"06_09_am_2"
                                    else:
                                        x.iloc[1] = head_val+"06_09_am_1"
                                if x.iloc[1] == '18-21 pm':
                                    if head_val+"18_21_pm_1" in df_header.iloc[1].values:
                                        x.iloc[1] = head_val+"18_21_pm_2"
                                    else:
                                        x.iloc[1] = head_val+"18_21_pm_1"
                                if x.iloc[1] == '21-22 pm':
                                    if head_val+"21_22_pm_1" in df_header.iloc[1].values:
                                        x.iloc[1] = head_val+"21_22_pm_2"
                                    else:
                                        x.iloc[1] = head_val+"21_22_pm_1"
                                if '05-06 am' in x.iloc[1]:
                                    if head_val+'05_06_am_&_09_18_pm_1' in df_header.iloc[1].values:
                                        x.iloc[1] = head_val + \
                                            '05_06_am_&_09_18_pm_2'
                                    else:
                                        x.iloc[1] = head_val + \
                                            '05_06_am_&_09_18_pm_1'
                                if 'rkvahr' in x.iloc[1]:
                                    if 'imp' in x.iloc[1]:
                                        x.iloc[1] = head_val+'rkvahr_imp'
                                    elif 'exp' in x.iloc[1]:
                                        x.iloc[1] = head_val+'rkvahr_exp'
                                if x.iloc[1] == '22-05 am':
                                    if head_val+"22_05_am_1" in df_header.iloc[1].values:
                                        x.iloc[1] = head_val+'22_05_am_2'
                                    else:
                                        x.iloc[1] = head_val+'22_05_am_1'
                                if 'grid' in x.iloc[1].lower() and 'failure' in x.iloc[1].lower():
                                    x.iloc[1] = "gf"
                                if 'feeder' in x.iloc[1].lower() and 'maintenance' in x.iloc[1].lower():
                                    x.iloc[1] = "fm"
                                if x.iloc[1] == "Scheduled Maintenance":
                                    x.iloc[1] = 'sch'
                                if x.iloc[1] == "Unscheduled Maintenance":
                                    x.iloc[1] = 'unsch'
                                if x.iloc[1] == "Manual Stoppage":
                                    x.iloc[1] = 'ms'
                                if x.iloc[1] == "Reading Not Avilable":
                                    x.iloc[1] = 'readNotAvail'
                                if x.iloc[1] == "Total" or x.iloc[1] == "Remarks":
                                    x.iloc[1] = x.iloc[1].lower()
                            try:
                                df.columns = df_header.iloc[1]
                                for column_val in df.iterrows():
                                    x = column_val[1]
                                    if re.match(r"\d{4}-\d{2}-\d{2}\s\d{2}:\d{2}:\d{2}", str(x.get('genDate'))):
                                        # try:
                                        db_command1 = f'INSERT into vestas_xl_daily_hist(gendate,mw,customername,htno,site,locno,readingtakentime,cml_runhrs,cml_genhrs,cml_g0,cml_gen,cml_totalprod,cml_totalimport,cml_06_09am_1,cml_18_21pm_1,cml_21_22pm_1,cml_05_06amand09_18pm_1,cml_22_05am_1,cml_totalexport,cml_06_09am_2,cml_18_21pm_2,cml_21_22pm_2,cml_05_06amand09_18pm_2,cml_22_05am_2,cml_rkvahr_imp,cml_rkvahr_exp,daily_runhrs,daily_genhrs,daily_g0,daily_gen,daily_totalprod,daily_totalimport,daily_06_09am_1,daily_18_21pm_1,daily_21_22pm_1,daily_05_06amand09_18pm_1,daily_22_05am_1,daily_totalexport,daily_06_09am_2,daily_18_21pm_2,daily_21_22pm_2,daily_05_06amand09_18pm_2,daily_22_05am_2,daily_rkvahr_imp,daily_rkvahr_exp,gf,fm,sch,unsch,manualstoppage,readingnotavailable,total,remarks) values("{x.get("genDate")}","{x.get("mw")}","{x.get("companyName")}","{x.get("htno")}","{x.get("site")}","{x.get("locNo")}","{x.get("reading_taken_time")}",{check_float_val(x.get("cml_run_hr"))},{check_float_val(x.get("cml_gen_hr"))},{check_float_val(x.get("cml_g_0"))},{check_float_val(x.get("cml_gen"))},{check_float_val(x.get("cml_total_prod"))},{check_float_val(x.get("cml_total_import"))},{check_float_val(x.get("cml_06_09_am_1"))},{check_float_val(x.get("cml_18_21_pm_1"))},{check_float_val(x.get("cml_21_22_pm_1"))},{check_float_val(x.get("cml_05_06_am_&_09_18_pm_1"))},{check_float_val(x.get("cml_22_05_am_1"))},{check_float_val(x.get("cml_total_export"))},{check_float_val(x.get("cml_06_09_am_2"))},{check_float_val(x.get("cml_18_21_pm_2"))},{check_float_val(x.get("cml_21_22_pm_2"))},{check_float_val(x.get("cml_05_06_am_&_09_18_pm_2"))},{check_float_val(x.get("cml_22_05_am_2"))},{check_float_val(x.get("cml_rkvahr_imp"))},{check_float_val(x.get("cml_rkvahr_exp"))},{check_float_val(x.get("daily_run_hr"))},{check_float_val(x.get("daily_gen_hr"))},{check_float_val(x.get("daily_g_0"))},{check_float_val(x.get("daily_gen"))},{check_float_val(x.get("Prod"))},{check_float_val(x.get("daily_total_import"))},{check_float_val(x.get("daily_06_09_am_1"))},{check_float_val(x.get("daily_18_21_pm_1"))},{check_float_val(x.get("daily_21_22_pm_1"))},{check_float_val(x.get("daily_05_06_am_&_09_18_pm_1"))},{check_float_val(x.get("daily_22_05_am_1"))},{check_float_val(x.get("daily_total_export"))},{check_float_val(x.get("daily_06_09_am_2"))},{check_float_val(x.get("daily_18_21_pm_2"))},{check_float_val(x.get("daily_21_22_pm_2"))},{check_float_val(x.get("daily_05_06_am_&_09_18_pm_2"))},{check_float_val(x.get("daily_22_05_am_2"))},{check_float_val(x.get("daily_rkvahr_imp"))},{check_float_val(x.get("daily_rkvahr_exp"))},{check_float_val(x.get("gf"))},{check_float_val(x.get("fm"))},{check_float_val(x.get("sch"))},{check_float_val(x.get("unsch"))},{check_float_val(x.get("ms"))},{check_float_val(x.get("readNotAvail"))},{check_float_val(x.get("total"))},"{x.get("remarks")}");'
                                        ebkwhValue = abs(float(check_float_val(x.get(
                                            "daily_total_export")))) - abs(float(check_float_val(x.get("daily_total_import"))))
                                        customerName = "SPI Power" if "spi" in x[2].lower() or "skr" in x[2].lower(
                                        ) else "KR Wind Energy" if "kr" in x[2].lower() else ''
                                        location_values = location.get(
                                            x.get('locNo'))
                                        db_command2 = f"INSERT into spi_windmill_gen_daily_report(gendate,companyname,locno,mckwhday,gf,fm,sch,unsch,genhrs,oprhrs,ebkwhday,mw,section,site,make,htno) values('{str(x.get('genDate')).split(' ')[0]}','{customerName}','{x.get('locNo')}',{float(check_float_val(x.get('Prod')))},{check_float_val(x.get('gf'))},{check_float_val(x.get('fm'))},{float(check_float_val(x.get('sch')))},{float(check_float_val(x.get('unsch')))},{float(check_float_val(x.get('daily_gen_hr')))},{float(check_float_val(x.get('daily_run_hr')))},{ebkwhValue},{float(check_float_val(x.get('mw')))},'{location_values[1]}','{x.get('site')}','{location_values[0]}','{location_values[3]}');"
                                        vestas_daily = check_valuein_reporting_layer(
                                            cursor, [str(x.get('genDate')).split(' ')[0], customerName, x.get('locNo')])
                                        # print(vestas_daily)
                                        if vestas_daily:
                                            if any([x.get('cml_run_hr'), x.get('cml_gen_hr'), x.get('cml_g_0'), x.get('cml_gen'), x.get("cml_total_prod")]):
                                                cursor.execute(db_command1)
                                                cursor.execute(db_command2)
                                                file_data.append({'gendate':str(x.get('genDate')).split(' ')[0],'mckwhday':float(check_float_val(x.get('Prod'))),'make':location_values[0]})
                                        # except Exception as dbe:
                                        #     logging.error(f"Error occured in row data insertion {dbe}")
                                logging.info(
                                    f"INFO :: Successfully inserted {sheet} sheet data into database of {file_name}")
                                logging.info(
                                    f"INFO :: Data in all the sheet from {file_name} is Successfully Inserted into vestas_xl_daily_hist and spi_windmill_gen_daily_report Database")
                                sending_mail(
                                    f"RAP Bot Successfull data uploaded notification for {customer_type}", f"Data from {file_name} is Successfully Inserted into  Database", "Admin")
                            except Exception as e:
                                dataBaseError.append(e)
                                logging.error(
                                    f"ERROR :: Error occured while inserting {sheet} sheet data from {file_name} file into database")
                                sending_mail(f"RAP Bot notification for error in Database insert",
                                             f"Data from {sheet} sheet data from {file_name} with {customer_type} type is not Inserted into  Database Error occured {e}", "Admin")
                    except Exception as e:
                        dataBaseError.append(e)
                        print(
                            "\n\Error occured while Inserting into vestas daily\n\n")
                        logging.error(
                            f"ERROR :: An error occured while inserting data from {file_name} into {e}")
                        sending_mail(f"RAP Bot notification for error in Database insert",
                                     f"Data from {file_name} or {customer_type} type is not Inserted into  Database Error occured {e}", "Admin")

                if 'suzlon_weekly' in customer_type:
                    excel_df = pd.read_excel(
                        file_path, sheet_name=None, header=None)
                    file_name = file_path.split('/')[-1]
                    logging.info(
                        "INFO ::-------- > Table used : suzlon_xl_weekly_hist and spi_windmill_gen_daily_report is get updated")
                    location_name = []
                    for sheetName in excel_df:
                        df = excel_df[sheetName].fillna('')
                        for x_i, x in df.iterrows():
                            if "date" in str(x[0]).lower():
                                df_header = df.iloc[x_i-1:x_i+1].fillna('')
                                break
                        df = df.iloc[x_i+1:]
                        head_val = ''
                        header = []
                        for x_i, x in df_header.iteritems():
                            if x.iloc[0]:
                                head_val = 'read_' if 'reading' in x.iloc[0].lower(
                                ) else 'calc_' if 'calculated' in x.iloc[0].lower() else ''
                            if 'date' in x.iloc[1].lower():
                                x.iloc[1] = 'genDate'
                            if x.iloc[1].lower() == 'mw' or x.iloc[1].lower() == 'site':
                                x.iloc[1] = x.iloc[1].lower()
                            if 'customer' in x.iloc[1].lower():
                                x.iloc[1] = 'companyName'
                            if 'htno' in re.sub(r"\s+", "", x.iloc[1].lower()):
                                x.iloc[1] = 'htno'
                            if 'locno' in re.sub(r"\s+", "", x.iloc[1].lower()):
                                x.iloc[1] = 'locno'
                            if 'total' in x.iloc[1].lower() and 'import' in x.iloc[1].lower():
                                x.iloc[1] = head_val+"total_import"
                            if 'total' in x.iloc[1].lower() and 'export' in x.iloc[1].lower():
                                x.iloc[1] = head_val+"total_export"
                            if '6.0am' in re.sub(r"\s+", "", x.iloc[1].lower()):
                                x.iloc[1] = head_val+"6am_to_9am_1" if head_val + \
                                    "6am_to_9am_1" not in df_header.iloc[1].values else head_val+"6am_to_9am_2"
                            if '6.0pm' in re.sub(r"\s+", "", x.iloc[1].lower()):
                                x.iloc[1] = head_val+"6pm_to_9pm_1" if head_val + \
                                    "6pm_to_9pm_1" not in df_header.iloc[1].values else head_val+"6pm_to_9pm_2"
                            if '9.0pm' in re.sub(r"\s+", "", x.iloc[1].lower()):
                                x.iloc[1] = head_val+"9pm_to_10pm_1" if head_val + \
                                    "9pm_to_10pm_1" not in df_header.iloc[1].values else head_val+"9pm_to_10pm_2"
                            if '5.0am' in re.sub(r"\s+", "", x.iloc[1].lower()):
                                x.iloc[1] = head_val+"5am_to_6am_and_9am_to_6pm_1" if head_val + \
                                    "5am_to_6am_and_9am_to_6pm_1" not in df_header.iloc[1].values else head_val+"5am_to_6am_and_9am_to_6pm_2"
                            if '10pm' in re.sub(r"\s+", "", x.iloc[1].lower()) and '5am' in re.sub(r"\s+", "", x.iloc[1].lower()):
                                x.iloc[1] = head_val+"10pm_to_5am_1" if head_val + \
                                    "10pm_to_5am_1" not in df_header.iloc[1].values else head_val+"10pm_to_5am_2"
                            if re.search(r"KVA(R|)H", x.iloc[1].upper().strip()):
                                if 'import' in x.iloc[1].lower() or 'export' in x.iloc[1].lower():
                                    if 'lag' in x.iloc[1].lower():
                                        x.iloc[1] = head_val+"kvarh_import_lag" if 'import' in x.iloc[1].lower(
                                        ) else head_val+"kvarh_export_lag"
                                    if 'lead' in x.iloc[1].lower():
                                        x.iloc[1] = head_val+"kvarh_import_lead" if 'import' in x.iloc[1].lower(
                                        ) else head_val+"kvarh_export_lead"
                                    if 'reading' in x.iloc[1].lower():
                                        x.iloc[1] = head_val+"kvah_import_reading" if 'import' in x.iloc[1].lower(
                                        ) else head_val+"kvah_export_reading"
                                    if "%" in x.iloc[1]:
                                        x.iloc[1] = head_val + \
                                            "percent_kvarh_import"
                            if 'month' in x.iloc[1].lower() and 'cumulative' in x.iloc[1].lower():
                                x.iloc[1] = head_val+"month_cml"
                            if 'power' in x.iloc[1].lower() and 'factor' in x.iloc[1].lower():
                                x.iloc[1] = head_val+"power_factor"
                        df.columns = df_header.iloc[1].unique()
                        try:
                            for data_i, data in df.iterrows():
                                if re.match(r"\d{4}-\d{2}-\d{2}\s\d{2}:\d{2}:\d{2}", str(data.get('genDate'))):
                                    db_command1 = f"INSERT INTO suzlon_xl_weekly_hist(gendate,mw,customername,htno,locno,reading_totalimport,reading_06_09am_1,reading_06_09pm_1,reading_09_10pm_1,reading_05_06amand09_06pm_1,reading_10pm_05am_1,reading_totalexport,reading_06_09am_2,reading_06_09pm_2,reading_09_10pm_2,reading_05_06amand09_06pm_2,reading_10pm_05am_2,reading_kvarhimportlag,reading_kvarhimportlead,reading_kvarhexportlag,reading_kvarhexportlead,reading_kvahimportreading,reading_kvahexportreading,reading_powerfactor,reading_percent_kvahimport,reading_monthcumulative,calc_totalimport,calc_06_09am_1,calc_06_09pm_1,calc_09_10pm_1,calc_05_06amand09_06pm_1,calc_10pm_05am_1,calc_totalexport,calc_06_09am_2,calc_06_09pm_2,calc_09_10pm_2,calc_05_06amand09_06pm_2,calc_10pm_05am_2,calc_kvarhimportlag,calc_kvarhimportlead,calc_kvarhexportlag,calc_kvarhexportlead,calc_kvahimportreading,calc_kvahexportreading,calc_powerfactor,calc_percent_kvahimport,calc_monthcumulative) values('{str(data.get('genDate')).split(' ')[0]}',{float(check_float_val(data.get('mw')))},'{str(data.get('companyName'))}','{str(data.get('htno'))}','{str(data.get('locno'))}',{float(check_float_val(data.get('read_total_import')))},{float(check_float_val(data.get('read_6am_to_9am_1')))},{float(check_float_val(data.get('read_6pm_to_9pm_1')))},{float(check_float_val(data.get('read_9pm_to_10pm_1')))},{float(check_float_val(data.get('read_5am_to_6am_and_9am_to_6pm_1')))},{float(check_float_val(data.get('read_10pm_to_5am_1')))},{float(check_float_val(data.get('read_total_export')))},{float(check_float_val(data.get('read_6am_to_9am_2')))},{float(check_float_val(data.get('read_6pm_to_9pm_2')))},{float(check_float_val(data.get('read_9pm_to_10pm_2')))},{float(check_float_val(data.get('read_5am_to_6am_and_9am_to_6pm_2')))},{float(check_float_val(data.get('read_10pm_to_5am_2')))},{float(check_float_val(data.get('read_kvarh_import_lag')))},{float(check_float_val(data.get('read_kvarh_import_lead')))},{float(check_float_val(data.get('read_kvarh_export_lag')))},{float(check_float_val(data.get('read_kvarh_export_lead')))},{float(check_float_val(data.get('read_kvah_import_reading')))},{float(check_float_val(data.get('read_kvah_export_reading')))},{float(check_float_val(data.get('read_power_factor')))},{float(check_float_val(data.get('read_percent_kvarh_import')))},{float(check_float_val(data.get('read_month_cml')))},{float(check_float_val(data.get('calc_total_import')))},{float(check_float_val(data.get('calc_6am_to_9am_1')))},{float(check_float_val(data.get('calc_6pm_to_9pm_1')))},{float(check_float_val(data.get('calc_9pm_to_10pm_1')))},{float(check_float_val(data.get('calc_5am_to_6am_and_9am_to_6pm_1')))},{float(check_float_val(data.get('calc_10pm_to_5am_1')))},{float(check_float_val(data.get('calc_total_export')))},{float(check_float_val(data.get('calc_6am_to_9am_2')))},{float(check_float_val(data.get('calc_6pm_to_9pm_2')))},{float(check_float_val(data.get('calc_9pm_to_10pm_2')))},{float(check_float_val(data.get('calc_5am_to_6am_and_9am_to_6pm_2')))},{float(check_float_val(data.get('calc_10pm_to_5am_2')))},{float(check_float_val(data.get('calc_kvarh_import_lag')))},{float(check_float_val(data.get('calc_kvarh_import_lead')))},{float(check_float_val(data.get('calc_kvarh_export_lag')))},{float(check_float_val(data.get('calc_kvarh_export_lead')))},{float(check_float_val(data.get('calc_kvah_import_reading')))},{float(check_float_val(data.get('calc_kvah_export_reading')))},{float(check_float_val(data.get('calc_power_factor')))},{float(check_float_val(data.get('calc_percent_kvarh_import')))},{float(check_float_val(data.get('calc_month_cml')))});"
                                    if any([str(data.get('companyName')).replace('NaT', ''), str(data.get('htno')).replace('NaT', ''), str(data.get('locno')).replace('NaT', '')]):
                                        customerName = "SPI Power" if "spi" in re.sub(r"\s+", '', data.get('companyName')).lower() or "skr" in re.sub(r"\s+", '', data.get(
                                            'companyName')).lower() else "KR Wind Energy" if "kr" in re.sub(r"\s+", '', data.get('companyName')).lower() else ''
                                        locNoVal = re.sub(
                                            r"\s+", '', data.get('locno')) if "TP06" not in data.get('locno') else "TP6"
                                        ebkwhday = abs(float(check_float_val(data.get('calc_total_export')))) - abs(
                                            float(check_float_val(data.get('calc_total_import'))))
                                        db_command2 = f"update spi_windmill_gen_daily_report set ebkwhday={float(check_float_val(ebkwhday))} where gendate='{str(data.get('genDate')).split(' ')[0]}' and locno='{locNoVal}' and companyname='{customerName}';"
                                        cursor.execute(db_command1)
                                        cursor.execute(db_command2)
                                        file_data.append({'gendate':str(data.get('genDate')).split(' ')[0],'ebkwhday':float(check_float_val(ebkwhday)),'make':'suzlon_weekly','locNoVal':locNoVal,'customerName':customerName,"FileName":file_name})

                            print("\n\nSuccessfully Inserted in suzlon weekly\n\n")
                            logging.info(
                                f"INFO ::Successfully inserted {sheetName} sheet data into database of {file_name}")
                            logging.info(
                                f"INFO ::Data from {file_name} is Successfully Inserted into suzlon_xl_weekly_hist and spi_windmill_gen_daily_report Database get updated")
                            location_name.append(sheetName)
                            # sending_mail(
                            #     f"RAP Bot Successfull data uploaded notification for {customer_type}", f"Data from {file_name} is and sheet name {sheetName} Successfully Inserted into  Database", "Admin")
                        except Exception as dbe:
                            dataBaseError.append(dbe)
                            logging.error(
                                f"ERROR :: Error occured while inserting {sheetName} sheet data from {file_name} file into database")
                            sending_mail(f"RAP Bot notification for error in Database insert",
                                         f"Data from {sheetName} sheet data from {file_name} with {customer_type} type is not Inserted into  Database Error occured {dbe}", "Admin")
                    sending_mail("RAP Bot Databases Successfull insertion",f"Bot has successfully inserted {customer_type} the data into the database and the locations are {', '.join(location_name)}","Admin")
            except Exception as e:
                dataBaseError.append(e)
                print("Error is : ", e)
                logging.error(f"ERROR :: Error occured in suzlon_weekly data insesrtion Error : {e}")
                sending_mail(
                    "RAP Bot Error notification", f"Error Occured suzlon_weekly in DB : {e} in {customer_type} and in the file {file_path}", "Admin")

            connection.commit()
            cursor.close()

        else:
            logging.error("ERROR :: Data base connection is not established")

            sending_mail("RAP Bot Error Notification","RAP BOT Have not estabished the connection with DB please do check the configurations","ADMIN")
        connection.close()
    except Exception as e:
        logging.error("ERROR :: Issue occured in read_excel() function")
        dataBaseError.append(e)
        print("The error is \t:", e)
        sending_mail("RAP Bot Error Notification","RAP BOT Have not estabished the connection with DB please do check the configurations","ADMIN")


def start_program(browser):
    browser.get(config["Website"]["url"])
    logging.info("INFO :: Bot run starts here")
    login_gmail(browser)


# Bot run starts here
# browser.implicitly_wait(30)


config = configparser.ConfigParser()
bot_run = True
bot_run_status = True
while(bot_run):
    config.read(os.path.join(os.path.dirname(__file__),'Config_file' ,'task.ini'))
    current_time = convert_time_zone(datetime.now()).replace(second=0, microsecond=0)
    bot_time = datetime.strptime(config['Bot']['schedule_time'], '%I:%M %p').replace(day=current_time.day, month=current_time.month, year=current_time.year, tzinfo=tz.gettz('Asia/Kolkata'))
    bot_run = int(config['Bot']['run'])
    if current_time == bot_time:
        os.makedirs(f'logs/{convert_time_zone(datetime.now()).strftime("%Y")}/{convert_time_zone(datetime.now()).strftime("%b")}',exist_ok=True)
        logging.basicConfig(filename=f'logs/{convert_time_zone(datetime.now()).strftime("%Y")}/{convert_time_zone(datetime.now()).strftime("%b")}/spi_log_{convert_time_zone(datetime.now()).date().day}.log',
                    format='%(asctime)s %(message)s',
                    filemode='a',
                    level=logging.DEBUG)
        dataBaseError=[]
        print(os.path.join(os.path.dirname(__file__),'chromedriver'))
        options = webdriver.ChromeOptions()
        suzlon_weekly_file = []
        download_file_path = os.path.join(os.path.dirname(
            __file__), config['Path']['download_path'])
        copy_file_path = os.path.join(os.path.dirname(
            __file__), config['Path']['copy_path'])
        os.makedirs(download_file_path, exist_ok=True)
        os.makedirs(copy_file_path, exist_ok=True)
        prefs = {"download.default_directory": download_file_path}
        options.add_experimental_option("prefs", prefs)
        options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.39 Safari/537.36")
        options.add_argument("--start-maximized")
        options.add_argument('--headless')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        browser = webdriver.Chrome(options=options)
        print("Path : ", download_file_path, "  ", copy_file_path)        
        sending_mail("RAP Bot started","RAP Bot started running","ADMIN")
        sending_mail("RAP Bot started","RAP Bot started running","Bussiness")
        logging.info("INFO :: Bot started to run")
        logging.info(
            'INFO :: Bot run time ---> {} and Current time ---> {}'.format(bot_time, current_time))
        file_data = []
        start_program(browser)
        print(bot_time)
        browser.quit()
        if dataBaseError:
            sending_mail("RAP Bot Partialy Completed","Some issue occured RAP Bot have partially inserted the data RAP Admin would look after the issue",'Bussiness')
        else:
            sending_mail("RAP Bot Successfully Completed","RAP Bot have inserted all the data Into the Database",'Bussiness')

        sending_mail("RAP Bot Completed","RAP Bot Completed","ADMIN")

logging.info("INFO :: BOT Stopped working")
sending_mail("RAP BOT Stopped working","RAP Bot Config file have been changed and bot STOPPED.",'ADMIN')