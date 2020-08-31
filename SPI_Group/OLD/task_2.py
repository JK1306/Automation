from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
import logging
import os
import configparser
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import shutil
from datetime import datetime
import time
from zipfile import ZipFile
from glob import glob
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
from openpyxl import Workbook
import openpyxl
import glob
import mysql.connector
from mysql.connector import Error
import json
import re
from selenium.webdriver.firefox.options import Options as FirefoxOptions

for handler in logging.root.handlers[:]:
    logging.root.removeHandler(handler)

logging.basicConfig(filename='task.log',
                    format='%(asctime)s %(message)s',
                    filemode='a',
                    level=logging.DEBUG)

config = configparser.ConfigParser()

config.read(os.path.join(os.path.dirname(__file__), 'task.ini'))

# Experimental

options = webdriver.ChromeOptions()

os.makedirs(config['Path']['download_path'],exist_ok=True)
os.makedirs(config['Path']['copy_path'],exist_ok=True)

prefs = {"download.default_directory": config['Path']['download_path']}
options.add_experimental_option("prefs", prefs)
options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.68 Safari/537.36")
#chromOpt.set_headless(headless=True)
#chromOpt.add_argument('--headless')
#options = webdriver.ChromeOptions()
options.add_argument('--headless')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
browser = webdriver.Chrome(options=options)
download_file_path = os.path.join(os.path.dirname(
    __file__), config['Path']['download_path'])
copy_file_path = os.path.join(os.path.dirname(
    __file__), config['Path']['copy_path'])
print("Path : ", download_file_path, "  ", copy_file_path)
print(config)


def last_mod_time(fname):
    folder_time = os.path.getmtime(fname)
    return os.path.getmtime(fname)


def move_zip_file(browser, customer_type):
    SECONDS_IN_DAY = 400
    # now= datetime.now().time().second
    now = time.time()
    before = now - SECONDS_IN_DAY
    copy_path = download_file_path
    #  logging.info(f"RAPBot has started moving the file to {output_path}")
    for file_name in os.listdir(copy_path):
        target_path = os.path.join(copy_path, file_name)
        if last_mod_time(target_path) > before:
            return file_name


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
    # browser.find_element_by_xpath('//input[@type="email" and @aria-label="Email or phone"]').send_keys(config['Login Details']['user_name'])
    WebDriverWait(browser, 45).until(EC.element_to_be_clickable(
        (By.XPATH, '//input[@type="email" and @aria-label="Email or phone"]'))).send_keys(config['Login Details']['user_name'])
    # WebDriverWait(browser, 45).until(EC.element_to_be_clickable((By.XPATH,'//input[@id="identifierId"]'))).send_keys(config['Login Details']['user_name'])
    print('Email ENtered')
    browser.find_element_by_xpath('//button[span="Next"]').click()
    WebDriverWait(browser, 45).until(EC.element_to_be_clickable(
        (By.XPATH, '//input[@type="password" and @aria-label="Enter your password"]'))).send_keys(config['Login Details']['password'])
    print("Password Entered")
    browser.implicitly_wait(10)
    WebDriverWait(browser, 30).until(EC.element_to_be_clickable(
        (By.XPATH, '//button[span="Next"]'))).click()
    browser.implicitly_wait(20)
    validate_mail(browser)


def download_button_click(browser, customer_type, exception_case_file=None):
    browser.implicitly_wait(30)
    ele_len = len(browser.find_elements_by_xpath(
        f'//div[@class="aQH"]/span[@download_url]'))
    print("Element Length : ", ele_len)
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
                    time.sleep(10)
                    move_downloaded_file(browser, customer_type, filename) if not exception_case_file else move_downloaded_file(
                        browser, customer_type, filename, True)
    else:
        pass
        # send mail for no document attached in mail
        # sending_mail()


def validate_mail(browser):
    count = 0
    WebDriverWait(browser, 30).until(EC.element_to_be_clickable(
        (By.XPATH, '//tr[contains(@class,"zA")]')))
    element_len = browser.find_elements_by_xpath(
        f'//tr[contains(@class,"zA")]')
    subject_val = {}
    file_path = []
    mail_tracker = {}
    suzlonCheckFilePath = os.path.dirname(__file__)+"/suzlonCheck.json"
    today = datetime.now()
    for x in config['Subject']:
        for y in config['Subject'][x].split(','):
            subject_val[y] = x
    exception_flag = [True for x in config['Exception']
                      if config['Exception'][x] == "ON"]
    if not exception_flag:
        for ele in range(1, len(element_len)+1):
            # browser.implicitly_wait(30)
            # WebDriverWait(browser,30).until(EC.((By.XPATH,f'//tr[contains(@class,"zA")][{ele}]/td[5]/div')))
            content_page = browser.find_elements_by_xpath(
                '//*[@id=":1"]/div/div[2]/div/table/tr/td[1]/div[2]')
            while len(content_page) != 0:
                content_page = browser.find_elements_by_xpath(
                    '//*[@id=":1"]/div/div[2]/div/table/tr/td[1]/div[2]')
            element = browser.find_element_by_xpath(
                f'//tr[contains(@class,"zA")][{ele}]/td[5]/div')
            mail_check_elemt = browser.find_element_by_xpath(
                f'//tr[contains(@class,"zA")][{ele}]/td[4]/div[1]/span/span').get_attribute('email')
            time_check = browser.find_element_by_xpath(
                f'//tr[contains(@class,"zA")][{ele}]/td[8]/span').get_attribute('title')
            mail_id = [config['Mail'][x] for x in config['Mail']]
            is_customer_mail = False
            current_date = datetime.now()
            mail_recived_time = datetime.strptime(
                time_check, '%a, %b %d, %Y, %I:%M %p')
            subject_check = browser.find_element_by_xpath(
                f'//tr[contains(@class,"zA")][{ele}]/td[5]/div[1]/div[1]/div[1]/span/span').text
            print(mail_check_elemt, " : ", time_check, " : ", subject_check)
            customer_type = ''
            for x in subject_val:
                if x in subject_check:
                    is_customer_mail = True
                    # customer_type = 'suzlon' if "suzlon" in subject_val[x] else 'vestas'
                    customer_type = subject_val[x]
                    break
            print(customer_type)
            if is_customer_mail and current_date.date().day == mail_recived_time.date().day:
                element.click()
                mail_element = browser.find_elements_by_xpath(
                    f'//div[@class="aQH"]/span[@download_url]')
                vestas_limit_end_time = datetime.strptime(
                    config["Mail Time"]['vestas_end_time'], '%I:%M %p')
                vestas_limit_end_time = vestas_limit_end_time.replace(
                    day=datetime.now().day, month=datetime.now().month, year=datetime.now().year)
                vestas_limit_start_time = datetime.strptime(
                    config["Mail Time"]['vestas_start_time'], '%I:%M %p')
                vestas_limit_start_time = vestas_limit_start_time.replace(
                    day=datetime.now().day, month=datetime.now().month, year=datetime.now().year)

                suzlon_limit_end_time = datetime.strptime(
                    config["Mail Time"]['suzlon_end_time'], '%I:%M %p')
                suzlon_limit_end_time = suzlon_limit_end_time.replace(
                    day=datetime.now().day, month=datetime.now().month, year=datetime.now().year)
                suzlon_limit_start_time = datetime.strptime(
                    config["Mail Time"]['suzlon_start_time'], '%I:%M %p')
                suzlon_limit_start_time = suzlon_limit_start_time.replace(
                    day=datetime.now().day, month=datetime.now().month, year=datetime.now().year)

                # suzlon daily download
                if "suzlon" in customer_type and "daily" in subject_check.lower():
                    if mail_check_elemt and mail_check_elemt in mail_id and suzlon_limit_end_time > mail_recived_time and suzlon_limit_start_time < mail_recived_time:
                        download_button_click(browser, customer_type)
                        # if customer_type not in mail_tracker:
                        #     mail_tracker[customer_type] = [subject_check]
                        # else:
                        #     mail_tracker[customer_type].append(subject_check)
                        # if len(mail_tracker[customer_type]) <= 4:
                        #     download_button_click(browser,customer_type)
                        # else:
                        #     print("excess suzlon daily mail")
                        #     sending_mail("RAP Bot notification","excess suzlone daily mail")
                        #     # send mail regarding excess mail in suzlon daily
                        #     pass

                # suzlon weekly download
                elif "suzlon" in customer_type and "weekly" in subject_check.lower():
                    download_button_click(browser, customer_type)
                    # suzlonWeekData = {"Download" : 1,"LastDownload":today.strftime("%d/%m/%Y %I:%M %p")}
                    # if os.path.exists(suzlonCheckFilePath):
                    #     with open(suzlonCheckFilePath,"r") as suzlonVal:
                    #         suzlonWeekDataRead = json.loads(suzlonVal.read())
                    #     if int(suzlonWeekDataRead['Download']) < 2 and today.isocalendar()[1] == datetime.strptime(suzlonWeekDataRead["LastDownload"],"%d/%m/%Y %I:%M %p").isocalendar()[1]:
                    #         suzlonWeekDataRead['Download'] += 1
                    #         suzlonWeekData = suzlonWeekDataRead
                    #         download_button_click(browser,customer_type)
                    #     elif today.isocalendar()[1] != datetime.strptime(suzlonWeekDataRead["LastDownload"],"%d/%m/%Y %I:%M %p").isocalendar()[1]:
                    #         download_button_click(browser,customer_type)
                    #     elif today.isocalendar()[1] == datetime.strptime(suzlonWeekDataRead["LastDownload"],"%d/%m/%Y %I:%M %p").isocalendar()[1]:
                    #         suzlonWeekData = suzlonWeekDataRead
                    #         print("excess suzlon weekly email")
                    #         sending_mail("RAP Bot notification","excess suzlone weekly mail")

                    #         # send email for excess weekly data sent
                    # else:
                    #     download_button_click(browser,customer_type)
                    # with open(suzlonCheckFilePath,"w+") as suzlonVal:
                    #     suzlonVal.write(json.dumps(suzlonWeekData))

                # vestas daily download
                elif "vestas" in customer_type:
                    if mail_recived_time < vestas_limit_end_time and mail_recived_time > vestas_limit_start_time:
                        download_button_click(browser, customer_type)
                        # print("------------------",mail_tracker[customer_type],"\n")
                        # mail_tracker[customer_type] = 1 if customer_type not in mail_tracker else mail_tracker[customer_type]+1
                        # if mail_tracker[customer_type] <=2:
                        #     download_button_click(browser,customer_type)
                        # else:
                        #     print("Excess vestas daily mail")
                        #     sending_mail("RAP Bot notification","excess vestas daily mail")
                        #     # send email for excess vestas daily mail
                        #     pass

                # email click back button
                email_back = ActionChains(browser)
                back_button = browser.find_element_by_xpath(
                    '//*[@id=":4"]/div[2]/div[1]/div/div[1]/div')
                email_back.move_to_element(back_button)
                browser.find_element_by_xpath(
                    '//div[@title="Back to Inbox" and @role="button"]').click()
            elif current_date.date().day != mail_recived_time.date().day:
                break

        # check all the mail are sent properly and notify admin in case of error
        # send_error_mail(browser,mail_tracker,today,suzlonCheckFilePath)
    else:
        exception_case(browser)

    # excel_file_path = extract_file(browser,file_saved_path)
    # read_excel_file(browser,excel_file_path)

        # except Exception as e:
        #     print(e)
        #     print("Except part")
        #     break


def send_error_mail(browser, mail_tracker, today, suzlonCheckFilePath):
    for x in mail_tracker:
        if 'suzlon' in x:
            if len(mail_tracker[x]) < int(config['No. of Mails']['suzlon_daily']):
                # send mail for less no of emails sent for suzlon daily
                spi_daily_count = 0
                kr_daily_count = 0
                for y in mail_tracker[x]:
                    if 'spi' in y.lower() or 'skr' in y.lower():
                        spi_daily_count += 1
                    elif 'k r' in y.lower():
                        kr_daily_count += 1
                if spi_daily_count < 2:
                    print("SPI mail is not yet sent")
                    sending_mail("RAP Bot notification",
                                 "SPI mail is not yet sent")
                    # send email for spi daily mail is not been sent
                    pass
                if kr_daily_count < 2:
                    print("KR suzlone mail is not yet sent")
                    sending_mail("RAP Bot notification",
                                 "KR suzlone mail is not yet sent")
                    # send email for kr daily mail is not been sent
                    pass
        elif 'vestas' in x:
            if mail_tracker[x] < int(config['No. of Mails']['vestas_daily']):
                print("Vestas daily mail is not yet sent")
                sending_mail("RAP Bot notification",
                             "Vestas daily mail is not yet sent")
                # send mail for less no of emails sent for vestas daily
                pass
    if today.strftime('%a') == 'Fri':
        if os.path.exists(suzlonCheckFilePath):
            with open(suzlonCheckFilePath, "r") as suzlonVal:
                suzlonWeekDataRead = json.loads(suzlonVal.read())
            if int(suzlonWeekDataRead['Download']) < 2:
                # send mail notif. for suzlon weekly not been sent
                sending_mail("RAP Bot notification",
                             "Suzlon weekly mail is not yet sent")
                pass


def exception_case(browser):
    exception_customer = [x for x in config['Exception']
                          if config['Exception'][x].upper() == "ON"]
    element_len = browser.find_elements_by_xpath(
        f'//tr[contains(@class,"zA")]')
    for email_index in range(1, len(element_len)+1):
        email_element = browser.find_element_by_xpath(
            f'//tr[contains(@class,"zA")][{email_index}]/td[5]/div')
        for company in exception_customer:
            mail_subject = config['Subject'][company]
            # mail_time = datetime.strptime(config['Exception Config'][company],'%d/%m/%Y %I:%M %p')
            mail_time = config['Exception Config'][company].split('-')
            mail_time = [datetime.strptime(
                x.strip(), '%d/%m/%Y %I:%M %p') for x in mail_time]
            mail_check_elemt = browser.find_element_by_xpath(
                f'//tr[contains(@class,"zA")][{email_index}]/td[4]/div[1]/span/span').get_attribute('email')
            time_check = browser.find_element_by_xpath(
                f'//tr[contains(@class,"zA")][{email_index}]/td[8]/span').get_attribute('title')
            subject_check = browser.find_element_by_xpath(
                f'//tr[contains(@class,"zA")][{email_index}]/td[5]/div[1]/div[1]/div[1]/span/span').text
            time_check = datetime.strptime(
                time_check, '%a, %b %d, %Y, %I:%M %p')
            # if mail_subject in subject_check and time_check == mail_time:
            if mail_subject in subject_check:
                if mail_time[0] <= time_check and time_check >= mail_time[1]:
                    email_element.click()
                    download_button_click(browser, company, True)


def extract_file(browser, excel_dest_path):
    # print("REACHED extract file")
    file_name = set()
    file_list = os.listdir(excel_dest_path)
    ret_file_name = set()
    for file in file_list:
        # print(file)
        if os.path.isdir(os.path.join(excel_dest_path, file)):
            ret_file_name = extract_file(
                browser, os.path.join(excel_dest_path, file))
        else:
            if 'zip' in file:
                zip_file_path = os.path.join(excel_dest_path, file)
                with ZipFile(zip_file_path, "r") as zip_file:
                    zip_file.extractall(excel_dest_path)
                os.remove(zip_file_path)
                ret_file_name = extract_file(browser, excel_dest_path)
            else:
                file_name.add(os.path.join(excel_dest_path, file))
        file_name.update(ret_file_name)
    return file_name


def move_downloaded_file(browser, customer_type, file_name, exception=None):
    dow_path = os.path.join(download_file_path, file_name)
    file_date = datetime.now().strftime("%d")
    file_month = datetime.now().strftime("%m")
    print(copy_file_path)
    if not exception:
        des_path = os.path.join(
            copy_file_path, "SPI\\{}\\{}".format(file_month, file_date))
    else:
        des_path = os.path.join(copy_file_path, "SPI\\{}\\{}\\{}".format(
            file_month, file_date, "Exception_run"))
    os.makedirs(des_path+"\\{}".format(customer_type), exist_ok=True)
    try:
        shutil.move(dow_path, des_path+"\\{}".format(customer_type))
    except Exception as e:
        logging.info(e)
    print("Destination Path: ", os.path.join(
        des_path, customer_type, file_name), "   ", customer_type)
    excel_file_path = os.path.join(des_path, customer_type, file_name)
    print(excel_file_path)
    read_excel_file(browser, excel_file_path, customer_type)
    # return des_path


def read_excel_file(browser, file_path, customer_type):
    def read_location_master(cursor):
        cursor.execute("select * from location_master;")
        location_data = cursor.fetchall()
        location_dic = {}
        for x in location_data:
            location_dic[x[0]] = [x[3], x[4], x[5]]
        return location_dic

    def check_float_val(data):
        try:
            float(data)
            return data if float(data) else 0.0
        except:
            return 0.0
    # try:
    connection = mysql.connector.connect(host=config["DB Config"]["host"],
                                         port=config["DB Config"]["port"],
                                         database=config["DB Config"]["database"],
                                         user=config["DB Config"]["user_name"],
                                         password=config["DB Config"]["paswd"])
    if connection.is_connected():
        cursor = connection.cursor()
        location = read_location_master(cursor)
        try:
            if "suzlon_daily" in customer_type:
                sheet_val = pd.read_excel(file_path, sheet_name=None)
                for sheet_name in sheet_val:
                    if "generation" in sheet_name.lower():
                        doc_val = sheet_val[sheet_name].fillna('')
                        for y in doc_val.columns:
                            if "date" in y.lower():
                                doc_val.rename(
                                    columns={y: 'genDate'}, inplace=True)
                            if "customer" in y.lower() or 'company' in y.lower():
                                doc_val.rename(
                                    columns={y: 'customerName'}, inplace=True)
                            if "state" in y.lower() or "site" in y.lower() or "section" in y.lower() or y.lower() == "mw" or y.lower() == "gf" or y.lower() == "fm" or y.lower() == "s" or y.lower() == "u" or y.lower() == "nor":
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
                        for column_val in doc_val.iterrows():
                            x = column_val[1]
                            if re.match(r"\d{4}-\d{2}-\d{2}\s\d{2}:\d{2}:\d{2}", str(x[0])):
                                db_command = f"insert into suzlon_xl_daily_hist(gendate,customername,state,site,section,mw,locno,genkwhday,genkwhmtd,genkwhytd,plfday,plfmtd,plfytd,mcavail,gf,fm,s,u,nor,genhrs,oprhrs) values('{str(x.get('genDate')).split(' ')[0]}','{x.get('customerName')}','{x.get('state')}','{x.get('site')}','{x.get('section')}',{float(check_float_val(x.get('mw')))},'{x.get('locNo')}',{float(check_float_val(x.get('genkwhDay')))},{float(check_float_val(x.get('genkwhMtd')))},{float(check_float_val(x.get('genkwhYtd')))},{float(check_float_val(x.get('plfDay')))},{float(check_float_val(x.get('plfMtd')))},{float(check_float_val(x.get('plfYtd')))},{float(check_float_val(x.get('mcAvail')))},{float(check_float_val(x.get('gf')))},{float(check_float_val(x.get('fm')))},{float(check_float_val(x.get('s')))},{float(check_float_val(x.get('u')))},{float(check_float_val(x.get('nor')))},{float(check_float_val(x.get('genHrs')))},{float(check_float_val(x.get('oprHrs')))});"
                                customerName = "SPI Power" if "spi" in re.sub(r"\s+", '', x.get('customerName')).lower() or "skr" in re.sub(r"\s+", '', x.get(
                                    'customerName')).lower() else "KR Wind Energy" if "kr" in re.sub(r"\s+", '', x.get('customerName')).lower() else ''
                                locNoVal = re.sub(
                                    r"\s+", '', x.get('locNo')) if "TP06" not in x.get('locNo') else "TP6"
                                location_values = location.get(locNoVal)
                                db_command2 = f"insert into spi_windmill_gen_daily_report(gendate,companyname,locno,mckwhday,gf,fm,sch,unsch,genhrs,oprhrs,mw,section,site,make) values('{str(x.get('genDate')).split(' ')[0]}','{customerName}','{locNoVal}',{float(check_float_val(x.get('genkwhDay')))},{float(check_float_val(x.get('gf')))},{float(check_float_val(x.get('fm')))},{float(check_float_val(x.get('s')))},{float(check_float_val(x.get('u')))},{float(check_float_val(x.get('genHrs')))},{float(check_float_val(x.get('oprHrs')))},{float(check_float_val(x.get('mw')))},'{x.get('section')}','{x.get('site')}','{location_values[0]}');"
                                cursor.execute(db_command)
                                cursor.execute(db_command2)
                        print("\n\nSuccessfully Inserted in suzlon daily\n\n")

            if "vestas_daily" in customer_type:
                df_dic = pd.read_excel(file_path, sheet_name=None, header=None)
                df_header = ""
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
                                x.iloc[1] = head_val+'05_06_am_&_09_18_pm_2'
                            else:
                                x.iloc[1] = head_val+'05_06_am_&_09_18_pm_1'
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

                    df.columns = df_header.iloc[1]
                    for column_val in df.iterrows():
                        x = column_val[1]
                        if re.match(r"\d{4}-\d{2}-\d{2}\s\d{2}:\d{2}:\d{2}", str(x[0])):
                            db_command1 = f'INSERT into vestas_xl_daily_hist(gendate,mw,customername,htno,site,locno,readingtakentime,cml_runhrs,cml_genhrs,cml_g0,cml_gen,cml_totalprod,cml_totalimport,cml_06_09am_1,cml_18_21pm_1,cml_21_22pm_1,cml_05_06amand09_18pm_1,cml_22_05am_1,cml_totalexport,cml_06_09am_2,cml_18_21pm_2,cml_21_22pm_2,cml_05_06amand09_18pm_2,cml_22_05am_2,cml_rkvahr_imp,cml_rkvahr_exp,daily_runhrs,daily_genhrs,daily_g0,daily_gen,daily_totalprod,daily_totalimport,daily_06_09am_1,daily_18_21pm_1,daily_21_22pm_1,daily_05_06amand09_18pm_1,daily_22_05am_1,daily_totalexport,daily_06_09am_2,daily_18_21pm_2,daily_21_22pm_2,daily_05_06amand09_18pm_2,daily_22_05am_2,daily_rkvahr_imp,daily_rkvahr_exp,gf,fm,sch,unsch,manualstoppage,readingnotavailable,total,remarks) values("{x.get("genDate")}","{x.get("mw")}","{x.get("companyName")}","{x.get("htno")}","{x.get("site")}","{x.get("locNo")}","{x.get("reading_taken_time")}",{check_float_val(x.get("cml_run_hr"))},{check_float_val(x.get("cml_gen_hr"))},{check_float_val(x.get("cml_g_0"))},{check_float_val(x.get("cml_gen"))},{check_float_val(x.get("cml_total_prod"))},{check_float_val(x.get("cml_total_import"))},{check_float_val(x.get("cml_06_09_am_1"))},{check_float_val(x.get("cml_18_21_pm_1"))},{check_float_val(x.get("cml_21_22_pm_1"))},{check_float_val(x.get("cml_05_06_am_&_09_18_pm_1"))},{check_float_val(x.get("cml_22_05_am_1"))},{check_float_val(x.get("cml_total_export"))},{check_float_val(x.get("cml_06_09_am_2"))},{check_float_val(x.get("cml_18_21_pm_2"))},{check_float_val(x.get("cml_21_22_pm_2"))},{check_float_val(x.get("cml_05_06_am_&_09_18_pm_2"))},{check_float_val(x.get("cml_22_05_am_2"))},{check_float_val(x.get("cml_rkvahr_imp"))},{check_float_val(x.get("cml_rkvahr_exp"))},{check_float_val(x.get("daily_run_hr"))},{check_float_val(x.get("daily_gen_hr"))},{check_float_val(x.get("daily_g_0"))},{check_float_val(x.get("daily_gen"))},{check_float_val(x.get("Prod"))},{check_float_val(x.get("daily_total_import"))},{check_float_val(x.get("daily_06_09_am_1"))},{check_float_val(x.get("daily_18_21_pm_1"))},{check_float_val(x.get("daily_21_22_pm_1"))},{check_float_val(x.get("daily_05_06_am_&_09_18_pm_1"))},{check_float_val(x.get("daily_22_05_am_1"))},{check_float_val(x.get("daily_total_export"))},{check_float_val(x.get("daily_06_09_am_2"))},{check_float_val(x.get("daily_18_21_pm_2"))},{check_float_val(x.get("daily_21_22_pm_2"))},{check_float_val(x.get("daily_05_06_am_&_09_18_pm_2"))},{check_float_val(x.get("daily_22_05_am_2"))},{check_float_val(x.get("daily_rkvahr_imp"))},{check_float_val(x.get("daily_rkvahr_exp"))},{check_float_val(x.get("gf"))},{check_float_val(x.get("fm"))},{check_float_val(x.get("sch"))},{check_float_val(x.get("unsch"))},{check_float_val(x.get("ms"))},{check_float_val(x.get("readNotAvail"))},{check_float_val(x.get("total"))},"{x.get("remarks")}");'
                            ebkwhValue = abs(float(check_float_val(x.get(
                                "daily_total_export")))) - abs(float(check_float_val(x.get("daily_total_import"))))
                            customerName = "SPI Power" if "spi" in x[2].lower() or "skr" in x[2].lower(
                            ) else "KR Wind Energy" if "kr" in x[2].lower() else ''
                            location_values = location.get(x.get('locNo'))
                            db_command2 = f"INSERT into spi_windmill_gen_daily_report(gendate,companyname,locno,mckwhday,gf,fm,sch,unsch,genhrs,oprhrs,ebkwhday,mw,section,site,make) values('{str(x.get('genDate')).split(' ')[0]}','{customerName}','{x.get('locNo')}',{float(check_float_val(x.get('Prod')))},{check_float_val(x.get('gf'))},{check_float_val(x.get('fm'))},{float(check_float_val(x.get('sch')))},{float(check_float_val(x.get('unsch')))},{float(check_float_val(x.get('daily_gen_hr')))},{float(check_float_val(x.get('daily_run_hr')))},{ebkwhValue},{float(check_float_val(x.get('mw')))},'{location_values[1]}','{x.get('site')}','{location_values[0]}');"
                            if any([x.get('cml_run_hr'), x.get('cml_gen_hr'), x.get('cml_g_0'), x.get('cml_gen'), x.get("cml_total_prod")]):
                                # try:
                                cursor.execute(db_command1)
                                cursor.execute(db_command2)
                    print("\n\nSuccessfully Inserted in vestas daily\n\n")

            if 'suzlon_weekly' in customer_type:
                excel_df = pd.read_excel(
                    file_path, sheet_name=None, header=None)
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
                                    x.iloc[1] = head_val+"percent_kvarh_import"
                        if 'month' in x.iloc[1].lower() and 'cumulative' in x.iloc[1].lower():
                            x.iloc[1] = head_val+"month_cml"
                        if 'power' in x.iloc[1].lower() and 'factor' in x.iloc[1].lower():
                            x.iloc[1] = head_val+"power_factor"
                    df.columns = df_header.iloc[1].unique()
                    for data_i, data in df.iterrows():
                        db_command1 = f"INSERT INTO suzlon_xl_weekly_hist(gendate,mw,customername,htno,locno,reading_totalimport,reading_06_09am_1,reading_06_09pm_1,reading_09_10pm_1,reading_05_06amand09_06pm_1,reading_10pm_05am_1,reading_totalexport,reading_06_09am_2,reading_06_09pm_2,reading_09_10pm_2,reading_05_06amand09_06pm_2,reading_10pm_05am_2,reading_kvarhimportlag,reading_kvarhimportlead,reading_kvarhexportlag,reading_kvarhexportlead,reading_kvahimportreading,reading_kvahexportreading,reading_powerfactor,reading_percent_kvahimport,reading_monthcumulative,calc_totalimport,calc_06_09am_1,calc_06_09pm_1,calc_09_10pm_1,calc_05_06amand09_06pm_1,calc_10pm_05am_1,calc_totalexport,calc_06_09am_2,calc_06_09pm_2,calc_09_10pm_2,calc_05_06amand09_06pm_2,calc_10pm_05am_2,calc_kvarhimportlag,calc_kvarhimportlead,calc_kvarhexportlag,calc_kvarhexportlead,calc_kvahimportreading,calc_kvahexportreading,calc_powerfactor,calc_percent_kvahimport,calc_monthcumulative) values('{str(data.get('genDate')).split(' ')[0]}',{float(check_float_val(data.get('mw')))},'{str(data.get('companyName'))}','{str(data.get('htno'))}','{str(data.get('locno'))}',{float(check_float_val(data.get('read_total_import')))},{float(check_float_val(data.get('read_6am_to_9am_1')))},{float(check_float_val(data.get('read_6pm_to_9pm_1')))},{float(check_float_val(data.get('read_9pm_to_10pm_1')))},{float(check_float_val(data.get('read_5am_to_6am_and_9am_to_6pm_1')))},{float(check_float_val(data.get('read_10pm_to_5am_1')))},{float(check_float_val(data.get('read_total_export')))},{float(check_float_val(data.get('read_6am_to_9am_2')))},{float(check_float_val(data.get('read_6pm_to_9pm_2')))},{float(check_float_val(data.get('read_9pm_to_10pm_2')))},{float(check_float_val(data.get('read_5am_to_6am_and_9am_to_6pm_2')))},{float(check_float_val(data.get('read_10pm_to_5am_2')))},{float(check_float_val(data.get('read_kvarh_import_lag')))},{float(check_float_val(data.get('read_kvarh_import_lead')))},{float(check_float_val(data.get('read_kvarh_export_lag')))},{float(check_float_val(data.get('read_kvarh_export_lead')))},{float(check_float_val(data.get('read_kvah_import_reading')))},{float(check_float_val(data.get('read_kvah_export_reading')))},{float(check_float_val(data.get('read_power_factor')))},{float(check_float_val(data.get('read_percent_kvarh_import')))},{float(check_float_val(data.get('read_month_cml')))},{float(check_float_val(data.get('calc_total_import')))},{float(check_float_val(data.get('calc_6am_to_9am_1')))},{float(check_float_val(data.get('calc_6pm_to_9pm_1')))},{float(check_float_val(data.get('calc_9pm_to_10pm_1')))},{float(check_float_val(data.get('calc_5am_to_6am_and_9am_to_6pm_1')))},{float(check_float_val(data.get('calc_10pm_to_5am_1')))},{float(check_float_val(data.get('calc_total_export')))},{float(check_float_val(data.get('calc_6am_to_9am_2')))},{float(check_float_val(data.get('calc_6pm_to_9pm_2')))},{float(check_float_val(data.get('calc_9pm_to_10pm_2')))},{float(check_float_val(data.get('calc_5am_to_6am_and_9am_to_6pm_2')))},{float(check_float_val(data.get('calc_10pm_to_5am_2')))},{float(check_float_val(data.get('calc_kvarh_import_lag')))},{float(check_float_val(data.get('calc_kvarh_import_lead')))},{float(check_float_val(data.get('calc_kvarh_export_lag')))},{float(check_float_val(data.get('calc_kvarh_export_lead')))},{float(check_float_val(data.get('calc_kvah_import_reading')))},{float(check_float_val(data.get('calc_kvah_export_reading')))},{float(check_float_val(data.get('calc_power_factor')))},{float(check_float_val(data.get('calc_percent_kvarh_import')))},{float(check_float_val(data.get('calc_month_cml')))});"
                        customerName = "SPI Power" if "spi" in re.sub(r"\s+", '', x[2]).lower() or "skr" in re.sub(
                            r"\s+", '', x[2]).lower() else "KR Wind Energy" if "kr" in re.sub(r"\s+", '', x[2]).lower() else ''
                        locNoVal = re.sub(
                            r"\s+", '', x[4]) if "TP06" not in x[4] else "TP6"
                        ebkwhday = abs(float(check_float_val(data.get('calc_total_export')))) - abs(
                            float(check_float_val(data.get('calc_total_import'))))
                        db_command2 = f"update spi_windmill_gen_daily_report set ebkwhday={float(check_float_val(ebkwhday))} where gendate='{str(data.get('genDate')).split(' ')[0]}' and locno='{locNoVal}' and customername='{customerName}';"
                        cursor.execute(db_command1)
                        cursor.execute(db_command2)
                    print("\n\nSuccessfully Inserted in suzlon weekly\n\n")

            sending_mail("RAP Bot Databases Successfull insertion",
                         f"Bot has successfully inserted {customer_type} the data into the database", "Bussiness")
        except Exception as e:
            print("Error is : ", e)
            sending_mail("RAP Bot notification",
                         f"Error Occured in DB : {e}", "Admin")
        connection.commit()

        cursor.close()
    # except Exception as e:
    #     print("The error is \t:",e)


browser.get(config["Website"]["url"])
# browser.implicitly_wait(30)
# login_gmail(browser)
# browser.quit()
