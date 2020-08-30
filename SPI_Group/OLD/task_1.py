from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from pandas import DataFrame
import logging,os
import configparser
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
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

for handler in logging.root.handlers[:]:
    logging.root.removeHandler(handler)

logging.basicConfig(filename='task.log',
                    format='%(asctime)s %(message)s',
                    filemode='a',
                    level = logging.DEBUG)

import configparser
config = configparser.ConfigParser()

config.read(os.path.dirname(__file__)+'/task.ini')

# Experimental
chromOpt = webdriver.ChromeOptions()
prefs = {"download.default_directory" : r"C:\Users\Jaikishore\Downloads\SPI"}
chromOpt.add_experimental_option("prefs",prefs)
browser= webdriver.Chrome(os.path.dirname(__file__)+"/../chromedriver.exe",chrome_options=chromOpt)

download_file_path = config['Path']['download_path']
copy_file_path = config['Path']['copy_path']

def last_mod_time(fname):
    folder_time= os.path.getmtime(fname)
    return os.path.getmtime(fname)

def move_zip_file(browser,customer_type):
    SECONDS_IN_DAY = 400
    # now= datetime.now().time().second
    now = time.time()
    before = now - SECONDS_IN_DAY
    copy_path=download_file_path
    #  logging.info(f"RAPBot has started moving the file to {output_path}")
    for file_name in os.listdir(copy_path):
        target_path = os.path.join(copy_path, file_name)
        if last_mod_time(target_path) > before:
            return file_name

def sending_mail(subject,body_mes):
    # outlook_id = "jaikishore.gopalakrishnan@rap.ventures"
    msg = MIMEMultipart()
    msg['From'] = config["Login Details"]["user_name"] # from address
    msg['To'] = config["Mail"]["vestas_mail"] # to address
    # msg['Subject'] = "Choice Reports RAPBot notification"
    msg['Subject'] = subject
    body = f'{body_mes}'
    msg.attach(MIMEText(body, 'plain'))
    server = smtplib.SMTP('smtp.office365.com', '587')  ### put your relevant SMTP here
    server.ehlo()
    server.starttls()
    server.ehlo()
    # server.login(your mail id, your password)  ### if applicable
    server.login(config["Login Details"]["user_name"],config["Login Details"]["password"])
    server.send_message(msg)
    server.quit()

def login_gmail(browser):
    browser.find_element_by_xpath('//input[@type="email" and @aria-label="Email or phone"]').send_keys(config['Login Details']['user_name'])
    browser.find_element_by_xpath('//button[span="Next"]').click()
    # browser.implicitly_wait(30)
    WebDriverWait(browser, 45).until(EC.element_to_be_clickable((By.XPATH,'//input[@type="password" and @aria-label="Enter your password"]'))).send_keys(config['Login Details']['password'])
    # browser.find_element_by_xpath('//input[@type="password" and @aria-label="Enter your password"]').send_keys(config['Login Details']['password'])
    browser.implicitly_wait(10)
    WebDriverWait(browser, 30).until(EC.element_to_be_clickable((By.XPATH,'//button[span="Next"]'))).click()
    # browser.find_element_by_xpath('//button[span="Next"]').click()
    browser.implicitly_wait(30)
    download_attached_document(browser)

def download_button_click(browser):
    ele_len = len(browser.find_elements_by_xpath(f'//div[@aria-label="attachments"]/div'))
    if ele_len ==1:
        action = ActionChains(browser)
        file_element = browser.find_element_by_xpath(f'//div[@aria-label="attachments"]/div')
        action.move_to_element(file_element).perform()
        browser.find_element_by_xpath(f'//i[@data-icon-name="Download"]').click()
        filename = browser.find_element_by_xpath(f'//div[@aria-label="attachments"]/div/div/div/div[2]/div/div[1]').get_attribute('title')
        print(filename)
        time.sleep(20)
        destination_path = move_downloaded_file(browser,customer_type,filename)
    else:
        try:
            WebDriverWait(browser, 25).until(EC.element_to_be_clickable((By.XPATH,'//span[@class = "ms-Button-label label-48"][contains(text(), "Download all")]'))).click()
            time.sleep(20)
        except Exception as e:
            email_tabs = browser.find_elements_by_xpath()
        file_path = move_zip_file(browser,customer_type)
        destination_path = move_downloaded_file(browser,customer_type,file_path)

def download_attached_document(browser):
    count = 0
    element_len = browser.find_elements_by_xpath(f'//tr[contains(@class,"zA")]')
    print(len(element_len))
    ele = 1
    # for ele in range(1,len(element_len)+1):
        # try:
    element = browser.find_element_by_xpath(f'//tr[contains(@class,"zA")][{ele}]')
    mail_check_elemt = browser.find_element_by_xpath(f'//tr[contains(@class,"zA")][{ele}]/td[4]/div[1]/span/span').get_attribute('email')
    mail_id = [config['Mail'][x] for x in config['Mail']]
    WebDriverWait(browser, 45).until(EC.element_to_be_clickable((By.XPATH,f'//tr[contains(@class,"zA")][{ele}]/td[4]'))).click()
    browser.implicitly_wait(30)
    print(browser.find_element_by_xpath('//h2[contains(@id,":k")]').text)
        # subject_check = element.get_attribute('aria-label')
        # subject_len = browser.find_elements_by_xpath(f'//tr[contains(@class,"zA")][{ele}]/td[5]/div[1]/div/div/span[@data-thread-id]')
        # subject_len = browser.find_elements_by_xpath(f'//tr{ele}//span[@class="bog"]/span')
        # subject_len = WebDriverWait(browser, 65).until(EC.presence_of_element_located((By.XPATH,f'//tr{ele}//span[@class="bog"]/span')))
        # subject_check = browser.find_element_by_xpath(f'//tr[contains(@class,"zA")][{ele}]/td[4]/div[1]/span[@data-thread-id]')
        # subject_check = browser.findELement(By.XPATH(f'//tr[contains(@class,"zA")][{ele}]/td[4]/div[1]/span[2]')).getText()
        # print("Length : ",len(subject_len))
        # print([x.text for x in subject_len])
        # date_check = 
        # is_customer_mail = False
        # customer_type = ''
        # if mail_check_elemt and mail_check_elemt in mail_id:
        #     for x in config['Subject']:
        #         if config['Subject'][x] in subject_check:
        #             # if 
        #             is_customer_mail = True
        #             customer_type = 'suzlon' if "suzlon" in x else 'vestas'

        #     # if config['Subject']['keyword'] in element.get_attribute('aria-label') and config['Mail']['recipient_mail'] == mail_check_elemt
        #     if is_customer_mail:
        #         element.click()
        #         mail_element = browser.find_elements_by_xpath(f'//div[@class="_2le66D_cFAbkq67CrgZcmE"]')
        #         print("*************************",len(mail_element))
        #         time_element = browser.find_element_by_xpath(f'//div[@class="DWrY3hKxZTZNTwt3mx095"]').text
        #         # print(time_element)
        #         mail_recived_time = datetime.strptime(time_element,'%a %m/%d/%Y %I:%M %p')
        #         fixed_mail_time = config["Mail Time"]['temp_time']
        #         fixed_mail_time = datetime.strptime(fixed_mail_time,'%I:%M %p')
        #         fixed_mail_time = fixed_mail_time.replace(day=datetime.now().day,month=datetime.now().month,year=datetime.now().year)
        #         print(mail_recived_time.time())
        #         print("-----------",fixed_mail_time.time())
        #         if customer_type == "suzlon" and "daily" in subject_check:
        #             if mail_recived_time <= fixed_mail_time:
        #                 # for mail_id in range(2,len(mail_element)+1):
        #                 #     WebDriverWait(browser, 25).until(EC.element_to_be_clickable((By.XPATH,f'//div[@class="_2le66D_cFAbkq67CrgZcmE"][{mail_id}]'))).click()
        #                 download_button_click(browser)

        #             else:
        #                 sending_mail("RAPBOT notification",f"{customer_type} Daily status report is not been sent in please do check it")
                
        #         elif customer_type == "suzlon" and "weekly" in subject_check.lower():
        #             today = datetime.now()
        #             suzlonCheckFilePath = os.path.dirname(__file__)+"/suzlonCheck.json"
        #             suzlonWeekData = {"Download" : 1,"LastDownload":today.strftime("%d/%m/%Y")}
        #             if os.path.exists(suzlonCheckFilePath):
        #                 with open(suzlonCheckFilePath,"r") as suzlonVal:
        #                     suzlonWeekData = json.loads(suzlonVal.read())
        #                 if int(suzlonWeekData['Download']) < 2 and today.isocalendar()[1] == datetime.strptime(suzlonWeekData["LastDownload"],"%d/%m/%Y").isocalendar()[1]:
        #                     suzlonWeekData['Download'] += 1
        #                     download_button_click(browser)
        #                 elif today.isocalendar()[1] != datetime.strptime(suzlonWeekData["LastDownload"],"%d/%m/%Y").isocalendar()[1]:
        #                     download_button_click(browser)
        #             else:
        #                 download_button_click(browser)
        #             with open(suzlonCheckFilePath,"w+") as suzlonVal:
        #                 suzlonVal.write(json.dumps(suzlonWeekData))
        #                 # suzlonVal.write(str(suzlonWeekVal))

        #         elif customer_type == "suzlon":
        #             download_button_click


                    
    # excel_file_path = extract_file(browser,destination_path)
    # read_excel_file(browser,excel_file_path)

        # except Exception as e:
        #     print(e)
        #     print("Except part")
        #     break


def extract_file(browser,excel_dest_path):
    # print("REACHED extract file")
    file_name=set()
    file_list = os.listdir(excel_dest_path)
    ret_file_name=set()
    for file in file_list:
        # print(file)
        if os.path.isdir(os.path.join(excel_dest_path,file)):
            ret_file_name = extract_file(browser,os.path.join(excel_dest_path,file))
        else:
            if 'zip' in file:
                zip_file_path = os.path.join(excel_dest_path,file)
                with ZipFile(zip_file_path,"r") as zip_file:
                    zip_file.extractall(excel_dest_path)
                os.remove(zip_file_path)
                ret_file_name=extract_file(browser,excel_dest_path)
            else:
                file_name.add(os.path.join(excel_dest_path,file))
        file_name.update(ret_file_name)
    return file_name

def move_downloaded_file(browser,customer_type,file_name):
    dow_path = os.path.join(download_file_path,file_name)
    file_date = datetime.now().strftime("%d")
    file_month = datetime.now().strftime("%m")
    print(copy_file_path)
    des_path = os.path.join(copy_file_path,"SPI\\{}\\{}".format(file_month,file_date))
    print("Destination Path: ",des_path)
    os.makedirs(des_path+"\\{}".format(customer_type),exist_ok=True)
    try:    
        shutil.move(dow_path,des_path+"\\{}".format(customer_type))
    except Exception as e:
        logging.info(e)
    return des_path

def read_excel_file(browser,file_path):
    # print("FILE PATH \t",file_path)
    while(file_path):
        x=file_path.pop()
        # if "suzlon" in x.split('\\')[-1].lower() or 'vestas' in x.split('\\')[-1].lower() or "location" in x.split('\\')[-1].lower():
        if "suzlon" in x.split('\\')[-1].lower() or 'vestas' in x.split('\\')[-1].lower():
            dest_path = x.split("\\")
            dest_path.insert(-1,"OUPUT")
            os.makedirs("\\".join(dest_path[:-1]),exist_ok=True)
            dest_path = "\\".join(dest_path).replace('xls','xlsx') if "xlsx" not in x else "\\".join(dest_path)
            print(dest_path)
            excel_writer = pd.ExcelWriter(dest_path,engine='xlsxwriter')
            if "suzlon" in x.split('\\')[-1].lower() or "vestas" in x.split('\\')[-1].lower():
                # check_val = "Gen Date" if "suzlon" in x.split('\\')[-1].lower() else "Date"
                check_val = "Date"
                data = pd.read_excel(x,header=None,sheet_name=None)
                for sn in data:
                    df = data[sn]
                    # skip_val = df[df[0]==check_val].index[0]
                    skip_val = [x_i for x_i,item in df[0].iteritems() if "Date" in str(item)][0]
                    excel_data = df.iloc[skip_val:]
                    header_val = excel_data.iloc[:1]
                    excel_data = excel_data.iloc[1:]
                    excel_data.columns=header_val.iloc[0]
                    excel_data.to_excel(excel_writer,sheet_name=sn,index=False)
                excel_writer.save()
            # else:
            #     df=pd.read_excel(x,header=0)
            #     df.to_excel(dest_path,index=False)
                # try:
                #     connection  = mysql.connector.connect(host="139.59.46.107",
                #                                     port=6603,
                #                                     database="spi_group_windmill_data",
                #                                     user="root",
                #                                     password="root")
                #     if connection.is_connected():
                #         cursor = connection.cursor()
                #         # for x in files_path:
                #         if "suzlon" in dest_path.split('\\')[-1].lower() and "weekly" not in dest_path.split('\\')[-1].lower():
                #         # if 'location' in x.split('\\')[-1].lower():
                #             print(dest_path)
                #             wb = openpyxl.load_workbook(dest_path)
                #             # shet_obj = wb.active
                #             for sheet_name in wb.sheetnames:
                #                 customer_name = "M/s. KR WIND ENERGY LLP"
                #                 for x in wb[sheet_name]:
                #                     # if x[0].value and "Date" not in str(x[0].value) and "intial" not in str(x[0].value).lower() and "total" not in str(x[0].value).lower() and 'locationno' not in str(x[0].value).lower():
                #                     if x[0] and re.match(r"\d{4}-\d{2}-\d{2}",str(x[0])):
                #                         if "suzlon" in dest_path.split('\\')[-1].lower() and "weekly" not in dest_path.split('\\')[-1].lower():
                #                             db_command1 = f"INSERT INTO suzlon_xl_daily_hist(gendate,customername,state,site,section,mw,locno,genkwhday,genkwhmtd,genkwhytd,plfday,plfmtd,plfytd,mcavail,gf,fm,s,u,nor,genhrs,oprhrs) VALUES('{str(x[0].value).split(' ')[0]}','{x[1].value}','{x[2].value}','{x[3].value}','{x[4].value}',{float(check_float_val(x[5].value))},'{x[6].value}',{float(check_float_val(x[7].value))},{float(check_float_val(x[8].value))},{float(check_float_val(x[9].value))},{float(check_float_val(x[10].value))},{float(check_float_val(x[11].value))},{float(check_float_val(x[12].value))},{float(check_float_val(x[13].value)) if type(x[13].value)!=str else 0.0},{float(check_float_val(x[14].value))},{float(check_float_val(x[15].value))},{float(check_float_val(x[16].value))},{float(check_float_val(x[17].value))},{float(check_float_val(x[18].value))},{float(check_float_val(x[19].value))},{float(check_float_val(x[20].value))});"
                #                             db_command2=f"INSERT into spi_windmill_gen_daily_report(gendate,customername,locono,genkwhday,gf,fm,s,u,genhrs,oprhrs,netkwhday) values('{str(x[0].value).split(' ')[0]}','{x[1].value}','{x[6].value}',{float(check_float_val(x[7].value))},{float(check_float_val(x[14].value))},{float(check_float_val(x[15].value))},{float(check_float_val(x[16].value))},{float(check_float_val(x[17].value))},{float(check_float_val(x[19].value))},{float(check_float_val(x[20].value))});"
                #                             cursor.execute(db_command1)
                #                             # cursor.execute(db_command2)

                #                         if "suzlon" in dest_path.split('\\')[-1].lower() and "weekly" in dest_path.split('\\')[-1].lower():
                #                             db_command = f"INSERT INTO suzlon_xl_weekly_hist(gendate,mw,customername,htno,locno,reading_totalimport,reading_06_09am_1,reading_06_09pm_1,reading_09_10pm_1,reading_05_06amand09_06pm_1,reading_10pm_05am_1,reading_totalexport,reading_06_09am_2,reading_06_09pm_2,reading_09_10pm_2,reading_05_06amand09_06pm_2,reading_10pm_05am_2,reading_kvarhimportlag,reading_kvarhimportlead,reading_kvarhexportlag,reading_kvarhexportlead,reading_kvahimportreading,reading_kvahexportreading,reading_powerfactor,reading_percent_kvahimport,reading_monthcumulative,calc_totalimport,calc_06_09am_1,calc_06_09pm_1,calc_09_10pm_1,calc_05_06amand09_06pm_1,calc_10pm_05am_1,calc_totalexport,calc_06_09am_2,calc_06_09pm_2,calc_09_10pm_2,calc_05_06amand09_06pm_2,calc_10pm_05am_2,calc_kvarhimportlag,calc_kvarhimportlead,calc_kvarhexportlag,calc_kvarhexportlead,calc_kvahimportreading,calc_kvahexportreading,calc_powerfactor,calc_percent_kvahimport,calc_monthcumulative) values('{str(x[0].value).split(' ')[0]}',{float(check_float_val(x[1].value))},'{str(x[2].value)}','{str(x[3].value)}','{str(x[4].value)}',{float(check_float_val(x[5].value))},{float(check_float_val(x[6].value))},{float(check_float_val(x[7].value))},{float(check_float_val(x[8].value))},{float(check_float_val(x[9].value))},{float(check_float_val(x[10].value))},{float(check_float_val(x[14].value))},{float(check_float_val(x[15].value))},{float(check_float_val(x[16].value))},{float(check_float_val(x[17].value))},{float(check_float_val(x[18].value))},{float(check_float_val(x[19].value))},{float(check_float_val(x[11].value))},{float(check_float_val(x[12].value))},{float(check_float_val(x[20].value))},{float(check_float_val(x[21].value))},{float(check_float_val(x[13].value))},{float(check_float_val(x[22].value))},{float(check_float_val(x[23].value))},{float(check_float_val(x[24].value))},{float(check_float_val(x[25].value))},{float(check_float_val(x[26].value))},{float(check_float_val(x[27].value))},{float(check_float_val(x[28].value))},{float(check_float_val(x[29].value))},{float(check_float_val(x[30].value))},{float(check_float_val(x[31].value))},{float(check_float_val(x[32].value))},{float(check_float_val(x[36].value))},{float(check_float_val(x[37].value))},{float(check_float_val(x[38].value))},{float(check_float_val(x[39].value))},{float(check_float_val(x[40].value))},{float(check_float_val(x[41].value))},{float(check_float_val(x[33].value))},{float(check_float_val(x[34].value))},{float(check_float_val(x[42].value))},{float(check_float_val(x[43].value))},{float(check_float_val(x[35].value))},{float(check_float_val(x[44].value))},{float(check_float_val(x[45].value))},{float(check_float_val(x[46].value))});"
                #                             cursor.execute(db_command)

                #                         if "vestas" in dest_path.split('\\')[-1].lower():
                #                             db_command1 = f"INSERT into vestas_xl_daily_hist(gendate,mw,customername,htno,locno,readingtakentime,cml_runhrs,cml_genhrs,cml_g0,cml_gen,cml_totalprod,cml_totalimport,cml_06_09am_1,cml_18_21pm_1,cml_21_22pm_1,cml_05_06amand09_18pm_1,cml_22_05am_1,cml_totalexport,cml_06_09am_2,cml_18_21pm_2,cml_21_22pm_2,cml_05_06amand09_18pm_2,cml_22_05am_2,cml_rkvahr_imp,cml_rkvahr_exp,daily_runhrs,daily_genhrs,daily_g0,daily_gen,daily_totalprod,daily_totalimport,daily_06_09am_1,daily_18_21pm_1,daily_21_22pm_1,daily_05_06amand09_18pm_1,daily_22_05am_1,daily_totalexport,daily_06_09am_2,daily_18_21pm_2,daily_21_22pm_2,daily_05_06amand09_18pm_2,daily_22_05am_2,daily_rkvahr_imp,daily_rkvahr_exp,gf,fm,sch,unsch,manualstoppage,readingnotavailable,total,remarks) values('{str(x[0].value).split(' ')[0]}',{float(check_float_val(x[1].value))},'{customer_name}','{x[2].value}','{x[3].value}','{x[4].value}',{float(check_float_val(x[5].value))},{float(check_float_val(x[6].value))},{float(check_float_val(x[7].value))},{float(check_float_val(x[8].value))},{float(check_float_val(x[9].value))},{float(check_float_val(x[10].value))},{float(check_float_val(x[11].value))},{float(check_float_val(x[12].value))},{float(check_float_val(x[13].value))},{float(check_float_val(x[14].value))},{float(check_float_val(x[15].value))},{float(check_float_val(x[16].value))},{float(check_float_val(x[17].value))},{float(check_float_val(x[18].value))},{float(check_float_val(x[19].value))},{float(check_float_val(x[20].value))},{float(check_float_val(x[21].value))},{float(check_float_val(x[22].value))},{float(check_float_val(x[23].value))},{float(check_float_val(x[24].value))},{float(check_float_val(x[25].value))},{float(check_float_val(x[26].value))},{float(check_float_val(x[27].value))},{float(check_float_val(x[28].value))},{float(check_float_val(x[29].value))},{float(check_float_val(x[30].value))},{float(check_float_val(x[31].value))},{float(check_float_val(x[32].value))},{float(check_float_val(x[33].value))},{float(check_float_val(x[34].value))},{float(check_float_val(x[35].value))},{float(check_float_val(x[36].value))},{float(check_float_val(x[37].value))},{float(check_float_val(x[38].value))},{float(check_float_val(x[39].value))},{float(check_float_val(x[40].value))},{float(check_float_val(x[41].value))},{float(check_float_val(x[42].value))},{float(check_float_val(x[43].value))},{float(check_float_val(x[44].value))},{float(check_float_val(x[45].value))},{float(check_float_val(x[46].value))},{float(check_float_val(x[47].value))},{float(check_float_val(x[48].value))},{float(check_float_val(x[49].value))},'{x[50].value}');"
                #                             genkwhValue = float(check_float_val(x[16].value)) - float(check_float_val(x[10].value))
                #                             db_command2=f"INSERT into spi_windmill_gen_daily_report(gendate,customername,locono,genkwhday,gf,fm,s,u,genhrs,oprhrs,netkwhday) values('{str(x[0].value).split(' ')[0]}','{customer_name}','{x[3].value}',{genkwhValue},{float(check_float_val(x[43].value))},{float(check_float_val(x[44].value))},{float(check_float_val(x[45].value))},{float(check_float_val(x[46].value))},{float(check_float_val(x[6].value))},{float(check_float_val(x[5].value))});"
                #                             cursor.execute(db_command1)
                #                             # cursor.execute(db_command2)

                #                         # if "location" in x.split('\\')[-1].lower():
                #                         #     db_command=f"INSERT into location_master(locno,weghtno,wegcapacitykw,make) values('{str(x[0].value)}',{int(x[1].value)},{int(x[2].value)},'{x[3].value}');"
                #                         #     cursor.execute(db_command)
                #                         print("---------",db_command)
                #         connection.commit()
                #         cursor.close()
                # except Exception as e:
                #     print("The error is \t:",e)

browser.get(config["Website"]["url"])
# browser.implicitly_wait(20)
login_gmail(browser)
# sending_mail("This is test message")
# read_excel_file(browser,file_path)