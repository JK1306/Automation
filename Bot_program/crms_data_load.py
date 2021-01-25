from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium import webdriver
import configparser
import shutil
import traceback
import time
from datetime import datetime
from datetime import timedelta
import logging
import db
import os

def is_page_loaded(browser):
    print("In page load")
    result = browser.execute_script('return document.readyState;')
    print(result)
    while result != "complete":
        result = browser.execute_script('return document.readyState;')

def take_browser_ss(browser):
    current_date = datetime.now()
    screen_shot_path = os.path.join(os.path.dirname(__file__),'Screenshots')
    ss_file_path = os.path.join(screen_shot_path,f'suzlon_ss_{current_date.strftime("%Y%m%d")}.png')
    os.makedirs(screen_shot_path,exist_ok=True)
    count = 1
    while True:
        if not os.path.exists(ss_file_path):
            browser.save_screenshot(ss_file_path)
            return ss_file_path
        else:
            ss_file_path = os.path.join(os.path.dirname(ss_file_path),f"suzlon_ss_{current_date.strftime('%Y%m%d')}_{count}.png")
            count += 1

def find_element_xpath(browser,xpath_str,click_flag=False,send_key=None):
    while True:
        try:
            if click_flag:
                WebDriverWait(browser,45).until(EC.element_to_be_clickable((By.XPATH,xpath_str))).click()
            elif send_key:
                print('Send key flag is enabled')
                WebDriverWait(browser,45).until(EC.element_to_be_clickable((By.XPATH,xpath_str))).send_keys(send_key)
            else:
                WebDriverWait(browser,45).until(EC.element_to_be_clickable((By.XPATH,xpath_str)))
            break
        except:
            pass

def is_ajax_loaded(browser):
    load_screen = browser.find_element_by_xpath('//*[@id="ContentPlaceHolder1_UpdateProgress1"]').get_attribute('style')
    while 'none' not in load_screen.lower():
        load_screen = browser.find_element_by_xpath('//*[@id="ContentPlaceHolder1_UpdateProgress1"]').get_attribute('style')


def fetch_downloaded_file(download_file_path):
    logging.debug('Entered into fetch_downloaded_file function')
    recent_file_time = None
    files = []
    file_name = None
    while not files or '.xls' not in file_name:
        for x_i,x in enumerate(os.listdir(download_file_path)):
            if x not in files and 'crdownload' not in x:
                if not recent_file_time:
                    recent_file_time = os.path.getmtime(os.path.join(download_file_path,x))
                    file_name = x
                else:
                    current_file_time = os.path.getmtime(os.path.join(download_file_path,x))
                    if current_file_time > recent_file_time:
                        recent_file_time = current_file_time
                        file_name = x
        current_time = time.time() - 400
        if file_name and 'crdownload' not in file_name and '.xls' in file_name and current_time <= recent_file_time:
            files.append(file_name)
        # print(files)
            # file_name = None
        recent_file_time = None
    return files[0]

def move_file(file_name,source_path,dest_path):
    os.makedirs(dest_path,exist_ok=True)
    shutil.move(source_path,dest_path)
    return os.path.join(dest_path,file_name)

def dashboard(browser,config,download_file_path,exception_flag):
    WebDriverWait(browser,45).until(EC.element_to_be_clickable((By.XPATH,'//div[@class="ContentBlock"]/h2')))
    WebDriverWait(browser,45).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ContentPlaceHolder1_DDLCustomer"]'))).click()
    customer_list = browser.find_elements_by_xpath('//*[@id="ContentPlaceHolder1_DDLCustomer"]/option')
    customer_list = [x.text for x in customer_list]
    for x in range(len(customer_list)):
        if 'select' not in customer_list[x].lower():
            print(customer_list[x])
            find_element_xpath(browser,f'//*[@id="ContentPlaceHolder1_DDLCustomer"]/option[{x+1}]',True)
            find_element_xpath(browser,'//*[@id="ContentPlaceHolder1_DDLMainSite"]',True)
            is_ajax_loaded(browser)
            site_list = browser.find_elements_by_xpath(f'//*[@id="ContentPlaceHolder1_DDLMainSite"]/option')
            site_list = [txt.text for txt in site_list]
            for y in range(1,len(site_list)):
                print(site_list[y])
                find_element_xpath(browser,f'//*[@id="ContentPlaceHolder1_DDLMainSite"]/option[{str(y+1)}]',True)
                find_element_xpath(browser,'//*[@id="ContentPlaceHolder1_DDLSite"]',True)
                sector_list = browser.find_elements_by_xpath('//*[@id="ContentPlaceHolder1_DDLSite"]/option')
                sector_list = [txt.text for txt in sector_list]
                for z in range(1,len(sector_list)):
                    print(sector_list[z])
                    find_element_xpath(browser,f'//*[@id="ContentPlaceHolder1_DDLSite"]/option[{str(z+1)}]',True)
                    is_ajax_loaded(browser)
                    current_date = datetime.now()
                    # current_date = datetime(2022,4,1)
                    current_year = current_date.year if current_date.month>=4 else current_date.year-1
                    if exception_flag:
                        mail_time = config['Exception Config']['suzlon_daily'].split('-')
                        from_date = datetime.strptime(mail_time[0].strip(), '%d/%m/%Y %I:%M %p').strftime('%d-%b-%Y')
                        to_date = datetime.strptime(mail_time[1].strip(), '%d/%m/%Y %I:%M %p').strftime('%d-%b-%Y')
                        # mail_time = [datetime.strptime(x.strip(), '%d/%m/%Y %I:%M %p') for x in mail_time]    
                    else:
                        from_date = '01-Apr-{}'.format(str(current_year))
                        # to_date = '01-Jan-2021'
                        to_date = current_date.strftime('%d-%b-%Y')
                    browser.execute_script(f"document.getElementById('ContentPlaceHolder1_txtFromDate').value='{from_date}';")
                    browser.execute_script(f"document.getElementById('ContentPlaceHolder1_txtToDate').value='{to_date}';")
                    find_element_xpath(browser,'//*[@id="ContentPlaceHolder1_BtnViewRpt"]',click_flag=True)
                    is_page_loaded(browser)
                    result_list = browser.find_elements_by_xpath('//*[@id="ContentPlaceHolder1_gvDailyGenData"]/tbody/tr[td]')
                    print('---------',len(result_list))
                    print(browser.find_element_by_xpath('//*[@id="ContentPlaceHolder1_txtFromDate"]').text)
                    for res in range(len(result_list)):
                        browser.find_element_by_xpath(f'//*[@id="ContentPlaceHolder1_gvDailyGenData"]/tbody/tr[td][{res+1}]/td[3]/a').click()
                        is_page_loaded(browser)
                        time.sleep(2)
                        downloaded_file = fetch_downloaded_file(download_file_path)
                        print('-------------',downloaded_file)
                        move_dir = os.path.join(os.path.dirname(__file__),'Files','Suzlon_daily',f'{current_date.strftime("%Y%m%d")}',customer_list[x],site_list[y],sector_list[z])
                        os.makedirs(move_dir,exist_ok=True)
                        file_path = os.path.join(download_file_path,downloaded_file)
                        file_path = move_file(downloaded_file,file_path,move_dir)
                        db.convert_xml_to_df(file_path,config)            
            
def login(browser,config,download_file_path,exception_flag):
    WebDriverWait(browser,45).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="txtUserId"]'))).send_keys(config['Bot']['login_id'])
    WebDriverWait(browser,45).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="txtPassword"]'))).send_keys(config['Bot']['paswd'])
    WebDriverWait(browser,45).until(EC.element_to_be_clickable((By.XPATH,'//*[@id="img_login"]'))).click()
    is_page_loaded(browser)
    try:
        err_msg_ele = browser.find_element_by_xpath('//*[@id="lblErrorMsg"]')
        err_msg = err_msg_ele.text
        if err_msg:
            print(err_msg)
            mail_time = config['Exception Config']['suzlon_daily'].split('-')
            from_date = datetime.strptime(mail_time[0].strip(), '%d/%m/%Y %I:%M %p').strftime('%d-%b-%Y')
            to_date = datetime.strptime(mail_time[1].strip(), '%d/%m/%Y %I:%M %p').strftime('%d-%b-%Y')
            print('From date : ',from_date)
            print('To date : ',to_date)
            ss_file_path = take_browser_ss(browser)
            db.send_mail(config,'RAPBot CRMS Log-in Failed',f'RAPBot have noted that Error occured during Login the error msg is "{err_msg}". PLease do look after the issue',0,ss_file_path)
    except:
        dashboard(browser,config,download_file_path,exception_flag)

def start(browser,config,download_file_path,exception_flag=None):
    try:
        for handler in logging.root.handlers[:]:
            logging.root.removeHandler(handler)
        log_file = os.path.join(os.path.dirname(__file__),'Logs',f'suzlone_data_load_{datetime.now().strftime("%d%m%Ys")}.log')
        os.makedirs(os.path.dirname(log_file),exist_ok=True)
        logging.basicConfig(filename=log_file,
                            format='%(asctime)s %(message)s',
                            filemode='a',
                            level = logging.DEBUG)
        logging.debug('Started to process ftp confirm')
        browser.get(config['Bot']['url'])
        logging.debug('URL get launched in browser')
        login(browser,config,download_file_path,exception_flag)
    except:
        error_msg = traceback.format_exc()
        print(error_msg)
        logging.debug(f'Error occured in ftp_confirm file the error : {error_msg}')
        db.send_mail(config,'RAPBot Error Notification',f'Hi,\n\tRAPBot have noticed that Error occured in ftp_confirm error msg : {error_msg}',0)

if __name__ == "__main__":
    config = configparser.ConfigParser()
    config.read(os.path.join(os.path.dirname(__file__),'Config','config.ini'))
    chrome_options = webdriver.ChromeOptions()
    # chrome_options.add_argument('disable-infobars')
    download_file_path = os.path.join(os.path.dirname(__file__),'Download')
    os.makedirs(download_file_path,exist_ok=True)
    download_file_path = os.path.abspath(download_file_path)
    prefs = {"download.default_directory": download_file_path}
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_experimental_option("prefs", prefs)
    browser = webdriver.Chrome(config['Path']['chromedriver'],options=chrome_options)
    browser.maximize_window()
    start(browser,config,download_file_path)
