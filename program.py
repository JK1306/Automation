import os
import json
import logging
from datetime import datetime
from selenium import webdriver  
from selenium.webdriver.common.keys import Keys  
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.firefox.options import Options as FirefoxOptions

def create_json_file():
    today = datetime.now()
    file_path = os.path.dirname(__file__)+"/test.json"
    print(file_path)
    Data = {"Download" : 1,"LastDownload":today.strftime("%d/%m/%Y %I:%M %p")}
    if os.path.exists(file_path):
        with open(file_path,"r") as suzlonVal:
            read_Data = json.loads(suzlonVal.read())
        if int(read_Data['Download']) < 2 and today.isocalendar()[1] == datetime.strptime(read_Data["LastDownload"],"%d/%m/%Y %I:%M %p").isocalendar()[1]:
            read_Data['Download'] += 1
            Data = read_Data
        elif today.isocalendar()[1] == datetime.strptime(read_Data["LastDownload"],"%d/%m/%Y %I:%M %p").isocalendar()[1]:
            print("It came here")
            Data = read_Data
            pass
            # send email for excess weekly data sent
    with open(file_path,"w+") as suzlonVal:
        suzlonVal.write(json.dumps(Data))

def run_selenium_headless():
    chrome_options = Options()  
    # chrome_options = webdriver.ChromeOptions()
    # chrome_options.add_argument("--headless")
    chrome_options.add_argument('--disable-gpu')
    driver = webdriver.Chrome(executable_path=r"C:\Users\Jaikishore\Documents\RAP\SPI Group\chromedriver.exe")
    driver.get("http://www.google.com")
    print(driver.find_element_by_xpath('//*[@id="tsf"]/div[2]/div[1]/div[3]/center/input[2]').get_attribute('value'))

def run_firefox():
    firefox_option = FirefoxOptions()
    firefox_option.add_argument('--headless')
    browser = webdriver.Firefox(executable_path=r"C:\Users\Jaikishore\Downloads\geckodriver-v0.27.0-win64\geckodriver.exe",options=firefox_option)
    browser.get("http://www.google.com")
    print(browser.find_element_by_xpath('//*[@id="tsf"]/div[2]/div[1]/div[3]/center/input[2]').get_attribute('value'))

def create_log():
    logging.info("Hellow world")
    logging.warning("This is warning log")
    logging.error("This is error log")
# create_json_file()
# run_selenium_headless()
# run_firefox()
create_log()