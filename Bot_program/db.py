from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.text import MIMEText
import smtplib
import mysql.connector as mc
import pandas as pd
import configparser
import re
import os
from bs4 import BeautifulSoup
from datetime import datetime
import logging
import traceback

def send_mail(config, subject, body, mail_type, image_file_path=None):
    try:
        logging.info("Enters send_mail function")
        msg = MIMEMultipart()
        msg['From'] = config['mail config']['from']
        if mail_type:
            # bussiness
            msg['To'] = ', '.join(config['mail config']['business'].split(','))
        else:
            # admin
            msg['To'] = ', '.join(config['mail config']['admin'].split(','))
        msg['Subject'] = subject
        body = f'{body}'
        msg.attach(MIMEText(body, 'plain'))
        if image_file_path:
            print(image_file_path)
            img_data = open(image_file_path,'rb').read()
            image = MIMEImage(img_data,name=os.path.basename(image_file_path))
            msg.attach(image)
        server = smtplib.SMTP(config['mail config']['host'], config['mail config']['port']) 
        server.ehlo()
        server.starttls()
        server.ehlo()
        server.login(config['mail config']['from'],config['mail config']['paswd'])
        server.sendmail(msg['From'],msg['To'].split(','),msg.as_string())
        server.quit()
    except:
        logging.error(f"Error occured in : {traceback.format_exc()}")

def get_cursor(config):
    connect = mc.connect(
        host=config.get('DB Config','host'),
        port=config["DB Config"]["port"],
        database=config["DB Config"]["database"],
        user=config["DB Config"]["user_name"],
        password=config["DB Config"]["paswd"]
    )
    if connect.is_connected():
        print("Connected")
        # connect.autocommit =True
        cursor = connect.cursor()
    return [connect,cursor]

def check_float_val(data):
    try:
        float(data)
        return float(data) if float(data) else 0.0
    except:
        return 0.0

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

def insert_into_db(config,data_type,doc_val,file_name,customer_type):
    duplicate_record_sd = []
    recordInserted = []
    db_cur = get_cursor(config)
    connection = db_cur[0]
    cursor = db_cur[1]
    location = read_location_master(cursor)
    if data_type=='generation':
        print("It entered into generation")
        logging.info("INFO :: -------- > Table used : suzlon_xl_daily_hist and spi_windmill_gen_daily_report")
        for column_val in doc_val.iterrows():
            try:
                x = column_val[1]
                if re.match(r"\d{4}-\d{2}-\d{2}\s\d{2}:\d{2}:\d{2}", str(x.get('genDate'))) or re.match(r"\d{2}-[A-z]{3}-\d{4}", str(x.get('genDate'))):

                    genDate = str(x.get('genDate')).split(' ')[0] if re.match(r"\d{4}-\d{2}-\d{2}\s\d{2}:\d{2}:\d{2}", str(x[0])) else datetime.strptime(x.get('genDate'), "%d-%b-%Y").strftime("%Y-%m-%d")
                    print(x.get('customerName'))
                    customerName = "SPI Power" if "spi" in re.sub(r"\s+", '', x.get('customerName')).lower() or "skr" in re.sub(r"\s+", '', x.get('customerName')).lower() else "KR Wind Energy" if "kr" in re.sub(r"\s+", '', x.get('customerName')).lower() else ''
                    print(customerName)

                    locNoVal = re.sub(r"\s+", '', x.get('locNo')) if "TP06" not in x.get('locNo') else "TP6"
                    
                    location_values = location.get(locNoVal)

                    if not check_valuein_reporting_layer(cursor,[genDate,customerName,x.get('locNo')]):
                        db_command = f"update spi_windmill_gen_daily_report set\
                        mckwhday={float(check_float_val(x.get('genkwhDay')))},\
                        gf={float(check_float_val(x.get('gf')))},\
                        fm={float(check_float_val(x.get('fm')))},\
                        sch={float(check_float_val(x.get('s')))},\
                        unsch={float(check_float_val(x.get('u')))},\
                        genhrs={float(check_float_val(x.get('genHrs')))},\
                        oprhrs={float(check_float_val(x.get('oprHrs')))},\
                        mw={float(check_float_val(x.get('mw')))}\
                        where gendate='{genDate}' and companyname='{customerName}' and locno='{x.get('locNo')}';"
                        logging.info(f'Execute command : {db_command}')
                        cursor.execute(db_command)

                        query = f"SELECT mckwhday,gf,fm,sch,unsch,genhrs,oprhrs,mw FROM spi_windmill_gen_daily_report where gendate='{genDate}' and companyname='{customerName}' and locno='{x.get('locNo')}';"
                        
                        xml_val = (float(check_float_val(x.get('genkwhDay'))),float(check_float_val(x.get('gf'))),float(check_float_val(x.get('fm'))),float(check_float_val(x.get('s'))),float(check_float_val(x.get('u'))),float(check_float_val(x.get('genHrs'))),float(check_float_val(x.get('oprHrs'))),float(check_float_val(x.get('mw'))))
                        
                        cursor.execute(query)
                        
                        db_val = cursor.fetchall()[0]
                        insert_flag = all([x==y for x,y in zip(xml_val,db_val)])
                        if not insert_flag:
                            db_command2 = f"insert into suzlon_xl_daily_hist(gendate,customername,state,site,section,mw,locno,genkwhday,genkwhmtd,genkwhytd,plfday,plfmtd,plfytd,mcavail,gf,fm,s,u,nor,genhrs,oprhrs) values('{genDate}','{x.get('customerName')}','{x.get('state')}','{x.get('site')}','{x.get('section')}',{float(check_float_val(x.get('mw')))},'{x.get('locNo')}',{float(check_float_val(x.get('genkwhDay')))},{float(check_float_val(x.get('genkwhMtd')))},{float(check_float_val(x.get('genkwhYtd')))},{float(check_float_val(x.get('plfDay')))},{float(check_float_val(x.get('plfMtd')))},{float(check_float_val(x.get('plfYtd')))},{float(check_float_val(x.get('mcAvail')))},{float(check_float_val(x.get('gf')))},{float(check_float_val(x.get('fm')))},{float(check_float_val(x.get('s')))},{float(check_float_val(x.get('u')))},{float(check_float_val(x.get('nor',x.get('rna'))))},{float(check_float_val(x.get('genHrs')))},{float(check_float_val(x.get('oprHrs')))});"
                            cursor.execute(db_command2)
                    else:
                        db_command = f"insert into suzlon_xl_daily_hist(gendate,customername,state,site,section,mw,locno,genkwhday,genkwhmtd,genkwhytd,plfday,plfmtd,plfytd,mcavail,gf,fm,s,u,nor,genhrs,oprhrs) values('{genDate}','{x.get('customerName')}','{x.get('state')}','{x.get('site')}','{x.get('section')}',{float(check_float_val(x.get('mw')))},'{x.get('locNo')}',{float(check_float_val(x.get('genkwhDay')))},{float(check_float_val(x.get('genkwhMtd')))},{float(check_float_val(x.get('genkwhYtd')))},{float(check_float_val(x.get('plfDay')))},{float(check_float_val(x.get('plfMtd')))},{float(check_float_val(x.get('plfYtd')))},{float(check_float_val(x.get('mcAvail')))},{float(check_float_val(x.get('gf')))},{float(check_float_val(x.get('fm')))},{float(check_float_val(x.get('s')))},{float(check_float_val(x.get('u')))},{float(check_float_val(x.get('nor',x.get('rna'))))},{float(check_float_val(x.get('genHrs')))},{float(check_float_val(x.get('oprHrs')))});"
                        logging.info(f'Execute command : {db_command}')
                        db_command2 = f"insert into spi_windmill_gen_daily_report(gendate,companyname,locno,mckwhday,gf,fm,sch,unsch,genhrs,oprhrs,mw,section,site,make,htno) values('{genDate}','{customerName}','{locNoVal}',{float(check_float_val(x.get('genkwhDay')))},{float(check_float_val(x.get('gf')))},{float(check_float_val(x.get('fm')))},{float(check_float_val(x.get('s')))},{float(check_float_val(x.get('u')))},{float(check_float_val(x.get('genHrs')))},{float(check_float_val(x.get('oprHrs')))},{float(check_float_val(x.get('mw')))},'{x.get('section')}','{x.get('site')}','{location_values[0]}','{location_values[3]}');"
                        logging.info(f'Execute command : {db_command2}')
                        cursor.execute(db_command)
                        cursor.execute(db_command2)
            except Exception as e:
                traceback.print_exc()
                logging.error(f"ERROR :: An error occured while inserting GENERAL data from {file_name} into suzlon_xl_daily_hist and spi_windmill_gen_daily_report Database ----------------------> {e}")
                send_mail(config,f"RAP Bot notification for error in Database insert",f"GENERATION DATA (general sheet) from {file_name} or {customer_type} type is not Inserted into  Database Error is : {e}", 0)
        print("\n\nSuccessfully Inserted in suzlon daily\n\n")
        logging.info(f"INFO :: Data from {file_name} is Successfully Inserted into suzlon_xl_daily_hist Database")
        if recordInserted:
            send_mail(f"RAP Bot Successfull data uploaded notification for {customer_type}", f"Data from {file_name} is Successfully Inserted into  Database", 0)
            logging.info(f"INFO :: Data from {file_name} is Successfully Inserted into spi_windmill_gen_daily_report Database")    
    connection.commit()
    cursor.close()

def convert_xml_to_df(file_path,config):
    data_df = pd.DataFrame()
    with open(file_path,'r') as file:
        soup = BeautifulSoup(file,'xml')
        for sheet in soup.findAll('Worksheet'): 
            if 'generation' in sheet['ss:Name'].lower():
                sheet_as_list = []
                for row in sheet.findAll('Row'):
                    row_as_list = []
                    for cell in row.findAll('Cell'):
                        if cell.Data:
                            row_as_list.append(cell.Data.text)
                        else:
                            row_as_list.append('')
                    sheet_as_list.append(row_as_list)
                data_df = data_df.append(sheet_as_list)
    data_df.columns = data_df.iloc[0]
    data_df.drop(0,inplace=True)
    # print('------------------------')
    print(data_df)
    read_data(config,file_path,data_df)

def read_data(config,file_path,doc_val):
    file_name = file_path.split('/')[-1]
    date_len = len(doc_val[['Gen. Date']].drop_duplicates())
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

    insert_into_db(config,'generation',doc_val,file_name,'suzlon_daily')
    # test_db(config,doc_val)

def test_db(config,doc_val):
    # query='SELECT * FROM spi_windmill_gen_daily_report;'
    db_connect = get_cursor(config)
    connection = db_connect[0]
    cursor = db_connect[1]
    # print(len(cursor.fetchall()))
    for _,x in doc_val.iterrows():
        genDate = str(x.get('genDate')).split(' ')[0] if re.match(r"\d{4}-\d{2}-\d{2}\s\d{2}:\d{2}:\d{2}", str(x[0])) else datetime.strptime(x.get('genDate'), "%d-%b-%Y").strftime("%Y-%m-%d")
        customerName = "SPI Power" if "spi" in re.sub(r"\s+", '', x.get('customerName')).lower() or "skr" in re.sub(r"\s+", '', x.get('customerName')).lower() else "KR Wind Energy" if "kr" in re.sub(r"\s+", '', x.get('customerName')).lower() else ''
        # print(genDate,'---',customerName,'---',x.get('locNo'))
        if check_valuein_reporting_layer(cursor,[genDate,customerName,x.get('locNo')]):
            query = f"SELECT mckwhday,gf,fm,sch,unsch,genhrs,oprhrs,mw FROM spi_group_windmill_data.spi_windmill_gen_daily_report where gendate='{genDate}' and companyname='{customerName}' and locno='{x.get('locNo')}';"
            xml_val = (float(check_float_val(x.get('genkwhDay'))),float(check_float_val(x.get('gf'))),float(check_float_val(x.get('fm'))),float(check_float_val(x.get('s'))),float(check_float_val(x.get('u'))),float(check_float_val(x.get('genHrs'))),float(check_float_val(x.get('oprHrs'))),float(check_float_val(x.get('mw'))))
            cursor.execute(query)
            db_val = cursor.fetchall()[0]
            insert_flag = all([x==y for x,y in zip(xml_val,db_val)])
            if not insert_flag:
                print('Perform Insert')
                print(xml_val)
                print(db_val)
            # if not (xml_val==db_val):
            #     print(xml_val)
            #     print('\n')
            #     print(db_val)
    cursor.close()

if __name__ == "__main__":
    config=configparser.ConfigParser()
    config.read(os.path.join(os.path.dirname(__file__),'Config','config.ini'))
    path=r'C:\Users\Jaikishore\Documents\RAP\SPI Group\Data_load_2\Download\DailyGenerationReport_637455303429672865.xls'
    convert_xml_to_df(path,config)
    # test_db(config)
