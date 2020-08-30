import pandas as pd
import os
from openpyxl import Workbook
import openpyxl
import glob,re
import mysql.connector
from mysql.connector import Error

def check_data(data,data_type):
    return data if data else '' if data_type == str else 0.0 
def check_float_val(data):
    try:
        float(data)
        return data if data else 0.0
    except:
        return 0.0

files = glob.glob(r'C:\Users\Jaikishore\Downloads\mskrenergyllpspipowerofwegsdailygenerationrepo\*')
# files = [r"C:\Users\Jaikishore\Downloads\Fw__Final_Format_for_Daily_MIS\Daily MIS Report -_FY 2019-20.xls"]
locatiotn_file = r"C:\Users\Jaikishore\Downloads\Location_Master.xlsx"
# df = pd.read_excel(,sheet_name=None)
connection = mysql.connector.connect(host="139.59.46.107",
                                port=6603,
                                database="spi_group_windmill_data",
                                user="root",
                                password="root")

def insert_into_tables():
    def read_location_master(cursor):
        cursor.execute("select * from location_master;")
        location_data = cursor.fetchall()
        location_dic = {}
        for x in location_data:
            location_dic[x[0]] = [x[3],x[4],x[5]]
        return location_dic
        
    if connection.is_connected():
        cursor = connection.cursor()
        inserted_val = {'suzlon':[],'vestas':[]}
        location = read_location_master(cursor)
        for file_name in files:
            sheet_val=pd.read_excel(file_name,sheet_name=None)
            print("\nFile Name : ",file_name)
            insert_df = ['KR Suzlon 1920','SPIP- Suzlon -FY 2019-20']
            df_header = ""
            for sheet_name in sheet_val:
                print(sheet_name)
                if True:
                    doc_val = sheet_val[sheet_name].fillna('')
                    for x_i,x in doc_val.iterrows():
                        if "date" in str(x[0]).lower():
                            df_header = doc_val.iloc[x_i-1:x_i+1].fillna('')
                            break
                    df = doc_val.iloc[x_i+1:]
                    head_val = ''
                    for x_i,x in df_header.iteritems():
                        if x.iloc[0]:
                            head_val = 'cml_' if 'cumulative' in x.iloc[0].lower() else 'daily_' if 'daily' in x.iloc[0].lower() else ''
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
                    # for suzlone daily
                    """
                    for y in doc_val.columns:
                        if "date" in y.lower():
                            doc_val.rename(columns={y:'genDate'},inplace=True)
                        if "customer" in y.lower() or 'company' in y.lower():
                            doc_val.rename(columns={y:'customerName'},inplace=True)
                        if "state" in y.lower() or "site" in y.lower() or "section" in y.lower() or y.lower() == "mw" or y.lower() == "gf" or y.lower() == "fm" or y.lower() == "s" or y.lower() == "u" or y.lower() == "nor":
                            doc_val.rename(columns={y:y.lower()},inplace=True)
                        if 'htsc' in y.lower():
                            doc_val.rename(columns={y:'htscNo'},inplace=True)
                        if 'loc' in y.lower():
                            doc_val.rename(columns={y:'locNo'},inplace=True)
                        if 'gen' in y.lower() and 'day' in y.lower():
                            doc_val.rename(columns={y:'genkwhDay'},inplace=True)
                        if 'gen' in y.lower() and 'mtd' in y.lower():
                            doc_val.rename(columns={y:'genkwhMtd'},inplace=True)
                        if 'gen' in y.lower() and 'ytd' in y.lower():
                            doc_val.rename(columns={y:'genkwhYtd'},inplace=True)
                        if 'plf' in y.lower() and 'day' in y.lower():
                            doc_val.rename(columns={y:'plfDay'},inplace=True)
                        if 'plf' in y.lower() and 'mtd' in y.lower():
                            doc_val.rename(columns={y:'plfMtd'},inplace=True)
                        if 'plf' in y.lower() and 'ytd' in y.lower():
                            doc_val.rename(columns={y:'plfYtd'},inplace=True)
                        if 'avail' in y.lower():
                            doc_val.rename(columns={y:'mcAvail'},inplace=True)
                        if 'hrs' in y.lower():
                            if 'gen' in y.lower():
                                doc_val.rename(columns={y:'genHrs'},inplace=True)
                            else:
                                doc_val.rename(columns={y:'oprHrs'},inplace=True)
                    """

                    for column_val in df.iterrows():
                        x = column_val[1]
                        if re.match(r"\d{4}-\d{2}-\d{2}\s\d{2}:\d{2}:\d{2}",str(x[0])):
                            # db_command = f"insert into suzlon_xl_daily_hist(gendate,customername,state,site,section,mw,locno,genkwhday,genkwhmtd,genkwhytd,plfday,plfmtd,plfytd,mcavail,gf,fm,s,u,nor,genhrs,oprhrs) values('{str(x.get('genDate')).split(' ')[0]}','{x.get('customerName')}','{x.get('state')}','{x.get('site')}','{x.get('section')}',{float(check_float_val(x.get('mw')))},'{x.get('locNo')}',{float(check_float_val(x.get('genkwhDay')))},{float(check_float_val(x.get('genkwhMtd')))},{float(check_float_val(x.get('genkwhYtd')))},{float(check_float_val(x.get('plfDay')))},{float(check_float_val(x.get('plfMtd')))},{float(check_float_val(x.get('plfYtd')))},{float(check_float_val(x.get('mcAvail')))},{float(check_float_val(x.get('gf')))},{float(check_float_val(x.get('fm')))},{float(check_float_val(x.get('s')))},{float(check_float_val(x.get('u')))},{float(check_float_val(x.get('nor')))},{float(check_float_val(x.get('genHrs')))},{float(check_float_val(x.get('oprHrs')))});"
                            # customerName = "SPI Power" if "spi" in re.sub(r"\s+",'',x.get('customerName')).lower() or "skr" in re.sub(r"\s+",'',x.get('customerName')).lower() else  "KR Wind Energy" if "kr" in re.sub(r"\s+",'',x.get('customerName')).lower() else ''
                            # locNoVal = re.sub(r"\s+",'',x.get('locNo')) if "TP06" not in x.get('locNo') else "TP6"
                            # location_values = location.get(locNoVal)
                            # db_command2=f"insert into spi_windmill_gen_daily_report(gendate,customername,locno,mckwhday,gf,fm,s,u,genhrs,oprhrs,mw,section,site,make) values('{str(x.get('genDate')).split(' ')[0]}','{customerName}','{locNoVal}',{float(check_float_val(x.get('genkwhDay')))},{float(check_float_val(x.get('gf')))},{float(check_float_val(x.get('fm')))},{float(check_float_val(x.get('s')))},{float(check_float_val(x.get('u')))},{float(check_float_val(x.get('genHrs')))},{float(check_float_val(x.get('oprHrs')))},{float(check_float_val(x.get('mw')))},'{location_values[2]}','{location_values[1]}','{location_values[0]}');"

                            # cursor.execute(db_command2)
                            # if "suzlon" in file_name.split('\\')[-1].lower() and "weekly" in file_name.split('\\')[-1].lower():
                            #     db_command1 = f"INSERT INTO suzlon_xl_weekly_hist(gendate,mw,customername,htno,locno,reading_totalimport,reading_06_09am_1,reading_06_09pm_1,reading_09_10pm_1,reading_05_06amand09_06pm_1,reading_10pm_05am_1,reading_totalexport,reading_06_09am_2,reading_06_09pm_2,reading_09_10pm_2,reading_05_06amand09_06pm_2,reading_10pm_05am_2,reading_kvarhimportlag,reading_kvarhimportlead,reading_kvarhexportlag,reading_kvarhexportlead,reading_kvahimportreading,reading_kvahexportreading,reading_powerfactor,reading_percent_kvahimport,reading_monthcumulative,calc_totalimport,calc_06_09am_1,calc_06_09pm_1,calc_09_10pm_1,calc_05_06amand09_06pm_1,calc_10pm_05am_1,calc_totalexport,calc_06_09am_2,calc_06_09pm_2,calc_09_10pm_2,calc_05_06amand09_06pm_2,calc_10pm_05am_2,calc_kvarhimportlag,calc_kvarhimportlead,calc_kvarhexportlag,calc_kvarhexportlead,calc_kvahimportreading,calc_kvahexportreading,calc_powerfactor,calc_percent_kvahimport,calc_monthcumulative) values('{str(x[0]).split(' ')[0]}',{float(check_float_val(x[1]))},'{str(x[2])}','{str(x[3])}','{str(x[4])}',{float(check_float_val(x[5]))},{float(check_float_val(x[6]))},{float(check_float_val(x[7]))},{float(check_float_val(x[8]))},{float(check_float_val(x[9]))},{float(check_float_val(x[10]))},{float(check_float_val(x[14]))},{float(check_float_val(x[15]))},{float(check_float_val(x[16]))},{float(check_float_val(x[17]))},{float(check_float_val(x[18]))},{float(check_float_val(x[19]))},{float(check_float_val(x[11]))},{float(check_float_val(x[12]))},{float(check_float_val(x[20]))},{float(check_float_val(x[21]))},{float(check_float_val(x[13]))},{float(check_float_val(x[22]))},{float(check_float_val(x[23]))},{float(check_float_val(x[24]))},{float(check_float_val(x[25]))},{float(check_float_val(x[26]))},{float(check_float_val(x[27]))},{float(check_float_val(x[28]))},{float(check_float_val(x[29]))},{float(check_float_val(x[30]))},{float(check_float_val(x[31]))},{float(check_float_val(x[32]))},{float(check_float_val(x[36]))},{float(check_float_val(x[37]))},{float(check_float_val(x[38]))},{float(check_float_val(x[39]))},{float(check_float_val(x[40]))},{float(check_float_val(x[41]))},{float(check_float_val(x[33]))},{float(check_float_val(x[34]))},{float(check_float_val(x[42]))},{float(check_float_val(x[43]))},{float(check_float_val(x[35]))},{float(check_float_val(x[44]))},{float(check_float_val(x[45]))},{float(check_float_val(x[46]))});"
                            #     inserted_val['suzlon'].append(db_command1)
                            #     print(len(inserted_val['suzlon']))
                                # customerName = "SPI Power" if "spi" in re.sub(r"\s+",'',x[2]).lower() or "skr" in re.sub(r"\s+",'',x[2]).lower() else  "KR Wind Energy" if "kr" in re.sub(r"\s+",'',x[2]).lower() else ''
                                # locNoVal = re.sub(r"\s+",'',x[4]) if "TP06" not in x[4] else "TP6"
                                # ebkwhday = abs(float(check_float_val(x[35]))) - abs(float(check_float_val(x[26])))
                                # db_command2=f"insert into spi_windmill_gen_daily_report(gendate,customername,locno,mckwhday,gf,fm,s,u,genhrs,oprhrs,ebkwhday,mw,section,site,make) values('{str(x[0]).split(' ')[0]}','{str(x[2])}','{str(x[4])}',)"
                                # db_command2=f"update spi_windmill_gen_daily_report set  ebkwhday={ebkwhday} where gendate='{str(x[0]).split(' ')[0]}' and locno='{locNoVal}' and (customername like '%{' '.join([x for x in str(x[2])])}%' or customername like '%{str(x[2])}%');"
                                # db_command2=f"update spi_windmill_gen_daily_report set ebkwhday={float(check_float_val(ebkwhday))} where gendate='{str(x[0]).split(' ')[0]}' and locno='{locNoVal}' and customername='{customerName}';"
                                # print(db_command2)
                            #     print(db_command2)
                            #     cursor.execute(db_command1)
                                # cursor.execute(db_command2)
                            # if "vestas" in file_name.split('\\')[-1].lower():
                                # db_command1 = f'INSERT into vestas_xl_daily_hist(gendate,mw,customername,htno,locno,readingtakentime,cml_runhrs,cml_genhrs,cml_g0,cml_gen,cml_totalprod,cml_totalimport,cml_06_09am_1,cml_18_21pm_1,cml_21_22pm_1,cml_05_06amand09_18pm_1,cml_22_05am_1,cml_totalexport,cml_06_09am_2,cml_18_21pm_2,cml_21_22pm_2,cml_05_06amand09_18pm_2,cml_22_05am_2,cml_rkvahr_imp,cml_rkvahr_exp,daily_runhrs,daily_genhrs,daily_g0,daily_gen,daily_totalprod,daily_totalimport,daily_06_09am_1,daily_18_21pm_1,daily_21_22pm_1,daily_05_06amand09_18pm_1,daily_22_05am_1,daily_totalexport,daily_06_09am_2,daily_18_21pm_2,daily_21_22pm_2,daily_05_06amand09_18pm_2,daily_22_05am_2,daily_rkvahr_imp,daily_rkvahr_exp,gf,fm,sch,unsch,manualstoppage,readingnotavailable,total,remarks) values("{str(x[0]).split(" ")[0]}",{float(check_float_val(x[1]))},"{x[2]}","{x[3]}","{x[4]}","{(x[5])}",{float(check_float_val(x[6]))},{float(check_float_val(x[7]))},{float(check_float_val(x[8]))},{float(check_float_val(x[9]))},{float(check_float_val(x[10]))},{float(check_float_val(x[11]))},{float(check_float_val(x[12]))},{float(check_float_val(x[13]))},{float(check_float_val(x[14]))},{float(check_float_val(x[15]))},{float(check_float_val(x[16]))},{float(check_float_val(x[17]))},{float(check_float_val(x[18]))},{float(check_float_val(x[19]))},{float(check_float_val(x[20]))},{float(check_float_val(x[21]))},{float(check_float_val(x[22]))},{float(check_float_val(x[23]))},{float(check_float_val(x[24]))},{float(check_float_val(x[25]))},{float(check_float_val(x[26]))},{float(check_float_val(x[27]))},{float(check_float_val(x[28]))},{float(check_float_val(x[29]))},{float(check_float_val(x[30]))},{float(check_float_val(x[31]))},{float(check_float_val(x[32]))},{float(check_float_val(x[33]))},{float(check_float_val(x[34]))},{float(check_float_val(x[35]))},{float(check_float_val(x[36]))},{float(check_float_val(x[37]))},{float(check_float_val(x[38]))},{float(check_float_val(x[39]))},{float(check_float_val(x[40]))},{float(check_float_val(x[41]))},{float(check_float_val(x[42]))},{float(check_float_val(x[43]))},{float(check_float_val(x[44]))},{float(check_float_val(x[45]))},{float(check_float_val(x[46]))},{float(check_float_val(x[47]))},{float(check_float_val(x[48]))},{float(check_float_val(x[49]))},{float(check_float_val(x[50]))},"{x[51]}");'
                                # 'customername', 'text', 'NO', '', NULL, ''

                            db_command1 = f'INSERT into vestas_xl_daily_hist(gendate,mw,customername,htno,site,locno,readingtakentime,cml_runhrs,cml_genhrs,cml_g0,cml_gen,cml_totalprod,cml_totalimport,cml_06_09am_1,cml_18_21pm_1,cml_21_22pm_1,cml_05_06amand09_18pm_1,cml_22_05am_1,cml_totalexport,cml_06_09am_2,cml_18_21pm_2,cml_21_22pm_2,cml_05_06amand09_18pm_2,cml_22_05am_2,cml_rkvahr_imp,cml_rkvahr_exp,daily_runhrs,daily_genhrs,daily_g0,daily_gen,daily_totalprod,daily_totalimport,daily_06_09am_1,daily_18_21pm_1,daily_21_22pm_1,daily_05_06amand09_18pm_1,daily_22_05am_1,daily_totalexport,daily_06_09am_2,daily_18_21pm_2,daily_21_22pm_2,daily_05_06amand09_18pm_2,daily_22_05am_2,daily_rkvahr_imp,daily_rkvahr_exp,gf,fm,sch,unsch,manualstoppage,readingnotavailable,total,remarks) values("{str(x.get("genDate")).split(" ")[0]}","{x.get("mw")}","{x.get("companyName")}","{x.get("htno")}","{x.get("site")}","{x.get("locNo")}","{x.get("reading_taken_time")}",{check_float_val(x.get("cml_run_hr"))},{check_float_val(x.get("cml_gen_hr"))},{check_float_val(x.get("cml_g_0"))},{check_float_val(x.get("cml_gen"))},{check_float_val(x.get("cml_total_prod"))},{check_float_val(x.get("cml_total_import"))},{check_float_val(x.get("cml_06_09_am_1"))},{check_float_val(x.get("cml_18_21_pm_1"))},{check_float_val(x.get("cml_21_22_pm_1"))},{check_float_val(x.get("cml_05_06_am_&_09_18_pm_1"))},{check_float_val(x.get("cml_22_05_am_1"))},{check_float_val(x.get("cml_total_export"))},{check_float_val(x.get("cml_06_09_am_2"))},{check_float_val(x.get("cml_18_21_pm_2"))},{check_float_val(x.get("cml_21_22_pm_2"))},{check_float_val(x.get("cml_05_06_am_&_09_18_pm_2"))},{check_float_val(x.get("cml_22_05_am_2"))},{check_float_val(x.get("cml_rkvahr_imp"))},{check_float_val(x.get("cml_rkvahr_exp"))},{check_float_val(x.get("daily_run_hr"))},{check_float_val(x.get("daily_gen_hr"))},{check_float_val(x.get("daily_g_0"))},{check_float_val(x.get("daily_gen"))},{check_float_val(x.get("Prod"))},{check_float_val(x.get("daily_total_import"))},{check_float_val(x.get("daily_06_09_am_1"))},{check_float_val(x.get("daily_18_21_pm_1"))},{check_float_val(x.get("daily_21_22_pm_1"))},{check_float_val(x.get("daily_05_06_am_&_09_18_pm_1"))},{check_float_val(x.get("daily_22_05_am_1"))},{check_float_val(x.get("daily_total_export"))},{check_float_val(x.get("daily_06_09_am_2"))},{check_float_val(x.get("daily_18_21_pm_2"))},{check_float_val(x.get("daily_21_22_pm_2"))},{check_float_val(x.get("daily_05_06_am_&_09_18_pm_2"))},{check_float_val(x.get("daily_22_05_am_2"))},{check_float_val(x.get("daily_rkvahr_imp"))},{check_float_val(x.get("daily_rkvahr_exp"))},{check_float_val(x.get("gf"))},{check_float_val(x.get("fm"))},{check_float_val(x.get("sch"))},{check_float_val(x.get("unsch"))},{check_float_val(x.get("ms"))},{check_float_val(x.get("readNotAvail"))},{check_float_val(x.get("total"))},"{x.get("remarks")}");'
                            print(db_command1)  
                            ebkwhValue = abs(float(check_float_val(x.get("daily_total_export")))) - abs(float(check_float_val(x.get("daily_total_import"))))
                            customerName = "SPI Power" if "spi" in x[2].lower() or "skr" in x[2].lower() else  "KR Wind Energy" if "kr" in x[2].lower() else ''
                            location_values = location.get(x.get('locNo'))
                            db_command2=f"INSERT into spi_windmill_gen_daily_report(gendate,companyname,locno,mckwhday,gf,fm,sch,unsch,genhrs,oprhrs,ebkwhday,mw,section,site,make) values('{str(x.get('genDate')).split(' ')[0]}','{customerName}','{x.get('locNo')}',{float(check_float_val(x.get('Prod')))},{check_float_val(x.get('gf'))},{check_float_val(x.get('fm'))},{float(check_float_val(x.get('sch')))},{float(check_float_val(x.get('unsch')))},{float(check_float_val(x.get('daily_gen_hr')))},{float(check_float_val(x.get('daily_run_hr')))},{ebkwhValue},{float(check_float_val(x.get('mw')))},'{location_values[1]}','{x.get('site')}','{location_values[0]}');"
                            if any([x.get('cml_run_hr'),x.get('cml_gen_hr'),x.get('cml_g_0'),x.get('cml_gen'),x.get("cml_total_prod")]):
                                # try:
                                cursor.execute(db_command1)
                                cursor.execute(db_command2)
                                # except Exception as e:
                                #     print("\n The exception is : ",e)
                                #     pass                            
                        # cursor.execute(db_command1)
                    connection.commit()

    connection.close()

def read_db_value():
    if connection.is_connected():
        cursor = connection.cursor()
        query = 'select * from vestas_xl_daily_hist;'
        cursor.execute(query)
        data = cursor.fetchall()
        print(len(data))
        sum_val = 0
        for x in data:
            print(x)
        connection.close()
        print("Sum is: ",sum_val)

def load_location_master():
    if connection.is_connected():
        cursor = connection.cursor()
        location_df = pd.read_excel(locatiotn_file)
        for x_i,x in location_df.iterrows():
            location_insert = f"insert into location_master(locno,weghtno,wegcapacitykw,make,site,section,state,company,alloc_company) values('{x.get('location_no')}',{x.get('weight_no')},{x.get('weg_capacity_kw')},'{x.get('make')}','{x.get('site')}','{x.get('section')}','{x.get('state')}','{x.get('company')}','{x.get('alloc_company')}');"
            cursor.execute(location_insert)
        connection.commit()
    connection.close()
            # print(x.get('company')," ",x.get('alloc_company')," ",x.get('location_no')," ",x.get('make')," ",x.get('site')," ",x.get('section')," ",x.get('state'))
insert_into_tables()
# read_db_value()

