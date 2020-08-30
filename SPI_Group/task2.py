import pandas as pd
import os
from openpyxl import Workbook
import openpyxl
import glob
import mysql.connector
from mysql.connector import Error

def check_float_val(val):
    return val if val else 0.0

# C:\Users\Jaikishore\Downloads\Fw__Final_Format_for_Daily_MIS\Daily MIS Report -1.xls
# files = glob.glob(r'C:\Users\Jaikishore\Documents\RAP\BOTTestLoad\*')
files = glob.glob(r'C:\Users\Jaikishore\Downloads\Fw__Final_Format_for_Daily_MIS\*')
for x in files:
    if "suzlon" in x.split('\\')[-1].lower() or 'vestas' in x.split('\\')[-1].lower() or "location" in x.split('\\')[-1].lower():
        dest_path = x.split("\\")
        dest_path.insert(-1,"OUPUT")
        os.makedirs("\\".join(dest_path[:-1]),exist_ok=True)
        dest_path = "\\".join(dest_path).replace('xls','xlsx') if "xlsx" not in x else "\\".join(dest_path)
        if "suzlon" in x.split('\\')[-1].lower() and "weekly" in x.split('\\')[-1].lower() or "vestas" in x.split('\\')[-1].lower():
            check_val = "Gen Date" if "suzlon" in x.split('\\')[-1].lower() else "Date"
            data = pd.read_excel(x,header=None)
            skip_val = data[data[0]==check_val].index[0]
            excel_data = data.iloc[skip_val:]
            header_val = excel_data.iloc[:1]
            excel_data = excel_data.iloc[1:]
            excel_data.columns=header_val.iloc[0]
            excel_data.to_excel(dest_path,index=False)
        # elif "location" in x.split('\\')[-1].lower():

        else:
            df=pd.read_excel(x,header=0)
            df.to_excel(dest_path,index=False)
dest_path=dest_path.split("\\")
dest_path="\\".join(dest_path[:-1])
files_path = glob.glob(dest_path+"\\*")
try:
    connection  = mysql.connector.connect(host="139.59.46.107",
                                    port=6603,
                                    database="spi_group_windmill_data",
                                    user="root",
                                    password="root")
    if connection.is_connected():
        cursor = connection.cursor()
        for x in files_path:
            if "suzlon" in x.split('\\')[-1].lower() and "weekly" not in x.split('\\')[-1].lower():
            # if 'location' in x.split('\\')[-1].lower():
                print(x)
                wb = openpyxl.load_workbook(x)
                shet_obj = wb.active
                customer_name = "M/s. KR WIND ENERGY LLP"
                for x in shet_obj:
                    # INTIAL
                    if x[0].value and "Date" not in str(x[0].value) and "intial" not in str(x[0].value).lower() and "total" not in str(x[0].value).lower() and 'locationno' not in str(x[0].value).lower():
                        if "suzlon" in x.split('\\')[-1].lower() and "weekly" not in x.split('\\')[-1].lower():
                            db_command1 = f"INSERT INTO suzlon_xl_daily_hist(gendate,customername,state,site,section,mw,locno,genkwhday,genkwhmtd,genkwhytd,plfday,plfmtd,plfytd,mcavail,gf,fm,s,u,nor,genhrs,oprhrs) VALUES('{str(x[0].value).split(' ')[0]}','{x[1].value}','{x[2].value}','{x[3].value}','{x[4].value}',{float(check_float_val(x[5].value))},'{x[6].value}',{float(check_float_val(x[7].value))},{float(check_float_val(x[8].value))},{float(check_float_val(x[9].value))},{float(check_float_val(x[10].value))},{float(check_float_val(x[11].value))},{float(check_float_val(x[12].value))},{float(check_float_val(x[13].value)) if type(x[13].value)!=str else 0.0},{float(check_float_val(x[14].value))},{float(check_float_val(x[15].value))},{float(check_float_val(x[16].value))},{float(check_float_val(x[17].value))},{float(check_float_val(x[18].value))},{float(check_float_val(x[19].value))},{float(check_float_val(x[20].value))});"
                            db_command2=f"INSERT into spi_windmill_gen_daily_report(gendate,customername,locono,genkwhday,gf,fm,s,u,genhrs,oprhrs,netkwhday) values('{str(x[0].value).split(' ')[0]}','{x[1].value}','{x[6].value}',{float(check_float_val(x[7].value))},{float(check_float_val(x[14].value))},{float(check_float_val(x[15].value))},{float(check_float_val(x[16].value))},{float(check_float_val(x[17].value))},{float(check_float_val(x[19].value))},{float(check_float_val(x[20].value))});"
                            cursor.execute(db_command1)
                            cursor.execute(db_command2)

                        if "suzlon" in x.split('\\')[-1].lower() and "weekly" in x.split('\\')[-1].lower():
                            db_command = f"INSERT INTO suzlon_xl_weekly_hist(gendate,mw,customername,htno,locno,reading_totalimport,reading_06_09am_1,reading_06_09pm_1,reading_09_10pm_1,reading_05_06amand09_06pm_1,reading_10pm_05am_1,reading_totalexport,reading_06_09am_2,reading_06_09pm_2,reading_09_10pm_2,reading_05_06amand09_06pm_2,reading_10pm_05am_2,reading_kvarhimportlag,reading_kvarhimportlead,reading_kvarhexportlag,reading_kvarhexportlead,reading_kvahimportreading,reading_kvahexportreading,reading_powerfactor,reading_percent_kvahimport,reading_monthcumulative,calc_totalimport,calc_06_09am_1,calc_06_09pm_1,calc_09_10pm_1,calc_05_06amand09_06pm_1,calc_10pm_05am_1,calc_totalexport,calc_06_09am_2,calc_06_09pm_2,calc_09_10pm_2,calc_05_06amand09_06pm_2,calc_10pm_05am_2,calc_kvarhimportlag,calc_kvarhimportlead,calc_kvarhexportlag,calc_kvarhexportlead,calc_kvahimportreading,calc_kvahexportreading,calc_powerfactor,calc_percent_kvahimport,calc_monthcumulative) values('{str(x[0].value).split(' ')[0]}',{float(check_float_val(x[1].value))},'{str(x[2].value)}','{str(x[3].value)}','{str(x[4].value)}',{float(check_float_val(x[5].value))},{float(check_float_val(x[6].value))},{float(check_float_val(x[7].value))},{float(check_float_val(x[8].value))},{float(check_float_val(x[9].value))},{float(check_float_val(x[10].value))},{float(check_float_val(x[14].value))},{float(check_float_val(x[15].value))},{float(check_float_val(x[16].value))},{float(check_float_val(x[17].value))},{float(check_float_val(x[18].value))},{float(check_float_val(x[19].value))},{float(check_float_val(x[11].value))},{float(check_float_val(x[12].value))},{float(check_float_val(x[20].value))},{float(check_float_val(x[21].value))},{float(check_float_val(x[13].value))},{float(check_float_val(x[22].value))},{float(check_float_val(x[23].value))},{float(check_float_val(x[24].value))},{float(check_float_val(x[25].value))},{float(check_float_val(x[26].value))},{float(check_float_val(x[27].value))},{float(check_float_val(x[28].value))},{float(check_float_val(x[29].value))},{float(check_float_val(x[30].value))},{float(check_float_val(x[31].value))},{float(check_float_val(x[32].value))},{float(check_float_val(x[36].value))},{float(check_float_val(x[37].value))},{float(check_float_val(x[38].value))},{float(check_float_val(x[39].value))},{float(check_float_val(x[40].value))},{float(check_float_val(x[41].value))},{float(check_float_val(x[33].value))},{float(check_float_val(x[34].value))},{float(check_float_val(x[42].value))},{float(check_float_val(x[43].value))},{float(check_float_val(x[35].value))},{float(check_float_val(x[44].value))},{float(check_float_val(x[45].value))},{float(check_float_val(x[46].value))});"
                            cursor.execute(db_command)

                        if "vestas" in x.split('\\')[-1].lower():
                            db_command1 = f"INSERT into vestas_xl_daily_hist(gendate,mw,customername,htno,locno,readingtakentime,cml_runhrs,cml_genhrs,cml_g0,cml_gen,cml_totalprod,cml_totalimport,cml_06_09am_1,cml_18_21pm_1,cml_21_22pm_1,cml_05_06amand09_18pm_1,cml_22_05am_1,cml_totalexport,cml_06_09am_2,cml_18_21pm_2,cml_21_22pm_2,cml_05_06amand09_18pm_2,cml_22_05am_2,cml_rkvahr_imp,cml_rkvahr_exp,daily_runhrs,daily_genhrs,daily_g0,daily_gen,daily_totalprod,daily_totalimport,daily_06_09am_1,daily_18_21pm_1,daily_21_22pm_1,daily_05_06amand09_18pm_1,daily_22_05am_1,daily_totalexport,daily_06_09am_2,daily_18_21pm_2,daily_21_22pm_2,daily_05_06amand09_18pm_2,daily_22_05am_2,daily_rkvahr_imp,daily_rkvahr_exp,gf,fm,sch,unsch,manualstoppage,readingnotavailable,total,remarks) values('{str(x[0].value).split(' ')[0]}',{float(check_float_val(x[1].value))},'{customer_name}','{x[2].value}','{x[3].value}','{x[4].value}',{float(check_float_val(x[5].value))},{float(check_float_val(x[6].value))},{float(check_float_val(x[7].value))},{float(check_float_val(x[8].value))},{float(check_float_val(x[9].value))},{float(check_float_val(x[10].value))},{float(check_float_val(x[11].value))},{float(check_float_val(x[12].value))},{float(check_float_val(x[13].value))},{float(check_float_val(x[14].value))},{float(check_float_val(x[15].value))},{float(check_float_val(x[16].value))},{float(check_float_val(x[17].value))},{float(check_float_val(x[18].value))},{float(check_float_val(x[19].value))},{float(check_float_val(x[20].value))},{float(check_float_val(x[21].value))},{float(check_float_val(x[22].value))},{float(check_float_val(x[23].value))},{float(check_float_val(x[24].value))},{float(check_float_val(x[25].value))},{float(check_float_val(x[26].value))},{float(check_float_val(x[27].value))},{float(check_float_val(x[28].value))},{float(check_float_val(x[29].value))},{float(check_float_val(x[30].value))},{float(check_float_val(x[31].value))},{float(check_float_val(x[32].value))},{float(check_float_val(x[33].value))},{float(check_float_val(x[34].value))},{float(check_float_val(x[35].value))},{float(check_float_val(x[36].value))},{float(check_float_val(x[37].value))},{float(check_float_val(x[38].value))},{float(check_float_val(x[39].value))},{float(check_float_val(x[40].value))},{float(check_float_val(x[41].value))},{float(check_float_val(x[42].value))},{float(check_float_val(x[43].value))},{float(check_float_val(x[44].value))},{float(check_float_val(x[45].value))},{float(check_float_val(x[46].value))},{float(check_float_val(x[47].value))},{float(check_float_val(x[48].value))},{float(check_float_val(x[49].value))},'{x[50].value}');"
                            genkwhValue = float(check_float_val(x[16].value)) - float(check_float_val(x[10].value))
                            db_command2=f"INSERT into spi_windmill_gen_daily_report(gendate,customername,locono,genkwhday,gf,fm,s,u,genhrs,oprhrs,netkwhday) values('{str(x[0].value).split(' ')[0]}','{customer_name}','{x[3].value}',{genkwhValue},{float(check_float_val(x[43].value))},{float(check_float_val(x[44].value))},{float(check_float_val(x[45].value))},{float(check_float_val(x[46].value))},{float(check_float_val(x[6].value))},{float(check_float_val(x[5].value))});"
                            cursor.execute(db_command1)
                            cursor.execute(db_command2)

                        if "location" in x.split('\\')[-1].lower():
                            db_command=f"INSERT into location_master(locno,weghtno,wegcapacitykw,make) values('{str(x[0].value)}',{int(x[1].value)},{int(x[2].value)},'{x[3].value}');"
                            cursor.execute(db_command)
                        print("---------",db_command)
        connection.commit()
        cursor.close()
except Exception as e:
    print("The error is \t:",e)   