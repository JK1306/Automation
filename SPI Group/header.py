import pandas as pd
import re

df = pd.read_excel(r"C:\Users\Jaikishore\Downloads\K R Wind Energy LLP 27082020.xlsx")
for x_i,x in df.iterrows():
    if "date" in str(x[0]).lower():
        df_header = df.iloc[x_i-1:x_i+1].fillna('')
        break
df = df.iloc[x_i+1:]
head_val = ''
header= []
for x_i,x in df_header.iteritems():
    if x.iloc[0]:
        head_val = 'reading_' if 'reading' in x.iloc[0].lower() else 'calc_' if 'calculated' in x.iloc[0].lower() else ''
    if 'date' in x.iloc[1].lower():
        x.iloc[1] = 'genDate'
    if x.iloc[1].lower() == 'mw' or x.iloc[1].lower() == 'site':
        x.iloc[1] = x.iloc[1].lower()
    if 'customer' in x.iloc[1].lower():
        x.iloc[1] = 'companyName'
    if 'htno' in re.sub(r"\s+","",x.iloc[1].lower()):
        x.iloc[1] = 'htno'
    if 'locno' in re.sub(r"\s+","",x.iloc[1].lower()):
        x.iloc[1] = 'locno'
    if 'total' in x.iloc[1].lower() and 'import' in x.iloc[1].lower():
        x.iloc[1] = head_val+"total_import"
    if 'total' in x.iloc[1].lower() and 'export' in x.iloc[1].lower():
        x.iloc[1] = head_val+"total_export"
    if '6.0am' in re.sub(r"\s+","",x.iloc[1].lower()):
        x.iloc[1] = head_val+"6am_to_9am_1" if head_val+"6am_to_9am_1" not in df_header.iloc[1].values else head_val+"6am_to_9am_2"
    if '6.0pm' in re.sub(r"\s+","",x.iloc[1].lower()):
        x.iloc[1] = head_val+"6pm_to_9pm_1" if head_val+"6pm_to_9pm_1" not in df_header.iloc[1].values else head_val+"6pm_to_9pm_2"
    if '9.0pm' in re.sub(r"\s+","",x.iloc[1].lower()):
        x.iloc[1] = head_val+"9pm_to_10pm_1" if head_val+"9pm_to_10pm_1" not in df_header.iloc[1].values else head_val+"9pm_to_10pm_2"
    if '5.0am' in re.sub(r"\s+","",x.iloc[1].lower()):
        print("-----------------",x.iloc[1])
        x.iloc[1] = head_val+"5am_to_6am_and_9am_to_6pm_1" if head_val+"5am_to_6am_and_9am_to_6pm_1" not in df_header.iloc[1].values else head_val+"5am_to_6am_and_9am_to_6pm_2"
    if '10pm' in re.sub(r"\s+","",x.iloc[1].lower()):
        x.iloc[1] = head_val+"10pm_to_5am_1" if head_val+"10pm_to_5am_1" not in df_header.iloc[1].values else head_val+"10pm_to_5am_2"
    if re.search(r"KVA(R|)H",x.iloc[1].upper().strip()):
        if 'import' in x.iloc[1].lower() or 'export' in x.iloc[1].lower():
            if 'lag' in x.iloc[1].lower():
                x.iloc[1] = head_val+"kvarh_import_lag" if 'import' in x.iloc[1].lower() else head_val+"kvarh_export_lag"
            if 'lead' in x.iloc[1].lower():
                x.iloc[1] = head_val+"kvarh_import_lead" if 'import' in x.iloc[1].lower() else head_val+"kvarh_export_lead"
            if 'reading' in x.iloc[1].lower():
                x.iloc[1] = head_val+"kvah_import_reading" if 'import' in x.iloc[1].lower() else head_val+"kvah_export_reading"
            if "%" in x.iloc[1]:
                x.iloc[1] = head_val+"percent_kvarh_import"
    if 'month' in x.iloc[1].lower() and 'cumulative' in x.iloc[1].lower():
        x.iloc[1] = head_val+"month_cml"
    if 'power' in x.iloc[1].lower() and 'factor' in x.iloc[1].lower():
        x.iloc[1] = head_val+"power_factor"
df_header.iloc[1].unique()