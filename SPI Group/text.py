import pandas as pd
import os
df = pd.read_excel(os.path.join(os.path.dirname(__file__),'Suzlon_Weekly_31072020.xls'))
print(df)