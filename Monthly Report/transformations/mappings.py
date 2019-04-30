import pandas as pd
import os

df_map = pd.read_excel(r'C:/Users/kocherlp/Bristol-Myers Squibb (O365-D)/Ramnauth, Kevin - Projects/telephony-data/data/Monthly Telephony High Level Script Automation_input_output.xlsx', sheet_name=None)

dfmap_fieldsci = df_map['Sheet2'].set_index('Call Type Name').to_dict()
dfmap_otc_amer = df_map['Sheet6'].set_index('Call Type Name').to_dict()
dfmap_otc_emea = df_map['Sheet4'].set_index('Call Type Name').to_dict()