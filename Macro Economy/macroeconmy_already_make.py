import pandas as pd
import datetime
import numpy as np

pd.set_option('display.max_rows', 500)
pd.set_option('display.max_columns', 500)
pd.set_option('display.width', 150)

now = datetime.datetime.now()
lower_bound_year = now.year - 3
files = ['JP_Consumer price index.csv' , 'JP Employment ratio_python.xlsx', 'JP_Retail sales value.csv', 'JP_Total Cash earning.csv', 'JP_Two plus household expenditure.xlsx']
for file in files:
    df = pd.read_excel(file, sheet_name = 'Seasonally adjusted value', skiprows = 6)
    row_number_start = df[df['Unnamed: 0'] == lower_bound_year].index[0] - 1
    df = df[row_number_start : row_number_start + 36]

    df = df.drop(['Unnamed: 1','Unnamed: 3'], axis = 1)

    df.columns = ['Year','Month','Both sexes','Male','Female','Both sexes','Male','Female','Both sexes','Male','Female','Both sexes','Male','Female','Both sexes','Male','Female','Both sexes','Male','Female']
    
    df_labour_force = df.iloc[:, 0:5]
    df_labour_force['Category'] = 'Labour Force'

    df2_Employed_Person = pd.concat([df.iloc[:, 0:2], df.iloc[:,5:8]], axis = 1)
    df2_Employed_Person['Category'] = 'Employed Person'

    df3_Employee = pd.concat([df.iloc[:, 0:2], df.iloc[:,8:11]], axis = 1)
    df3_Employee['Category'] = 'Employee'

    df4_UnEmployed_person = pd.concat([df.iloc[:, 0:2], df.iloc[:,11:14]], axis = 1)
    df4_UnEmployed_person['Category'] = 'UnEmployed Person'

    df5_NotinLabourForce = pd.concat([df.iloc[:, 0:2], df.iloc[:,14:17]], axis = 1)
    df5_NotinLabourForce['Category'] = 'Not in Labour Force'

    df6_Unemployement_rate = pd.concat([df.iloc[:, 0:2], df.iloc[:,17:20]], axis = 1)
    df6_Unemployement_rate['Category'] = 'Unemployment rate  (percent)'

    All_df = pd.concat([df_labour_force, df2_Employed_Person, df3_Employee, df4_UnEmployed_person, df5_NotinLabourForce, df6_Unemployement_rate])
    All_df.loc[All_df['Month'] != 'Feb.', 'Year']= np.nan

    All_df['Year'] = All_df['Year'].shift(-1)
    All_df['Year'] = All_df['Year'].ffill()

    All_df.to_excel("MacroEconomy.xlsx", index = False)