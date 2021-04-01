import pandas as pd
from datetime import datetime
from datetime import timedelta
import re

index = pd.read_excel('Brand Tracker/20210202 JP_Weekly brand tracker 2019 by Nov_Shortlisted Batch 2 (Price_Worth).xlsx', sheet_name='Index')
index = index[:-1]
index['Table'] = index['Table'].astype(int)
index['Title'] = index['Title'].str.replace("-","")
#index = index[index['Batch 1'] == 1]

year = 2019
dfs = []
print(list(index['Table']))
for inde in list(index['Table']):

    ind = inde
    if ind == 1069:
        continue
    a = index.loc[ind, 'Price/Worth Extracts']
    b = index.loc[ind, 'Table']
    print(a, b)
    if index.loc[ind, 'Price/Worth Extracts'] == 1:
        print('Executing for Table:', index.loc[ind, 'Table'], 'Title:', index.loc[ind , 'Title'])

        df = pd.read_excel(
            'Brand Tracker/20210202 JP_Weekly brand tracker 2019 by Nov_Shortlisted Batch 2 (Price_Worth).xlsx',
            sheet_name='Table' + str(index.loc[ind, 'Table']), skiprows=7)
        df = df.dropna()
        df = df.rename(columns={ df.columns[0]: "Type" })
        df_melt = pd.melt(df, id_vars=['Type'])
        df_melt.columns = ['Type', 'Week', 'Value']

        df_melt = df_melt[['Type', 'Week', 'Value']]
        df_melt.loc[df_melt['Value'] == '-', 'Value'] = ''

        if year == 2020:
            df_melt['Year'] = df_melt['Week'].str.split(' ').str[0]
            df_melt['WeekNo'] = df_melt['Week'].str.split(' ').str[1]
            df_melt['WeekDate'] = df_melt['Year'] + df_melt['WeekNo']
            df_melt['Date'] = df_melt.apply(lambda row: datetime.strptime(row['WeekDate'] + '-1', "%YW%W-%w") - timedelta(days=6), axis=1)
            df_melt['Month'] = df_melt['Date'].dt.month

        elif year == 2019:

            df_melt['Year'] = 2019
            df_melt['WeekNo'] = df_melt['Week'].str.split(' ').str[1].str.split('/').str[1]
            df_melt['WeekRange'] = df_melt['Week'].str.split(' ').str[1].str.split('/').str[1]
            df_melt['Month'] = df_melt['WeekRange'].str.split('-').str[1]
            df_melt['Day'] = df_melt['WeekRange'].str.split('-').str[0]
            df_melt['Date'] = pd.to_datetime(df_melt[['Year', 'Month', 'Day']])
            df_melt['Week'] = df_melt.apply(lambda row: row['Date'].isocalendar()[1], axis=1)

        df_melt['Year_Week'] = df_melt.apply(lambda row: str(row['Date'].isocalendar()[0]) + '_' + str(row['Date'].isocalendar()[1]).zfill(2), axis=1)
        df_melt['Year_Month'] = pd.to_datetime(df_melt['Date']).dt.strftime('%Y_%m')
        df_melt['Week'] = df_melt.apply(lambda row: row['Date'].isocalendar()[1], axis=1)
        df_melt['Value'] =  df_melt['Value'].apply(pd.to_numeric, errors='coerce')
        df_melt['Value'] = df_melt['Value'].fillna(0)
        df_melt = df_melt[['Year', 'Month', 'Week', 'Year_Week', 'Year_Month', 'Type', 'Value']]

        df_melt.columns = ['Year','Month', 'Week', 'Year_Week', 'Year_Month', 'Row_Title', 'Value']
        df_melt['Full_Title'] = index.loc[ind, 'Title']
        df_melt['Table'] = 'Table' + str(b)

        r1 = re.sub("[\(\[].*?[\)\]]", "", index.loc[ind, 'Title'])
        r2 = re.sub(r"^\s+|\s+$", '', r1).strip().split('  ')
        r3 = [item.strip() for item in r2 if len(item) >= 1]
        df_melt['Title_Core'] = r3[0]
        if len(r3) > 1:
            df_melt['Title_Type'] = r3[1]
        else:
            df_melt['Title_Type'] = ''
        #print(df_melt)

        dfs.append(df_melt)
    else:
        continue

index = pd.read_excel(
    'Brand Tracker/20210202 JP_Weekly brand tracker 2020 by Nov_Shortlisted Batch 2 (Price_Worth).xlsx',
    sheet_name='Index')
index = index[:-1]
index['Table'] = index['Table'].astype(int)
index['Title'] = index['Title'].str.replace("-", "")
# index = index[index['Batch 1'] == 1]

year = 2020
print(list(index['Table']))
for inde in list(index['Table']):
    ind = inde
    if ind == 443:
        continue
    if index.loc[ind, 'Price/Worth'] == 1:
        print('ind', ind)
        print('Executing for Table:', index.loc[ind, 'Table'], 'Title:', index.loc[ind , 'Title'])
        table_number = index.loc[ind, 'Table']
        df = pd.read_excel(
            'Brand Tracker/20210202 JP_Weekly brand tracker 2020 by Nov_Shortlisted Batch 2 (Price_Worth).xlsx',
            sheet_name='Table' + str(index.loc[ind , 'Table']), skiprows=7)
        df = df.dropna()
        df = df.rename(columns={df.columns[0]: "Type"})
        df_melt = pd.melt(df, id_vars=['Type'])
        df_melt.columns = ['Type', 'Week', 'Value']

        df_melt = df_melt[['Type', 'Week', 'Value']]
        df_melt.loc[df_melt['Value'] == '-', 'Value'] = ''
        if year == 2020:
            df_melt['Year'] = df_melt['Week'].str.split(' ').str[0]
            df_melt['WeekNo'] = df_melt['Week'].str.split(' ').str[1]
            df_melt['WeekDate'] = df_melt['Year'] + df_melt['WeekNo']
            df_melt['Date'] = df_melt.apply(
                lambda row: datetime.strptime(row['WeekDate'] + '-1', "%YW%W-%w") - timedelta(days=6), axis=1)
            df_melt['Month'] = df_melt['Date'].dt.month

        df_melt['Year_Week'] = df_melt.apply(
            lambda row: str(row['Date'].isocalendar()[0]) + '_' + str(row['Date'].isocalendar()[1]).zfill(2), axis=1)
        df_melt['Year_Month'] = pd.to_datetime(df_melt['Date']).dt.strftime('%Y_%m')
        df_melt['Week'] = df_melt.apply(lambda row: row['Date'].isocalendar()[1], axis=1)
        df_melt['Value'] = df_melt['Value'].apply(pd.to_numeric, errors='coerce')
        df_melt['Value'] = df_melt['Value'].fillna(0)
        df_melt = df_melt[['Year', 'Month', 'Week', 'Year_Week', 'Year_Month', 'Type', 'Value']]

        df_melt.columns = ['Year', 'Month', 'Week', 'Year_Week', 'Year_Month', 'Row_Title', 'Value']
        df_melt['Full_Title'] = index.loc[ind, 'Title']
        df_melt['Table'] = 'Table' + str(table_number)

        r1 = re.sub("[\(\[].*?[\)\]]", "", index.loc[ind, 'Title'])
        r2 = re.sub(r"^\s+|\s+$", '', r1).strip().split('  ')
        r3 = [item.strip() for item in r2 if len(item) >= 1]
        df_melt['Title_Core'] = r3[0]
        if len(r3) > 1:
            df_melt['Title_Type'] = r3[1]
        else:
            df_melt['Title_Type'] = ''
        # print(df_melt)

        dfs.append(df_melt)
    else:
        continue

bt = pd.concat(dfs, axis = 0)
bt = bt[['Table', 'Year', 'Month', 'Week', 'Year_Week', 'Year_Month', 'Row_Title', 'Full_Title', 'Title_Core', 'Title_Type', 'Value']]

mapping_table = pd.read_excel('Brand Tracker/20210202_Reference Table for Title_Type Tagging v02.xlsx', engine='openpyxl', sheet_name='Reference Table')
mapping_table.columns = ['Title_Type Tagging', 'Title_Type']

bt = pd.merge(bt, mapping_table, how="left", on="Title_Type")

bt.to_csv('BrandTracker.csv', index=0, encoding='utf-8-sig')

