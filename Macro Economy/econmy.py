from datetime import datetime

import pandas as pd
import os
import re
from dateutil import rrule

def split_on_dot(fileName):
    year = None
    try:
        year = fileName.split('.')[1]
    except:
        try:
            year = fileName.split('-')[1]
        except:
            year = ''
    return year

def main():
    folder_name = 'Dataset'
    files = ['JP_Employment ratio.xlsx', 'JP_Consumer price index.csv', 'JP Employment ratio_python.xlsx', 'JP_Retail sales value.csv', 'JP_Total Cash earning.csv', 'JP_Two plus household expenditure.xlsx']
    data = [{} for i in range(36)]
    for file in files:
        filepath = os.path.join(folder_name, file)
        file_to_read_extension = file.split('.')[1]
        if file_to_read_extension == 'xlsx':
            if file == 'JP_Employment ratio.xlsx':
                print(file)
                readexcelfile = pd.read_excel(filepath, engine='openpyxl', sheet_name='Seasonally adjusted value')
                count = 0
                for i,j in readexcelfile.iterrows():
                    if j[0] == 2018 or j[0] == 2019 or j[0] == 2020:
                        labourbothesexes = j[4]
                        labourmale = j[5]
                        labourfemale = j[6]
                        employeedbothsexes = j[7]
                        employeedmale = j[8]
                        employeedfemale = j[9]
                        employebothsexes = j[10]
                        employemale = j[11]
                        employefemale = j[12]
                        unemployebothsexes = j[13]
                        unemployebothmale = j[14]
                        unemployebothfemale = j[15]
                        notinlabourbothsexes = j[16]
                        notinlabourbothmale = j[17]
                        notinlabourbothfemale = j[18]
                        unemployeebothpercent = j[19]
                        unemployeemalepercent = j[20]
                        unemployeefemalepercent = j[21]
                        data[count]['year_week'] = str(j[0]) + '_ 1'
                        data[count]['year_month'] = str(j[0]) + '_' + re.findall('\d+', j[1] )[0]
                        data[count]['Year'] = str(j[0])
                        data[count]['Month'] = re.findall('\d+', j[1] )[0]
                        data[count]['Week'] = '1'
                        data[count]['FileName'] = ''
                        data[count]['Date'] = ''
                        data[count]['Labour Force (Both Sexes)'] = labourbothesexes
                        data[count]['Labour Force (Male)'] = labourmale
                        data[count]['Labour Force (Female)'] = labourfemale
                        data[count]['Employeed Person (Both Sexes)'] = employeedbothsexes
                        data[count]['Employeed Person (Male)'] = employeedmale
                        data[count]['Employeed Person (Female)'] = employeedfemale
                        data[count]['Employee (Both Sexes)'] = employebothsexes
                        data[count]['Employee (Male)'] = employemale
                        data[count]['Employee (Female)'] = employefemale
                        data[count]['UnEmployeed Person (Both Sexes)'] = unemployebothsexes
                        data[count]['UnEmployeed Person (Male)'] = unemployebothmale
                        data[count]['UnEmployeed Person (Female)'] = unemployebothfemale
                        data[count]['Not in Labour Force (Both Sexes)'] = notinlabourbothsexes
                        data[count]['Not in Labour Force (Male)'] = notinlabourbothmale
                        data[count]['Not in Labour Force (Female)'] = notinlabourbothfemale
                        data[count]['UnEmployment rate Percent (Both Sexes)'] = unemployeebothpercent
                        data[count]['UnEmployment rate Percent (Male)'] = unemployeemalepercent
                        data[count]['UnEmployment rate Percent (Female)'] = unemployeefemalepercent
                        count += 1
            if file == 'JP_Two plus household expenditure.xlsx':
                    print(file)
                    count = 0
                    readexcelfilece = pd.read_excel(filepath, engine='openpyxl')
                    for i, j in readexcelfilece.iterrows():
                        calyear = split_on_dot(j[0])
                        if calyear == '2018' or calyear == '2019' or calyear == '2020':
                            consumptioexpenditure = j[1]
                            consumptioexpenditureratio = j[2]
                            data[count]['Consumption expenditures for two-or-more-person households【yen】'] = consumptioexpenditure
                            data[count]['vs. Ya'] = consumptioexpenditureratio
                            count +=1
        elif file_to_read_extension == 'csv':
            if file == 'JP_Consumer price index.csv':
                print(file)
                readcsvfilejp = pd.read_csv(filepath)
                count = 0
                for i, j in readcsvfilejp.iterrows():
                    calyear = split_on_dot(j[0])
                    if calyear == '2018' or calyear == '2019' or calyear == '2020':
                        consumerpriceindex = j[1]
                        consumerpriceindexwithratio = j[2]
                        data[count]['Consumer Price Index (All items) 2015 base'] = consumerpriceindex
                        data[count]['vs Ya%'] = consumerpriceindexwithratio
                        count +=1
            if file == 'JP_Retail sales value.csv':
                print(file)
                readcsvfilerf = pd.read_csv(filepath)
                count = 0
                for i, j in readcsvfilerf.iterrows():
                    calyear = split_on_dot(j[0])
                    if calyear == '2018' or calyear == '2019' or calyear == '2020':
                        retailsalesvalue = j[1]
                        retailsalesvalueratio = j[2]
                        data[count]['Retail sales value (Nominal)【billion yen】'] = retailsalesvalue
                        data[count]['vs YA%'] = retailsalesvalueratio
                        count +=1
            if file == 'JP_Total Cash earning.csv':
                print(file)
                count =0
                readcsvfilece = pd.read_csv(filepath)
                for i, j in readcsvfilece.iterrows():
                    calyear = split_on_dot(j[1])
                    if calyear == '2018' or calyear == '2019' or calyear == '2020':
                        totalcashearning = j[2]
                        totalcashearningratio = j[3]
                        data[count]['Total cash earnings【yen】'] = totalcashearning
                        data[count]['vs. YA%'] = totalcashearningratio
                        count +=1

    start = datetime(year=2018, month=1, day=1)
    end = datetime(year=2020, month=12, day=31)
    data1 = []
    for dt in rrule.rrule(rrule.WEEKLY, dtstart=start, until=end):
        month = dt.month
        year = dt.year
        print(month, year)
        di = dt.isocalendar()
        print(di)
        week_num = di[1]
        for row in data:
            if row['year_month'] == str(year) + '_' + str(month):
                row1 = row.copy()
                row1['year_week'] = str(year) + '_' + str(week_num)
                row1['Week'] = week_num
                data1.append(row1)

    df1 = pd.DataFrame(data1)
    df1.to_excel('Cleaned Dataset/macroEconomy.xlsx', index=False)

if __name__ == '__main__':
    main()