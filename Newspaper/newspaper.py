from googletrans import Translator
import pandas as pd
import os
from collections import Counter

import datetime


def date_format_check(date):
    d = None
    try:
        d = datetime.datetime.strptime(date, 'yyyy-mm-dd Thh:mm:ss')
    except:
        try:
            d = datetime.datetime.strptime(date, '%d.%m..%Y')
        except:
            d = datetime.datetime.strptime(date, '%m.%d.%Y')
    return d

def split_on_space(data):
    data1 = None
    year = None
    month = None
    try:
        data1 = data.split(' ')[0]
        data2 = data1.split('-')
        year = data2[0]
        month = data2[1]
    except:
        data1 = ''
        year = ''
        month = ''
    return data1, year, month

def split_on_col(data):
    data = str(data)
    data1 = None
    try:
        data1 = data.split(' ')[0].split(':')[0]
    except:
        data1 = ''
    return data1

def my_mode(sample):
    c = Counter(sample)
    return [k for k, v in c.items() if v == c.most_common(1)[0][1]]

def calculate_Avg(array):
    if not array:
        return 0
    if len(array) == 0:
        return 0

    sum = 0
    for i in range(len(array)):
        sum += array[i]
    sum /= len(array)
    return sum

def main():

    data = []
    count = 0
    folder_name = 'Dataset'
    file = 'Newspaper.xlsx'
    count = news_data(folder_name, file, data, count)

    df1 = pd.DataFrame(data)
    df1.to_excel('Cleaned Dataset/Newspaper v 2.xlsx', index=False)

def news_data(folder_name, file, data, count):
    filepath = os.path.join(folder_name, file)
    readexcelfile = pd.read_excel(filepath, engine='openpyxl', sheet_name='Newspapers')
    for i, j in readexcelfile.iterrows():
        start_date, year, month = split_on_space(str(j[4]))
        # print(year, month)

        if start_date != 'NaT' and start_date != '':
            dt = datetime.datetime.strptime(start_date, '%Y-%m-%d')
            di = dt.isocalendar()
            week_num = di[1]
            data.append({})
            data[count]['Year_Month'] = str(year) + '_' + str(month).zfill(2)
            data[count]['Year_Week'] = str(year) + '_' + str(week_num).zfill(2)
            data[count]['Year'] = year
            data[count]['Month'] = month
            data[count]['Week'] = week_num
            data[count]['Newspaper Name'] = j[0]
            data[count]['Newspaper Title'] = j[1]
            data[count]['National vs Regional Name'] = j[2]
            data[count]['Issuance frequency'] = j[3]
            data[count]['Date of publication'] = start_date.replace('-', '/')
            data[count]['Number of circulation'] = j[11]
            data[count]['Overall Number of readers'] = j[12]
            data[count]['At the target Number of readers'] = j[13]
            count +=1

    return count

if __name__ == '__main__':
    main()