import pandas as pd
import os
from collections import Counter
from googletrans import Translator
import datetime

folder_name = '../Translation Data'
file = 'English Translations Dictionary.xlsx'
filepath = os.path.join(folder_name, file)
readexcelfile12 = pd.read_excel(filepath, engine='openpyxl', sheet_name='Translation Dictionary')
translation_dict = {}
for i, j in readexcelfile12.iterrows():
    japanese_language = j[0]
    english_trans = j[1]
    translation_dict[japanese_language] = english_trans

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
    try:
        data1 = data.split(' ')[0]
    except:
        data1 = ''
    return data1

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
    trans = Translator()
    data = []
    count = 0
    folder_name = 'Dataset'
    file = 'In-store Activity Coverage Phase 2 2018.xlsx'
    count = store_data_2018(folder_name, file, data, count,)
    file = 'In-store Activity Coverage_2019.xlsx'
    count = store_data_2019(folder_name, file, data, count)
    file = 'New_In-store Promoter Activity Coverage_2020.xlsx'
    count = store_data_2020(folder_name, file, data, count)
    df1 = pd.DataFrame(data)
    df1.to_excel('Cleaned Dataset/InStore v 4.xlsx', index=False)

def google_translator_my(value):
    if value in translation_dict:
        return translation_dict[value]
    return value

def store_data_2018(folder_name, file, data, count):
    print(file)
    filepath = os.path.join(folder_name, file)
    readexcelfile = pd.read_excel(filepath, engine='openpyxl', sheet_name='Raw')
    for i, j in readexcelfile.iterrows():
        job_number = j[0]
        start_date = split_on_space(str(j[1]))
        end_date = split_on_space(str(j[2]))
        if start_date != 'NaT' and end_date != 'NaT':
            dt = datetime.datetime.strptime(start_date, '%Y-%m-%d')
            di = dt.isocalendar()
            year = di[0]
            week_num = di[1]
            month_num = dt.month
            data.append({})
            data[count]['Year_Month'] = str(year) + '_' + str(month_num).zfill(2)
            data[count]['Year_Week'] = str(year) + '_' + str(week_num).zfill(2)
            data[count]['Year'] = year
            data[count]['Month'] = month_num
            data[count]['Week'] = week_num
            data[count]['Job No'] = job_number
            data[count]['Start Date'] = start_date.replace('-', '/')
            data[count]['End Date'] = end_date.replace('-', '/')
            data[count]['Chain'] = google_translator_my(j[3])
            data[count]['Store'] = google_translator_my(j[4])
            data[count]['Prefecture'] = google_translator_my(j[6])
            data[count]['Reach'] = j[7]
            data[count]['Session'] = j[8]
            data[count]['Phase'] = j[9]
            count +=1
    return count

def store_data_2019(folder_name, file, data, count):
    print(file)
    filepath = os.path.join(folder_name, file)
    readexcelfile = pd.read_excel(filepath, engine='openpyxl', sheet_name='Reach')
    for i, j in readexcelfile.iterrows():
        job_number = j[0]
        start_date = split_on_space(str(j[1]))
        end_date = split_on_space(str(j[2]))
        if start_date != 'NaT' and end_date != 'NaT':
            dt = datetime.datetime.strptime(start_date, '%Y-%m-%d')
            di = dt.isocalendar()
            year = di[0]
            week_num = di[1]
            month_num = di[2]
            data.append({})
            data[count]['Year_Month'] = str(year) + '_' + str(month_num).zfill(2)
            data[count]['Year_Week'] = str(year) + '_' + str(week_num).zfill(2)
            data[count]['Year'] = year
            data[count]['Month'] = month_num
            data[count]['Week'] = week_num
            data[count]['Job No'] = job_number
            data[count]['Start Date'] = start_date.replace('-', '/')
            data[count]['End Date'] = end_date.replace('-', '/')
            data[count]['Chain'] = google_translator_my(j[3])
            data[count]['Store'] = ''
            data[count]['Prefecture'] = google_translator_my(j[5])
            data[count]['Reach'] = j[6]
            data[count]['Session'] = ''
            data[count]['Phase'] = ''
            count +=1

    return count

def store_data_2020(folder_name, file, data, count):
    print(file)
    filepath = os.path.join(folder_name, file)
    readexcelfile = pd.read_excel(filepath, engine='openpyxl', sheet_name='raw data')
    for i, j in readexcelfile.iterrows():
        job_number = j[0]
        start_date = split_on_space(str(j[3]))
        end_date = split_on_space(str(j[5]))
        if start_date != 'NaT' and end_date != 'NaT':
            dt = datetime.datetime.strptime(start_date, '%Y-%m-%d')
            di = dt.isocalendar()
            year = di[0]
            week_num = di[1]
            month_num = di[2]
            cal_session = 0
            print(split_on_col(j[8]))
            if split_on_col(j[7]) == '10' and split_on_col(j[8]) == '18':
                cal_session = 1

            if split_on_col(j[7]) == '11' and split_on_col(j[8]) == '19':
                cal_session = 2
            data.append({})
            data[count]['Year_Month'] = str(year) + '_' + str(month_num).zfill(2)
            data[count]['Year_Week'] = str(year) + '_' + str(week_num).zfill(2)
            data[count]['Year'] = year
            data[count]['Month'] = month_num
            data[count]['Week'] = week_num
            data[count]['Job No'] = job_number
            data[count]['Start Date'] = start_date.replace('-', '/')
            data[count]['End Date'] = end_date.replace('-', '/')
            data[count]['Chain'] = google_translator_my(j[9])
            data[count]['Store'] = google_translator_my(j[10])
            data[count]['Prefecture'] = google_translator_my(j[17])
            data[count]['Reach'] = ''
            data[count]['Session'] = cal_session
            data[count]['Phase'] = ''
            count +=1

    return count

if __name__ == '__main__':
    main()