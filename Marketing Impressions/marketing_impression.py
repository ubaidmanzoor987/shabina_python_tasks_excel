from datetime import datetime

import pandas as pd
import os
import re
from collections import Counter
from dateutil import rrule

def split_on_slash(fileName):
    year = None
    try:
        year = fileName.split('/')[1].split('0')[1]
    except:
        try:
            year = fileName.split('/')[1]
        except:
            year = ''
    return year

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
    data1 = []
    count = 0
    folder_name = 'Dataset'
    file = 'TV data_2018_Phase_1.xlsx'
    count = same_content(folder_name, file, data1, 0, 1, 2, 3, 6, 9, 10, count)
    file = 'TV data_2018_Phase_2.xlsx'
    count = same_content(folder_name, file, data1, 0, 1, 2, 3, 4, 7, 8, count)
    file = 'TV and Digital_2019.xlsx'
    count = same_content(folder_name, file, data1, 0, 1, 2, 3, 4, 5, 6, count)
    file = 'TV and Digital_2020.xlsx'
    count = same_content(folder_name, file, data1, 0, 2, 3, 6, 12, 13, 14, count)

    df1 = pd.DataFrame(data1)
    df1.to_excel('Cleaned Dataset/Marketing_TV_v4.xlsx', index=False)

def same_content(folder_name, file, data1, r0, r1, col1, col2, col3, col4, col5, count):
    filepath = os.path.join(folder_name, file)
    readexcelfile = None
    if file == "TV data_2018_Phase_1.xlsx":
        readexcelfile = pd.read_excel(filepath, engine='openpyxl', sheet_name=' Aggregate')
    else:
        readexcelfile = pd.read_excel(filepath, engine='openpyxl')
    weekwise_data = {}
    for i, j in readexcelfile.iterrows():
        if not pd.isna(j[0]):
            dt = datetime.strptime(j[0], '%Y/%m/%d')
            d1 = dt.isocalendar()
            week_num = d1[1]
            if not j[2] in weekwise_data:
                weekwise_data[j[2]] = {}

            if week_num in weekwise_data[j[2]]:
                weekwise_data[j[2]][week_num].append(j)
            else:
                weekwise_data[j[2]][week_num] = [j]
    print(weekwise_data)
    for i in weekwise_data.keys():
        for k in weekwise_data[i].keys():
            for j in range(len(weekwise_data[i][k])):
                row = weekwise_data[i][k][j]
                station = row[col1]
                length = row[col2]
                trp = row[col3]
                if trp == "*":
                    trp = 0
                # accumalted_trp = row[col4]
                # if accumalted_trp == "*":
                #     accumalted_trp = 0
                fq1 = row[col5]
                if fq1 == "*":
                    fq1 = 0
                datefromdata = row[r0]
                timefromdata = row[r1]
                date = weekwise_data[i][k][0][0]
                date = datetime.strptime(date, '%Y/%m/%d')
                d1 = date.isocalendar()
                data1.append({})
                data1[count]['Year_Month'] = str(date.year) + '_' + str(date.month).zfill(2)
                data1[count]['Year_Week'] = str(date.year) + '_' + str(d1[1]).zfill(2)
                data1[count]['Year'] = date.year
                data1[count]['Month'] = date.month
                data1[count]['Week'] = d1[1]
                data1[count]['Date'] = datefromdata
                data1[count]['Time'] = timefromdata
                data1[count]['Channel/Station'] = station
                data1[count]['Secs/Length'] = length
                data1[count]['TRP'] = trp
                data1[count]['FQ:1'] = fq1
                data1[count]['Reach'] = trp/fq1
                count += 1
    return count

if __name__ == '__main__':
    main()