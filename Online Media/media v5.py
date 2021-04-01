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

def split_on_hyphen(dt):
    data1 = None
    year = None
    month = None
    try:
        data1 = dt.split(' ')[0].split('-')
        year = data1[0]
        month = data1[1]
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
    file = 'Online Ads data 2018 Phase 1.xlsx'
    count = media_data_2018(folder_name, file, data, count,)
    count = media_data_2018_2(folder_name, file, data, count,)
    file = 'Digital All 2018 Phase 2.xlsx'
    count = media_data_phase_2_2018(folder_name, file, data, count,)
    file = 'Online Media 2019.xlsx'
    count = media_data_2019(folder_name, file, data, count,)
    file = 'TV and Digital 2020.xlsx'
    count = media_data_2020(folder_name, file, data, count,)

    for row in data:
        print(row['Platform'])
        if row['Platform'] == 'COOKPAD' or row['Platform'] == 'COOKPAD　':
            print(row['Platform'])
            row['Platform'] = 'Cookpad'
    df1 = pd.DataFrame(data)
    df1.to_excel('Cleaned Dataset/Media v 4.xlsx', index=False)


def media_data_2018_2(folder_name, file, data, count):
    print('media_data_2018_2')
    filepath = os.path.join(folder_name, file)
    readexcelfile = pd.read_excel(filepath, engine='openpyxl', sheet_name='YouTube_Online delivery by site')
    # readexcelfile.to_excel('readed data.xlsx')
    weekwise_data = {}
    prevPlatform = None
    prevDevice = None
    prevMenu = None
    prevFormat = None
    for i, j in readexcelfile.iterrows():
        # print(
        dt = datetime.datetime.strptime('2018/5/1', '%Y/%m/%d')
        count1 = 0
        platform1 = j['Platform'] if j['Platform'] != '' and not pd.isna(j['Platform']) else prevPlatform
        if platform1 == 'COOKPAD':
            platform1 = 'Cookpad'
        device = j['device'] if j['device'] != '' and not pd.isna(j['device'])  else prevDevice
        menu = j['Menu'] if j['Menu'] != '' and not pd.isna(j['Menu']) else prevMenu
        format = j['Ads Fromat'] if j['Ads Fromat'] != '' and not pd.isna(j['Ads Fromat']) else prevFormat
        if device == 'Total':
            prevDevice = device
            continue
        imp_or_click = j['KPI']
        platform = platform1 + device + menu
        for k in range(5,58):
            dt1 = dt + datetime.timedelta(days=count1)
            d1 = dt1.isocalendar()
            week_num = d1[1]
            row = {'date': str(dt1), 'platform': platform1, 'device': device, 'menu': menu, 'format': format, imp_or_click: j[k], 'week_num': week_num}

            if not platform in weekwise_data:
                weekwise_data[platform] = {}

            if week_num in weekwise_data[platform]:
                weekwise_data[platform][week_num].append(row)
            else:
                weekwise_data[platform][week_num] = [row]

            count1 += 1
        prevPlatform = platform1
        prevDevice = device
        prevMenu = menu
        prevFormat = format

    # print(weekwise_data)
    for i in weekwise_data.keys():
        for k in weekwise_data[i].keys():
            clicks = 0
            impressions = 0
            for j in range(len(weekwise_data[i][k])):
                row = weekwise_data[i][k][j]

                if 'imp' in row and not pd.isna(row['imp']):
                    impressions += row['imp']
                elif 'clicks' in row and not pd.isna(row['clicks']):
                    clicks += row['clicks']
            data1, year, month = split_on_hyphen(row['date'])
            data.append({})
            format = None
            if not pd.isna(row['format']):
                format = row['format']
            else:
                format = ''
            data[count]['Year_Week'] = str(year) + '_' + str(row['week_num']).zfill(2)
            data[count]['Year_Month'] = str(year) + '_' + str(month).zfill(2)
            data[count]['Year'] = str(year)
            data[count]['Month'] = month
            data[count]['Week'] = row['week_num']
            data[count]['Platform'] = row['platform']
            data[count]['Device'] = row['device']
            data[count]['Menu'] = row['menu']
            data[count]['format'] = str(row['format']).replace('\n', '') if row['format'] else ''
            data[count]['clicks'] = int(clicks if not pd.isna(clicks) else 0)
            data[count]['Impressions'] = int(impressions if not pd.isna(impressions) else 0)
            count+=1
    return count

def media_data_2018(folder_name, file, data, count):
    print('media_data_2018')
    filepath = os.path.join(folder_name, file)
    readexcelfile = pd.read_excel(filepath, engine='openpyxl', sheet_name='SNS_Online delivery by site')
    # readexcelfile.to_excel('readed data.xlsx')
    weekwise_data = {}
    prevPlatform = None
    prevDevice = None
    prevMenu = None
    prevFormat = None

    for i, j in readexcelfile.iterrows():
        # print(
        dt = datetime.datetime.strptime('2018/5/4', '%Y/%m/%d')
        count1 = 0
        platform1 = j['Platform'] if j['Platform'] != '' and not pd.isna(j['Platform']) else prevPlatform
        if platform1 == 'COOKPAD':
            platform1 = 'Cookpad'
        device = j['device'] if j['device'] != '' and not pd.isna(j['device'])  else prevDevice
        menu = j['Menu'] if j['Menu'] != '' and not pd.isna(j['Menu']) else prevMenu
        if device == 'Total':
            prevDevice = device
            continue
        print('platform', platform1)
        print('device', device)
        format = j['Ads Fromat'] if j['Ads Fromat'] != '' and not pd.isna(j['Ads Fromat']) else prevFormat
        imp_or_click = j['KPI']
        platform = platform1 + device + menu
        for k in range(4,23):
            dt1 = dt + datetime.timedelta(days=count1)
            d1 = dt1.isocalendar()
            week_num = d1[1]
            row = {'date': str(dt1), 'platform': platform1, 'device': device, 'menu': menu, 'format': format, imp_or_click: j[k], 'week_num': week_num}

            if not platform in weekwise_data:
                weekwise_data[platform] = {}

            if week_num in weekwise_data[platform]:
                weekwise_data[platform][week_num].append(row)
            else:
                weekwise_data[platform][week_num] = [row]

            count1 += 1
        prevPlatform = platform1
        prevDevice = device
        prevMenu = menu
        prevFormat = format

    print(weekwise_data)
    for i in weekwise_data.keys():
        for k in weekwise_data[i].keys():
            # print(k)
            clicks = 0
            impressions = 0
            for j in range(len(weekwise_data[i][k])):
                row = weekwise_data[i][k][j]
                print(row)
                if 'imp' in row and not pd.isna(row['imp']):
                    impressions += row['imp']
                elif 'clicks' in row and not pd.isna(row['clicks']):
                    clicks += row['clicks']
            data1, year, month = split_on_hyphen(row['date'])
            data.append({})
            data[count]['Year_Week'] = str(year) + '_' + str(row['week_num']).zfill(2)
            data[count]['Year_Month'] = str(year) + '_' + str(month).zfill(2)
            data[count]['Year'] = str(year)
            data[count]['Month'] = month
            data[count]['Week'] = row['week_num']
            data[count]['Platform'] = row['platform']
            data[count]['Device'] = row['device']
            data[count]['Menu'] = row['menu']
            data[count]['format'] = row['format']
            data[count]['clicks'] = int(clicks if not pd.isna(clicks) else 0)
            data[count]['Impressions'] = int(impressions if not pd.isna(impressions) else 0)
            count+=1
    return count

def media_data_phase_2_2018(folder_name, file, data, count):
    filepath = os.path.join(folder_name, file)
    readexcelfile = pd.read_excel(filepath, engine='openpyxl', sheet_name='SNS_Online delivery by site')
    # readexcelfile.to_excel('readed data.xlsx')
    weekwise_data = {}
    prevPlatform = None
    prevDevice = None
    prevMenu = None
    prevFormat = None

    for i, j in readexcelfile.iterrows():
        # print(
        dt = datetime.datetime.strptime('2018/6/23', '%Y/%m/%d')
        count1 = 0
        platform1 = j['Platform'] if j['Platform'] != '' and not pd.isna(j['Platform']) else prevPlatform
        if platform1 == 'COOKPAD':
            platform1 = 'Cookpad'
        device = j['device'] if j['device'] != '' and not pd.isna(j['device'])  else prevDevice
        if device == 'Total':
            prevDevice = device
            continue
        menu = j['Menu'] if j['Menu'] != '' and not pd.isna(j['Menu']) else prevMenu

        format = j['Ads Fromat'] if j['Ads Fromat'] != '' and not pd.isna(j['Ads Fromat']) else prevFormat
        imp_or_click = j['KPI']
        platform = platform1 + device + menu
        keys = j.keys()
        for k1 in range(5,75):
            k = keys[k1]
            dt1 = dt + datetime.timedelta(days=count1)
            d1 = dt1.isocalendar()
            week_num = d1[1]
            row = {'date': str(dt1), 'platform': platform1, 'device': device, 'menu': menu, 'format': format, imp_or_click: j[k], 'week_num': week_num}

            if not platform in weekwise_data:
                weekwise_data[platform] = {}

            if week_num in weekwise_data[platform]:
                weekwise_data[platform][week_num].append(row)
            else:
                weekwise_data[platform][week_num] = [row]

            count1 += 1
        prevPlatform = platform1
        prevDevice = device
        prevMenu = menu
        prevFormat = format

    # print(weekwise_data)
    for i in weekwise_data.keys():
        for k in weekwise_data[i].keys():
            # print(k)
            clicks = 0
            impressions = 0
            for j in range(len(weekwise_data[i][k])):
                row = weekwise_data[i][k][j]
                print(row)
                if 'imp' in row and not pd.isna(row['imp']):
                    impressions += row['imp']
                elif 'clicks' in row and not pd.isna(row['clicks']):
                    clicks += row['clicks']
            data1, year, month = split_on_hyphen(row['date'])
            data.append({})
            data[count]['Year_Week'] = str(year) + '_' + str(row['week_num']).zfill(2)
            data[count]['Year_Month'] = str(year) + '_' + str(month).zfill(2)
            data[count]['Year'] = str(year)
            data[count]['Month'] = month
            data[count]['Week'] = row['week_num']
            data[count]['Platform'] = row['platform']
            data[count]['Device'] = row['device']
            data[count]['Menu'] = row['menu']
            data[count]['format'] = row['format']
            data[count]['clicks'] = int(clicks if not pd.isna(clicks) else 0)
            data[count]['Impressions'] = int(impressions if not pd.isna(impressions) else 0)
            count+=1
    return count

def media_data_2019(folder_name, file, data, count):
    filepath = os.path.join(folder_name, file)
    readexcelfile = pd.read_excel(filepath, engine='openpyxl', sheet_name='Online delivery by site (2)')
    weekwise_data = {}
    prevPlatform = None
    prevDevice = None
    prevMenu = None
    prevFormat = None

    for i, j in readexcelfile.iterrows():
        dt = datetime.datetime.strptime('2019/5/29', '%Y/%m/%d')
        count1 = 0
        platform1 = j['Platform'] if j['Platform'] != '' and not pd.isna(j['Platform']) else prevPlatform
        if platform1 == 'COOKPAD':
            platform1 = 'Cookpad'
        menu = j['Menu'] if j['Menu'] != '' and not pd.isna(j['Menu']) else prevMenu
        device = j['Device'] if j['Device'] != '' and not pd.isna(j['Device']) else prevDevice
        if device == 'Total':
            prevDevice = device
            continue
        print('platform', platform1)
        imp_or_click = j['KPI']
        platform = platform1 + menu
        keys = j.keys()
        print(keys)
        for k1 in range(4,66):
            k = keys[k1]
            dt1 = dt + datetime.timedelta(days=count1)
            d1 = dt1.isocalendar()
            week_num = d1[1]
            row = {'date': str(dt1), 'platform': platform1, 'menu': menu, imp_or_click: j[k], 'week_num': week_num, 'k': k, 'device': device}
            print(row)
            if not platform in weekwise_data:
                weekwise_data[platform] = {}

            if week_num in weekwise_data[platform]:
                weekwise_data[platform][week_num].append(row)
            else:
                weekwise_data[platform][week_num] = [row]

            count1 += 1
        prevPlatform = platform1
        prevDevice = device
        prevMenu = menu
        # prevFormat = format

    print(weekwise_data)
    for i in weekwise_data.keys():
        for k in weekwise_data[i].keys():
            # print(k)
            clicks = 0
            impressions = 0
            for j in range(len(weekwise_data[i][k])):
                row = weekwise_data[i][k][j]
                print(row)
                if 'imp' in row and not pd.isna(row['imp']):
                    impressions += row['imp']
                elif 'clicks' in row and not pd.isna(row['clicks']):
                    clicks += row['clicks']
            data1, year, month = split_on_hyphen(row['date'])
            data.append({})
            data[count]['Year_Week'] = str(year) + '_' + str(row['week_num']).zfill(2)
            data[count]['Year_Month'] = str(year) + '_' + str(month).zfill(2)
            data[count]['Year'] = str(year)
            data[count]['Month'] = month
            data[count]['Week'] = row['week_num']
            data[count]['Platform'] = row['platform']
            data[count]['Device'] = row['device']
            data[count]['Menu'] = row['menu']
            data[count]['format'] = ''
            data[count]['clicks'] = int(clicks if not pd.isna(clicks) else 0)
            data[count]['Impressions'] = int(impressions if not pd.isna(impressions) else 0)
            count+=1
    return count

def media_data_2020(folder_name, file, data, count):
    filepath = os.path.join(folder_name, file)
    readexcelfile = pd.read_excel(filepath, engine='openpyxl', sheet_name='Online delivery by site')
    weekwise_data = {}
    prevPlatform = None
    prevDevice = None
    prevMenu = None
    prevFormat = None

    for i, j in readexcelfile.iterrows():
        dt = datetime.datetime.strptime('2020/5/7', '%Y/%m/%d')
        count1 = 0
        platform1 = j['Platform'] if j['Platform'] != '' and not pd.isna(j['Platform']) else prevPlatform
        if platform1 == 'COOKPAD':
            platform1 = 'Cookpad'
        device = j['Device'] if j['Device'] != '' and not pd.isna(j['Device'])  else prevDevice
        print('device', device)
        if device == 'Total':
            prevDevice = device
            continue
        print('platform', platform1)
        print('device', device)
        menu = j['Menu'] if j['Menu'] != '' and not pd.isna(j['Menu']) else prevMenu


        format = j['Format'] if j['Format'] != '' and not pd.isna(j['Format']) else prevFormat
        imp_or_click = j['KPI']
        platform = platform1 + device + menu
        keys = j.keys()
        # print(keys)
        for k1 in range(5,95):
            k = keys[k1]
            dt1 = dt + datetime.timedelta(days=count1)
            d1 = dt1.isocalendar()
            week_num = d1[1]
            row = {'date': str(dt1), 'platform': platform1, 'device': device, 'menu': menu, 'format': format, imp_or_click: j[k], 'week_num': week_num}

            if not platform in weekwise_data:
                weekwise_data[platform] = {}

            if week_num in weekwise_data[platform]:
                weekwise_data[platform][week_num].append(row)
            else:
                weekwise_data[platform][week_num] = [row]

            count1 += 1
        prevPlatform = platform1
        prevDevice = device
        prevMenu = menu
        prevFormat = format

    # print(weekwise_data)
    for i in weekwise_data.keys():
        for k in weekwise_data[i].keys():
            # print(k)
            clicks = 0
            impressions = 0
            for j in range(len(weekwise_data[i][k])):
                row = weekwise_data[i][k][j]
                # print(row)
                if 'imp' in row and not pd.isna(row['imp']):
                    impressions += row['imp']
                elif 'clicks' in row and not pd.isna(row['clicks']):
                    clicks += row['clicks']
            data1, year, month = split_on_hyphen(row['date'])
            data.append({})
            data[count]['Year_Week'] = str(year) + '_' + str(row['week_num']).zfill(2)
            data[count]['Year_Month'] = str(year) + '_' + str(month).zfill(2)
            data[count]['Year'] = str(year)
            data[count]['Month'] = month
            data[count]['Week'] = row['week_num']
            data[count]['Platform'] = row['platform']
            data[count]['Device'] = row['device']
            data[count]['Menu'] = row['menu']
            data[count]['format'] = row['format']
            data[count]['clicks'] = int(clicks if not pd.isna(clicks) else 0)
            data[count]['Impressions'] = int(impressions if not pd.isna(impressions) else 0)
            count+=1
    return count

if __name__ == '__main__':
    main()