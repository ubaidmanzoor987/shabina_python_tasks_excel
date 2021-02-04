from datetime import datetime

import pandas as pd
import os
import re
from collections import Counter
from dateutil import rrule

def split_on_dash(fileName):
    year = None
    try:
        year = fileName.split('-')[0]
    except:
        try:
            year = fileName.split('-')[0]
        except:
            year = ''
    return year

def my_mode(sample):
    c = Counter(sample)
    return [k for k, v in c.items() if v == c.most_common(1)[0][1]]

def main():
    folder_name = 'Global Covid'
    data = [{} for i in range(800)]
    file = 'Consumer confidence index.xlsx'
    filepath = os.path.join(folder_name, file)
    print(file)
    readexcelfile = pd.read_excel(filepath, engine='openpyxl', sheet_name='JPKRAUS')
    count = 0
    for i, j in readexcelfile.iterrows():
        year = split_on_dash(j[5])
        if year == '2020':
            if j[0] == 'AUS' or j[0] == 'KOR' or j[0] == 'JPN':
                if j[0] == 'AUS':
                    j[0] = 'Australia'
                elif j[0] == 'KOR':
                    j[0] = 'Korea'
                elif j[0] == 'JPN':
                    j[0] = 'JAPAN'
                data[count]['Month'] = int(j[5].split('-')[1])
                data[count]['Country'] = j[0]
                data[count]['Consumer Confidence'] = j[6]
                count +=1
    counter = 1
    readexcelfile = pd.read_excel(filepath, engine='openpyxl', sheet_name='VN,SG')
    for i, j in readexcelfile.iterrows():
        year = j[1]
        month = j[0]
        if year == 2020:
            if month == 'JAN' or month == 'APR' or month == 'JUL':
                data[count]['Country'] = 'Viet nam'
                data[count]['Month'] = counter
                data[count]['Consumer Confidence'] = j[2]
                count += 1
            else:
                data[count]['Country'] = 'Viet Nam'
                data[count]['Month'] = counter
                data[count]['Consumer Confidence'] = ''
                count += 1
            counter +=1
    counter = 1
    for i, j in readexcelfile.iterrows():
        year = j[1]
        month = j[0]
        if year == 2020:
            if month == 'JAN' or month == 'APR' or month == 'JUL':
                data[count]['Country'] = 'Singapore'
                data[count]['Month'] = counter
                data[count]['Consumer Confidence'] = j[3]
                count += 1
            else:
                data[count]['Country'] = 'Singapore'
                data[count]['Month'] = counter
                data[count]['Consumer Confidence'] = ''
                count += 1
            counter+=1
        start = datetime(year=2020, month=1, day=1)
        end = datetime(year=2020, month=12, day=31)
        data1 = []
        for dt in rrule.rrule(rrule.WEEKLY, dtstart=start, until=end):
            month = dt.month
            year = dt.year
            di = dt.isocalendar()
            week_num = di[1]
            for row in data:
                if 'Month' in row and row['Month'] == month:
                    row1 = row.copy()
                    row1['year_week'] = str(year) + '_' + str(week_num)
                    row1['year_month'] = str(year) + '_' + str(month)
                    row1['Year'] = year
                    row1['Week'] = week_num
                    data1.append(row1)


    file = 'Number of cases.csv'
    filepath = os.path.join(folder_name, file)
    print(file)
    readcsvfilejp = pd.read_csv(filepath)
    weekwise_data = {}
    for i, j in readcsvfilejp.iterrows():
        if j[2] == 'Australia' or j[2] == 'Republic of Korea' or j[2] == 'Singapore' or j[2] == 'Japan' or j[2] == 'Viet Nam':
            dt = datetime.fromisoformat(j[0])
            if not dt.year == 2020:
                continue
            d1 = dt.isocalendar()
            week_num = d1[1]
            if not j[2] in weekwise_data:
                weekwise_data[j[2]] = {}

            if week_num in weekwise_data[j[2]]:
                weekwise_data[j[2]][week_num].append(j)
            else:
                weekwise_data[j[2]][week_num] = [j]

    for i in weekwise_data.keys():
        for k in weekwise_data[i].keys():
            new_cases = 0
            Cumulative_cases = 0
            total_deaths = 0
            Cumulative_deaths = 0
            for j in range(len(weekwise_data[i][k])):
                row = weekwise_data[i][k][j]
                new_cases += row[4]
                Cumulative_cases += row[5]
                total_deaths += row[6]
                Cumulative_deaths += row[7]
            our_row = None
            for x in range(len(data1)):
                row = data1[x]
                if (row['Country'].lower() == i.lower() or (row['Country'] == 'Korea' and i == 'Republic of Korea')) and row['Week'] == k:
                    our_row = row
                    break
            our_row['New Cases'] = new_cases
            our_row['Cumulative_cases'] = Cumulative_cases
            our_row['New_deaths'] = total_deaths
            our_row['Cumulative_deaths'] = Cumulative_deaths

    file = 'Stringency index.xlsx'
    print(file)
    filepath = os.path.join(folder_name, file)
    readcsvfilejp = pd.read_excel(filepath, engine='openpyxl', sheet_name=None)
    sheets = readcsvfilejp.keys()
    for sheet in sheets:
        sheetObj = pd.read_excel(filepath, engine='openpyxl', sheet_name = sheet)
        if not sheet == 'Stringency_Index_Cleaned':
            stingeny(sheetObj, sheet, data1)

    df1 = pd.DataFrame(data1)
    df1.to_excel('Covid6.xlsx', index=False)


def stingeny(sheetName, column_name, data1):
    weekwise_data = {}
    for i, k in sheetName.iterrows():
        if k[1] == 'Australia' or k[1] == 'South Korea' or k[1] == 'Singapore' or k[1] == 'Japan' or k[1] == 'Vietnam':
            count = 2
            for j in sheetName.keys()[2:367]:
                dt = datetime.strptime(j, '%d%b%Y')
                if not dt.year == 2020:
                    continue
                d1 = dt.isocalendar()
                week_num = d1[1]
                if not k[1] in weekwise_data:
                    weekwise_data[k[1]] = {}

                if week_num in weekwise_data[k[1]]:
                    weekwise_data[k[1]][week_num].append(k[count])
                else:
                    weekwise_data[k[1]][week_num] = [k[count]]
                count += 1

    for i in weekwise_data.keys():
        for k in weekwise_data[i].keys():
            if column_name == 'stringency_index' or column_name == 'government_response_index' or column_name == 'containment_health_index' or column_name == 'economic_support_index':
                new_cases = 0
                for j in range(len(weekwise_data[i][k])):
                    row = weekwise_data[i][k][j]
                    print(row)
                    new_cases += row
                new_cases /= len(weekwise_data[i][k])
                our_row = None
                for x in range(len(data1)):
                    row = data1[x]
                    print(i, k)
                    if (row['Country'].lower() == i.lower() or (row['Country'] == 'Korea' and i == 'South Korea') or ((row['Country'] == 'Viet nam' or row['Country'] == 'Viet Nam') and i == 'Vietnam')) and \
                            row['Week'] == k:
                        our_row = row
                        break
                our_row[column_name] = new_cases
            else:
                new_cases = my_mode(str(weekwise_data[i][k]))[0]
                our_row = None
                for x in range(len(data1)):
                    row = data1[x]
                    print(i, k)
                    if (row['Country'].lower() == i.lower() or (row['Country'] == 'Korea' and i == 'South Korea') or (
                            (row['Country'] == 'Viet nam' or row['Country'] == 'Viet Nam') and i == 'Vietnam')) and \
                            row['Week'] == k:
                        our_row = row
                        break
                our_row[column_name] = new_cases


if __name__ == '__main__':
    main()