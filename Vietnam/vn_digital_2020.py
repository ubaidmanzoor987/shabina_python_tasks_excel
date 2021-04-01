import pandas as pd
import os
import re
import datetime

from dateutil import rrule


def convert_to_date(start):
    dt = None
    try:
        dt = datetime.datetime.strptime(start, "%d/%m/%Y")
    except:
        try:
            start1 = start.strip()
            dt = datetime.datetime.strptime(start1, "%d/%m/%Y")
        except:
            print(start)
            try:
                dt1 = start.strftime("%m/%d/%Y")
                dt = datetime.datetime.strptime(dt1, "%d/%m/%Y")
            except:
                dt1 = start.strftime("%d/%m/%Y")
                dt = datetime.datetime.strptime(dt1, "%d/%m/%Y")

    return dt

def main():
    folder_name = 'Dataset'
    file = os.path.join(folder_name, '2. Sampling data 2020 _ 2019.xlsx')
    print(file)
    digital_data(file, folder_name)


def digital_data(file, folder_name):
    data = {}
    #Sheet 2019 Details
    prevDate = None
    readexcelfile = pd.read_excel(file, engine='openpyxl', sheet_name='2019 Details')
    dt_row = None
    for i, j in readexcelfile.iterrows():
        if (i >=0 and i<=8) or (i>=11 and i<=200) or i == 10 or i>204:
            continue
        # print(j[2])
        for k in range(11, 277):
            if i == 9:
                date = j[k] if j[k] != '' and not pd.isna(j[k]) else prevDate
                dt = convert_to_date(j[k]) if j[k] != '' and not pd.isna(j[k]) else convert_to_date(prevDate)
                month = dt.month
                di = dt.isocalendar()
                year = dt.year
                week_num = di[1]
                year_month_week = str(year) + '_' + str(month) + str(week_num).zfill(2)
                year_week = str(year) + '_' + str(week_num).zfill(2)
                if not year_month_week in data:
                    data[year_month_week] = {}
                data[year_month_week]['Year_Month'] = str(year) + '_' + str(month).zfill(2)
                data[year_month_week]['Year_Week'] = year_week
                data[year_month_week]['Year'] = year
                data[year_month_week]['Month'] = int(month)
                data[year_month_week]['Week'] = int(week_num)
                dt_row = j

            if i == 201:
                sampling_day = j[k]
                date = dt_row[k] if dt_row[k] != '' and not pd.isna(dt_row[k]) else prevDate
                dt = convert_to_date(dt_row[k]) if dt_row[k] != '' and not pd.isna(dt_row[k]) else convert_to_date(prevDate)
                year = dt.year
                di = dt.isocalendar()
                week_num = di[1]
                month = dt.month
                year_month_week = str(year) + '_' + str(month) + str(week_num).zfill(2)
                if 'Sampling Day' in data[year_month_week]:
                    data[year_month_week]['Sampling Day'] += sampling_day if not pd.isna(sampling_day) else 0
                else:
                    data[year_month_week]['Sampling Day'] = sampling_day if not pd.isna(sampling_day) else 0
            if i == 202:
                trial = j[k]
                date = dt_row[k] if dt_row[k] != '' and not pd.isna(dt_row[k]) else prevDate
                dt = convert_to_date(dt_row[k]) if dt_row[k] != '' and not pd.isna(dt_row[k]) else convert_to_date(prevDate)
                year = dt.year
                di = dt.isocalendar()
                week_num = di[1]
                month = dt.month
                year_month_week = str(year) + '_' + str(month) + str(week_num).zfill(2)
                if 'Sample delivered' in data[year_month_week]:
                    data[year_month_week]['Sample delivered'] += trial if not pd.isna(trial) else 0
                else:
                    data[year_month_week]['Sample delivered'] = trial if not pd.isna(trial) else 0
            if i == 203:
                sales = j[k]
                date = dt_row[k] if dt_row[k] != '' and not pd.isna(dt_row[k]) else prevDate
                dt = convert_to_date(dt_row[k]) if dt_row[k] != '' and not pd.isna(dt_row[k]) else convert_to_date(
                    prevDate)
                year = dt.year
                di = dt.isocalendar()
                week_num = di[1]
                month = dt.month
                year_month_week = str(year) + '_' + str(month) + str(week_num).zfill(2)
                if 'Sales (kg)' in data[year_month_week]:
                    data[year_month_week]['Sales (kg)'] += sales if not pd.isna(sales) else 0
                else:
                    data[year_month_week]['Sales (kg)'] = sales if not pd.isna(sales) else 0
            if i == 204:
                buyer = j[k]
                date = dt_row[k] if dt_row[k] != '' and not pd.isna(dt_row[k]) else prevDate
                dt = convert_to_date(dt_row[k]) if dt_row[k] != '' and not pd.isna(dt_row[k]) else convert_to_date(prevDate)
                year = dt.year
                di = dt.isocalendar()
                week_num = di[1]
                month = dt.month
                year_month_week = str(year) + '_' + str(month) + str(week_num).zfill(2)
                if 'Shoppers' in data[year_month_week]:
                    data[year_month_week]['Shoppers'] += buyer if not pd.isna(buyer) else 0
                else:
                    data[year_month_week]['Shoppers'] = buyer if not pd.isna(buyer) else 0

            prevDate = date

    # Sheet 2020 Details
    readexcelfile = pd.read_excel(file, engine='openpyxl', sheet_name='2020 Details')
    for i, j in readexcelfile.iterrows():
        dt = None
        if i == 0 or i==2 or i ==8 or i==13 or i==18 or i==24 or i==29 or i==35 or i==41 or i>45:
            print(j[0])
            continue

        if i == 1:
            dt = convert_to_date('04/01/2021')
        elif (i >= 3 and i<=7) or  (i >= 9 and i<=12) or (i >= 14 and i<=17) or (i >= 19 and i<=23) or (i>=25 and i<=28) or (i>=30 and i<=34) or (i>=36 or i<=40) or (i>=42 and i<=45):
            dt = convert_to_date('04/05/2020')
            dt = dt + datetime.timedelta(days=(int(j[0][1:])-1)*7)

        dt1 =dt
        month = dt1.month
        di = dt1.isocalendar()
        year = dt1.year
        week_num = di[1]
        year_month_week = str(year) + '_' + str(month) + str(week_num).zfill(2) + str(j[0])
        year_week = str(year) + '_' + str(week_num).zfill(2)

        if not year_month_week in data:
            data[year_month_week] = {
                'Sampling Day': 0,
                'Sample delivered': 0,
                'Shoppers': 0,
                'Sales (kg)': 0,
                'Reach': 0,
                'Sum of TRIALS G': 0,
                'Sum of TRIALS S': 0,
                'Sum of SALE G': 0,
                'Sum of SALE S': 0,
                'Sum of BUYER G': 0,
                'Sum of BUYER S': 0,

            }
        j3 = j[3] if not pd.isna(j[3]) else 0
        j4 = j[4] if not pd.isna(j[4]) else 0
        j5 = j[5] if not pd.isna(j[5]) else 0
        j6 = j[6] if not pd.isna(j[6]) else 0
        j7 = j[7] if not pd.isna(j[7]) else 0
        j8 = j[8] if not pd.isna(j[8]) else 0

        data[year_month_week]['Year_Month'] = str(year) + '_' + str(month).zfill(2)
        data[year_month_week]['Year_Week'] = year_week
        data[year_month_week]['Year'] = year
        data[year_month_week]['Month'] = int(month)
        data[year_month_week]['Week'] = int(week_num)
        data[year_month_week]['Sampling Day'] += j[1]
        data[year_month_week]['Sample delivered'] += j3 + j4
        data[year_month_week]['Shoppers'] += j7 + j8
        data[year_month_week]['Sales (kg)'] += j5 + j6
        data[year_month_week]['Reach'] += j[2]
        data[year_month_week]['Sum of TRIALS G'] += j3
        data[year_month_week]['Sum of TRIALS S'] += j4
        data[year_month_week]['Sum of SALE G'] += j5
        data[year_month_week]['Sum of SALE S'] += j6
        data[year_month_week]['Sum of BUYER G'] += j7
        data[year_month_week]['Sum of BUYER S'] += j8

    data2 = []
    for a in data.keys():
        print(data[a])
        data2.append(data[a])

    df1 = pd.DataFrame(data2)
    df1.to_excel('Cleaned Dataset/Vietnam weekly Digital Result 2020_2019 v 1.xlsx', index=False)


if __name__ == '__main__':
    main()