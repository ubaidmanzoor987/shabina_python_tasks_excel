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
            try:
                dt1 = start.strftime("%m/%d/%Y")
                dt = datetime.datetime.strptime(dt1, "%d/%m/%Y")
            except:
                dt1 = start.strftime("%d/%m/%Y")
                dt = datetime.datetime.strptime(dt1, "%d/%m/%Y")

    return dt

folder_name = 'Dataset'
file = '2. Sampling data 2020 _ 2019_Sample Extract_v02.xlsx'
filepath = os.path.join(folder_name, file)
readexcelfile12 = pd.read_excel(filepath, engine='openpyxl', sheet_name='Store Location Reference')
state_region_dict = {}
for i, j in readexcelfile12.iterrows():
    country = j[0]
    region = j[1]
    chain = j[2]
    state = j[3]
    store_name = j[4]
    state_region_dict[store_name] = [country, region, state]

print(state_region_dict.keys())

def find_state_region(chain_name):
    country = None
    region = None
    state = None
    if chain_name in state_region_dict:
        country = state_region_dict[chain_name][0]
        region = state_region_dict[chain_name][1]
        state = state_region_dict[chain_name][2]

    return country, region, state

filepath = os.path.join(folder_name, file)
readexcelfile12 = pd.read_excel(filepath, engine='openpyxl', sheet_name='Type Reference')
referenece_dict = {}
for i, j in readexcelfile12.iterrows():
    store_name = j[1]
    store_type = j[2]
    referenece_dict[store_name] = store_type


def find_store_type(store_name):
    store_type = None
    if store_name in referenece_dict:
        store_type = referenece_dict[store_name]

    return store_type


def main():
    folder_name = 'Dataset'
    file = os.path.join(folder_name, '2. Sampling data 2020 _ 2019_Sample Extract_v02.xlsx')
    print(file)
    digital_data(file)

def digital_data(file):
    data = {}
    data2 = {}
    data4 = {}
    #Sheet 2019 Details
    prevDate = None
    readexcelfile = pd.read_excel(file, engine='openpyxl', sheet_name='2019 Details')
    dt_row = None
    for i, j in readexcelfile.iterrows():
        if (i >=0 and i<=8) or i==10 or i>198 or  i==88 or i==117:
            # print(j[2])
            continue
        # print(j[1], j[2])
        store = j['OMG VIETNAM - RETAIL MARKETING SERVICE AGENCY']
        country, region, state = find_state_region(store)
        store_type = find_store_type(store)
        for k in range(11, 276):
            if i == 9:
                date = j[k] if j[k] != '' and not pd.isna(j[k]) else prevDate
                dt = convert_to_date(j[k]) if j[k] != '' and not pd.isna(j[k]) else convert_to_date(prevDate)
                month = dt.month
                di = dt.isocalendar()
                year = dt.year
                week_num = di[1]
                year_month_week = str(year) + '_' + str(month) + str(week_num).zfill(2) + str(date)
                year_week = str(year) + '_' + str(week_num).zfill(2)
                if not year_month_week in data:
                    data[year_month_week] = {}
                data[year_month_week]['Year_Month'] = str(year) + '_' + str(month).zfill(2)
                data[year_month_week]['Year_Week'] = year_week
                data[year_month_week]['Year'] = year
                data[year_month_week]['Month'] = int(month)
                data[year_month_week]['Week'] = int(week_num)
            if i>9:

                total_trials = None
                total_buyer = None
                if k < 116:
                    if (k-11)%3 == 0:
                        continue
                    type = j[1]
                    # print(store, region, store_type)
                    year_month_week_1 = list(data.keys())[int((k - 11) / 3)]
                    year_month_week = year_month_week_1 + store
                    if not year_month_week in data2:
                        data2[year_month_week] = {}
                    data2[year_month_week]['Year_Month'] = data[year_month_week_1]['Year_Month']
                    data2[year_month_week]['Year_Week'] = data[year_month_week_1]['Year_Week']
                    data2[year_month_week]['Year'] = data[year_month_week_1]['Year']
                    data2[year_month_week]['Month'] = data[year_month_week_1]['Month']
                    data2[year_month_week]['Week'] = data[year_month_week_1]['Week']
                    data2[year_month_week]['Country'] = 'Vietnam'
                    data2[year_month_week]['Region'] = region
                    data2[year_month_week]['State'] = state
                    data2[year_month_week]['Mt Chain'] = type
                    data2[year_month_week]['Name Store'] = store
                    data2[year_month_week]['Type'] = store_type
                    if (k-11)%3 == 1:
                        # print(k)
                        total_buyer = j[k]
                        data2[year_month_week]['Total Buyers'] = float(total_buyer)
                    if (k - 11)%3 == 2:
                        total_trials = j[k]
                        data2[year_month_week]['Total Trial'] = float(total_trials)
                else:
                    type = j[1]
                    # print(store, region, store_type)
                    print(k)
                    year_month_week_1 = list(data.keys())[int((116 - 11) / 3) + int((k - 116) / 4)]
                    year_month_week = year_month_week_1 + store
                    if not year_month_week in data2:
                        data2[year_month_week] = {}
                    data2[year_month_week]['Year_Month'] = data[year_month_week_1]['Year_Month']
                    data2[year_month_week]['Year_Week'] = data[year_month_week_1]['Year_Week']
                    data2[year_month_week]['Year'] = data[year_month_week_1]['Year']
                    data2[year_month_week]['Month'] = data[year_month_week_1]['Month']
                    data2[year_month_week]['Week'] = data[year_month_week_1]['Week']
                    data2[year_month_week]['Country'] = 'Vietnam'
                    data2[year_month_week]['Region'] = region
                    data2[year_month_week]['State'] = state
                    data2[year_month_week]['Mt Chain'] = type
                    data2[year_month_week]['Name Store'] = store
                    data2[year_month_week]['Type'] = store_type
                    if (k-116)%4 == 0:
                        data2[year_month_week]['Sales G'] = float(j[k])
                    if (k-116)%4 == 1:
                        data2[year_month_week]['Sales S'] = float(j[k])
                    if (k - 116)%4 == 2:
                        total_buyers = j[k]
                        data2[year_month_week]['Total Buyers'] = float(total_buyers)
                    if (k - 116)%4 == 3:
                        total_trials = j[k]
                        data2[year_month_week]['Total Trial'] = float(total_trials)

            prevDate = date

    # Sheet 2020 Details
    readexcelfile = pd.read_excel(file, engine='openpyxl', sheet_name='2020 Details')
    for i, j in readexcelfile.iterrows():
        date = j[5]
        month = date.month
        di = date.isocalendar()
        year = date.year
        week_num = di[1]
        year_week = str(year) + '_' + str(week_num).zfill(2)
        year_month_week = year_week + j[9]
        # print(year_month_week )
        j13 = j[13] if not pd.isna(j[13]) else 0
        j14 = j[14] if not pd.isna(j[14]) else 0
        j16 = j[16] if not pd.isna(j[16]) else 0
        j17 = j[17] if not pd.isna(j[17]) else 0
        j19 = j[19] if not pd.isna(j[19]) else 0
        j20 = j[20] if not pd.isna(j[20]) else 0

        if not year_month_week in data4:
            data4[year_month_week] = {}
            data4[year_month_week]['Year_Month'] = str(year) + '_' + str(month).zfill(2)
            data4[year_month_week]['Year_Week'] = year_week
            data4[year_month_week]['Year'] = year
            data4[year_month_week]['Month'] = int(month)
            data4[year_month_week]['Week'] = int(week_num)
            data4[year_month_week]['Country'] = 'Vietnam'
            data4[year_month_week]['Region'] = j[2]
            data4[year_month_week]['State'] = j[8]
            data4[year_month_week]['Mt Chain'] = j[6]
            data4[year_month_week]['Name Store'] = j[9]
            data4[year_month_week]['Type'] = j[10]
            data4[year_month_week]['TRAFFIC'] = j[11] if not pd.isna(j[11]) else 0
            data4[year_month_week]['TRIALS G'] = j[13] if not pd.isna(j[13]) else 0
            data4[year_month_week]['TRIALS S'] = j[14] if not pd.isna(j[14]) else 0
            data4[year_month_week]['Total Trial'] = j13 + j14
            data4[year_month_week]['BUYER G'] = j[16] if not pd.isna(j[16]) else 0
            data4[year_month_week]['BUYER S'] = j[17] if not pd.isna(j[17]) else 0
            data4[year_month_week]['Total Buyers'] = j16 + j17
            data4[year_month_week]['Sales G'] = j[19] if not pd.isna(j[19]) else 0
            data4[year_month_week]['Sales S'] = j[20] if not pd.isna(j[20]) else 0
            data4[year_month_week]['Total Sales'] = j19 + j20
            data4[year_month_week]['Total Reach'] = j[25] if not pd.isna(j[25]) else 0
            data4[year_month_week]['% people tried/bought Zespri Kiwfruit before'] = j[26] if not pd.isna(j[26]) else 0
        else:
            data4[year_month_week]['TRAFFIC'] += j[11] if not pd.isna(j[11]) else 0
            data4[year_month_week]['TRIALS G'] += j[13] if not pd.isna(j[13]) else 0
            data4[year_month_week]['TRIALS S'] += j[14] if not pd.isna(j[14]) else 0
            data4[year_month_week]['Total Trial'] += j13 + j14
            data4[year_month_week]['BUYER G'] += j[16] if not pd.isna(j[16]) else 0
            data4[year_month_week]['BUYER S'] += j[17] if not pd.isna(j[17]) else 0
            data4[year_month_week]['Total Buyers'] += j16 + j17
            data4[year_month_week]['Sales G'] += j[19] if not pd.isna(j[19]) else 0
            data4[year_month_week]['Sales S'] += j[20] if not pd.isna(j[20]) else 0
            data4[year_month_week]['Total Sales'] += j19+ j20
            data4[year_month_week]['Total Reach'] += int(j[25]) if not pd.isna(j[25]) else 0
            data4[year_month_week]['% people tried/bought Zespri Kiwfruit before'] += j[26]
    
    data3 = []
    for a in data2.keys():
        # print(data2[a])
        unique_key = data2[a]['Year_Week'] + data2[a]['Name Store']
        if not unique_key in data4:
            data4[unique_key] = data2[a]
        else:
            if "Sales G" in data2[a]:
                val1 = data4[unique_key]['Sales G'] if not pd.isna(data4[unique_key]['Sales G']) else 0
                val2 = data2[a]["Sales G"] if not pd.isna(data2[a]["Sales G"]) else 0
                data4[unique_key]['Sales G'] = val1 + val2
            if "Sales S" in data2[a]:
                val1 = data4[unique_key]['Sales S'] if not pd.isna(data4[unique_key]['Sales S']) else 0
                val2 = data2[a]["Sales S"] if not pd.isna(data2[a]["Sales S"]) else 0
                data4[unique_key]['Sales S'] = val1 + val2

            val1 = data4[unique_key]['Total Buyers'] if not pd.isna(data4[unique_key]['Total Buyers']) else 0
            val2 = data2[a]["Total Buyers"] if not pd.isna(data2[a]["Total Buyers"]) else 0
            data4[unique_key]['Total Buyers'] = val1 + val2

            val1 = data4[unique_key]['Total Trial'] if not pd.isna(data4[unique_key]['Total Trial']) else 0
            val2 = data2[a]["Total Trial"] if not pd.isna(data2[a]["Total Trial"]) else 0
            data4[unique_key]['Total Trial'] = val1 + val2

    for key in data4.keys():
        data3.append(data4[key])

    df1 = pd.DataFrame(data3)
    df1.to_excel('Cleaned Dataset/Vietnam weekly Digital Result 2020_2019 v 3.xlsx', index=False)


if __name__ == '__main__':
    main()