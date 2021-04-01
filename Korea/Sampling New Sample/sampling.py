import pandas as pd
import os
import re
import datetime

from dateutil import rrule


def split_on_hypyhen_and_space(date):
    start = None
    end = None
    try:
        start = date.split(' - ')[0]
        end = date.split(' - ')[1]
    except:
        date = str(date).split(' ')[0]
        try:
            d = datetime.datetime.strptime(date, '%Y-%m-%d')
            start = d.strftime("%d %b")
            end = ''
        except:
            start = ''
            end = ''
    return start, end


def main():
    data = {}
    folder_name = 'Dataset'
    file = os.path.join(folder_name, 'sampling New_Sample Extract.xlsx')
    print(file)
    data = {}
    data1 = {}
    activity_dict = {
        "Sungold Sampling": (2,8),
        "Roadshow2": (9,15),
        "Big Roadshow Sampling": (16,22),
        "Both Sampling": (23,29),
        "Non NZ Sampling": (30, 36),
        "Non-sampling Sampling": (37, 43),
        "Roadshow Sungold": (48, 54),
        "Roadshow Both": (55, 61),
        "Roadshow Non NZ": (62, 68),
        "Roadshow Nonsampling": (69, 75),
        "On Pack": (44, 44),
        "Point of sale materials": (45, 45),
        "Additional display": (46, 46),
        # "Green Promotion": (47, 47)
    }
    metric = [
        "Number of Store","Number of Session",
        "Number of Promoter", "Number of Sampled", "Fruit Cost", "Promoter Cost",
        "Total Cost"
    ]
    readexcelfile = pd.read_excel(file, engine='openpyxl', sheet_name='sampling New')
    for i, j in readexcelfile.iterrows():
        retailer_name = j[1]
        dt = datetime.datetime.strptime(str(j[0]), '%Y-%m-%d %H:%M:%S')
        year = dt.year
        month = dt.month
        d1 = dt.isocalendar()
        week_num = d1[1]
        year_week = str(year) + '_' + str(week_num).zfill(2)
        for key, value in activity_dict.items():
            if key == "On Pack" or key == "Point of sale materials" or key == "Additional display":
                year_week_retailer = str(year) + '_' + str(week_num).zfill(2) + retailer_name + key
                if not year_week_retailer in data1:
                    data1[year_week_retailer] = {}
                data1[year_week_retailer]['Year_Month'] = str(year) + '_' + str(month).zfill(2)
                data1[year_week_retailer]['Year_Week'] = year_week
                data1[year_week_retailer]['Year'] = str(year)
                data1[year_week_retailer]['Month'] = str(month)
                data1[year_week_retailer]['Week'] = int(week_num)
                data1[year_week_retailer]['Date'] = dt.date()
                data1[year_week_retailer]['Activity'] = key
                data1[year_week_retailer]['Retailer'] = retailer_name
                data1[year_week_retailer]['Count'] = j[value[0]]
                continue
            year_week_retailer = str(year) + '_' + str(week_num).zfill(2) + retailer_name + key
            if not year_week_retailer in data:
                data[year_week_retailer] = {}
            data[year_week_retailer]['Year_Month'] = str(year) + '_' + str(month).zfill(2)
            data[year_week_retailer]['Year_Week'] = year_week
            data[year_week_retailer]['Year'] = str(year)
            data[year_week_retailer]['Month'] = str(month)
            data[year_week_retailer]['Week'] = int(week_num)
            data[year_week_retailer]['Date'] = dt.date()
            data[year_week_retailer]['Activity'] = key
            data[year_week_retailer]['Retailer'] = retailer_name
            count = 0
            for k in range(value[0], value[1]+1):
                data[year_week_retailer][metric[count]] = j[k]
                count+=1

    data2 = []
    for a in data.keys():
        data2.append(data[a])


    df1 = pd.DataFrame(data2)
    df1.to_excel('Cleaned Dataset/sampling New_Sample Extract Result1 v 1.xlsx', index=False)

    data2 = []
    for a in data1.keys():
        data2.append(data1[a])

    df1 = pd.DataFrame(data2)
    df1.to_excel('Cleaned Dataset/sampling New_Sample Extract Result2 v 1.xlsx', index=False)


if __name__ == '__main__':
    main()