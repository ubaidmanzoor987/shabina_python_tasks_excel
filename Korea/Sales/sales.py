import pandas as pd
import os
import datetime

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

def sale_data(file, data, column_name1, column_name_2, hasRetailer = True):
    readcsvfile = pd.read_csv(file)
    for i, j in readcsvfile.iterrows():
        retailer_name = j[1]
        dt = datetime.datetime.strptime(str(j[0]), '%d/%m/%Y')
        year = dt.year
        month = dt.month
        d1 = dt.isocalendar()
        week_num = d1[1]
        year_week = str(year) + '_' + str(week_num).zfill(2)
        year_week_retailer = str(year) + '_' + str(week_num).zfill(2)


        if not year_week_retailer in data:
            data[year_week_retailer] = {}
        if hasRetailer:
            if not retailer_name in data[year_week_retailer]:
                data[year_week_retailer][retailer_name] = {}
            data[year_week_retailer][retailer_name]['Year_Month'] = str(year) + '_' + str(month).zfill(2)
            data[year_week_retailer][retailer_name]['Year_Week'] = year_week
            data[year_week_retailer][retailer_name]['Year'] = str(year)
            data[year_week_retailer][retailer_name]['Month'] = str(month)
            data[year_week_retailer][retailer_name]['Week'] = int(week_num)
            data[year_week_retailer][retailer_name]['Date'] = dt.date()
            data[year_week_retailer][retailer_name]['Retailer'] = retailer_name
            data[year_week_retailer][retailer_name][column_name1] = float(j[2]) if str(j[2]).lower() != 'a' else j[2]
            data[year_week_retailer][retailer_name][column_name_2] = float(j[3]) if str(j[3]).lower() != 'a' else j[3]
        else:
            for key in data[year_week_retailer].keys():
                retailer_name = key
                data[year_week_retailer][retailer_name]['Year_Month'] = str(year) + '_' + str(month).zfill(2)
                data[year_week_retailer][retailer_name]['Year_Week'] = year_week
                data[year_week_retailer][retailer_name]['Year'] = str(year)
                data[year_week_retailer][retailer_name]['Month'] = str(month)
                data[year_week_retailer][retailer_name]['Week'] = int(week_num)
                data[year_week_retailer][retailer_name]['Date'] = dt.date()
                data[year_week_retailer][retailer_name]['Retailer'] = retailer_name
                data[year_week_retailer][retailer_name][column_name1] = float(j[1]) if j[1] != 'A' else j[1]
                data[year_week_retailer][retailer_name][column_name_2] = float(j[2]) if j[2] != 'A' else j[2]


def main():
    folder_name = 'Dataset'
    data = {}
    file = os.path.join(folder_name, 'Sales Gold.csv')
    print(file)
    sale_data(file, data, 'Sales Gold NZ BL', 'Sales Gold NZ ZR')
    file = os.path.join(folder_name, 'Sales Green .csv')
    print(file)
    sale_data(file, data, 'Sales Green NZ BL', 'Sales Green NZ ZR')
    file = os.path.join(folder_name, 'Price JanUp.csv')
    print(file)
    sale_data(file, data, 'Gold Price', 'Green Price', hasRetailer=False)

    data2 = []
    for a in data.keys():
        for b in data[a].keys():
            data2.append(data[a][b])


    df1 = pd.DataFrame(data2)
    df1.to_excel('Cleaned Dataset/Sales Result1 v 1.xlsx', index=False)

if __name__ == '__main__':
    main()