from datetime import datetime, timedelta

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

def my_mode(lst):
    return max(set(lst), key=lst.count)
    # c = Counter(sample)
    # return [k for k, v in c.items() if v == c.most_common(1)[0][1]]
def check_date_forat(date_str):
    val = None
    try:
        val = datetime.strptime(date_str, '%Y-%m-%d')
    except:
        try:
            val = datetime.strptime(date_str, '%m/%d/%Y')
        except:
            try:
                val = datetime.strptime(date_str, '%m/%d/%Y')
            except:
                val = datetime.strptime(date_str, '%d-%m-%Y')
    return val

def check_date_forat1(date_str):
    val = None
    try:
        val = datetime.strptime(date_str, '%Y-%m-%d')
    except:
        try:
            val = datetime.strptime(date_str, '%d/%m/%Y')
        except:
            try:
                val = datetime.strptime(date_str, '%m/%d/%Y')
            except:
                val = datetime.strptime(date_str, '%d-%m-%Y')
    return val

def main():
    folder_name = 'Dataset'
    data = [{} for i in range(800)]
    file = 'Consumer confidence index.xlsx'
    filepath = os.path.join(folder_name, file)
    print(file)
    readexcelfile = pd.read_excel(filepath, engine='openpyxl', sheet_name='JPKRAUS')
    count = 0
    for i, j in readexcelfile.iterrows():
        year = split_on_dash(j[5])
        if year == '2020' or year == '2021':
            if j[0] == 'AUS' or j[0] == 'KOR' or j[0] == 'JPN':
                if j[0] == 'AUS':
                    j[0] = 'Australia'
                elif j[0] == 'KOR':
                    j[0] = 'Korea'
                elif j[0] == 'JPN':
                    j[0] = 'JAPAN'
                data[count]['Month'] = int(j[5].split('-')[1])
                data[count]['Year'] = int(year)
                data[count]['Country'] = j[0]
                data[count]['Consumer Confidence'] = j[6]
                count +=1
    counter = 1
    readexcelfile = pd.read_excel(filepath, engine='openpyxl', sheet_name='VN,SG')
    prevValueCountry = None
    for i, j in readexcelfile.iterrows():
        year = j[1]
        month = j[0]

        if year == 2020 or year == 2021:
            if month == 'JAN' or month == 'APR' or month == 'JUL' or month == 'OCT':
                data[count]['Country'] = 'Viet nam'
                data[count]['Year'] = year
                data[count]['Month'] = counter
                data[count]['Consumer Confidence'] = j[2]
                prevValueCountry = j[2]
                count += 1
            else:
                data[count]['Country'] = 'Viet Nam'
                data[count]['Year'] = year
                data[count]['Month'] = counter
                data[count]['Consumer Confidence'] = prevValueCountry
                count += 1
            counter +=1
    counter = 1
    prevValueCountry = None
    for i, j in readexcelfile.iterrows():
        year = j[1]
        month = j[0]
        if year == 2020 or year == 2021:
            if month == 'JAN' or month == 'APR' or month == 'JUL' or month == 'OCT':
                data[count]['Country'] = 'Singapore'
                data[count]['Month'] = counter
                data[count]['Year'] = year
                data[count]['Consumer Confidence'] = j[3]
                prevValueCountry = j[3]
                count += 1
            else:
                data[count]['Country'] = 'Singapore'
                data[count]['Month'] = counter
                data[count]['Year'] = year
                data[count]['Consumer Confidence'] = prevValueCountry
                count += 1
            counter+=1
        start = datetime(year=2020, month=1, day=1)
        end = datetime(year=2021, month=2, day=28)
        data1 = []
        for dt in rrule.rrule(rrule.WEEKLY, dtstart=start, until=end):
            month = dt.month
            year = dt.year
            di = dt.isocalendar()
            start_day = dt - timedelta(days=dt.weekday())
            end_day = start_day + timedelta(days=6)
            week_num = di[1]
            for row in data:
                row1 = None
                if 'Month' in row and row['Month'] == month and 'Year' in row and row['Year'] == year:
                    row1 = row.copy()
                    row1['year_week'] = str(year) + '_' + str(week_num).zfill(2)
                    row1['year_month'] = str(year) + '_' + str(month).zfill(2)
                    row1['Year'] = year
                    row1['Week'] = int(week_num)
                    data1.append(row1)
                elif 'Month' in row and row['Month'] == month:
                    row1 = row.copy()
                    row1['year_week'] = str(year) + '_' + str(week_num).zfill(2)
                    row1['year_month'] = str(year) + '_' + str(month).zfill(2)
                    row1['Year'] = year
                    row1['Week'] = int(week_num)
                    row1['Consumer Confidence'] = ''
                    data1.append(row1)

    df1 = pd.DataFrame(data1)
    # df1.to_excel('Covid v3 inter.xlsx', index=False)

    file = 'Number of cases (till Feb).csv'
    filepath = os.path.join(folder_name, file)
    print(file)
    readcsvfilejp = pd.read_csv(filepath)
    weekwise_data = {}
    for i, j in readcsvfilejp.iterrows():
        if j[2] == 'Australia' or j[2] == 'Republic of Korea' or j[2] == 'Singapore' or j[2] == 'Japan' or j[2] == 'Viet Nam':
            # dt = datetime.fromisoformat(j[0])
            dt = check_date_forat(j[0])
            if dt.year == 2020 or dt.year == 2021:
                d1 = dt.isocalendar()
                start_day = dt - timedelta(days=dt.weekday())
                year = start_day.year if start_day.year != 2019 else 2020
                week_num = str(year) + '_' + str(d1[1]).zfill(2)
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
                Cumulative_cases = row[5]
                total_deaths += row[6]
                Cumulative_deaths = row[7]
            our_row = None
            for x in range(len(data1)):
                row = data1[x]
                if (row['Country'].lower() == i.lower() or (row['Country'] == 'Korea' and i == 'Republic of Korea')) and row['year_week'] == k:
                    our_row = row
                    break
            print(row['year_week'])
            print(k)
            our_row['New Cases'] = new_cases
            our_row['Cumulative_cases'] = Cumulative_cases
            our_row['New_deaths'] = total_deaths
            our_row['Cumulative_deaths'] = Cumulative_deaths

    file = 'Mobility Data/2020_AU_Region_Mobility_Report.csv'
    filepath = os.path.join(folder_name, file)
    readcsvfilejp = pd.read_csv(filepath)
    weekwise_data = {}
    for i, j in readcsvfilejp.iterrows():
        if j[1] == 'Australia' or j[1] == 'South Korea' or j[1] == 'Singapore' or j[1] == 'Japan' or j[1] == 'Vietnam':
            dt = check_date_forat1(j[7])
            if j[1] == 'South Korea' or j[1] == 'Singapore' or j[1] == 'Vietnam':
                dt = check_date_forat(j[7])
            if (dt.year == 2020 or dt.year == 2021) and pd.isna(j[2]) and pd.isna(j[3]) and pd.isna(j[4]):
                d1 = dt.isocalendar()
                start_day = dt - timedelta(days=dt.weekday())
                year = start_day.year if start_day.year != 2019 else 2020
                week_num = str(year) + '_' + str(d1[1]).zfill(2)
                if not j[1] in weekwise_data:
                    weekwise_data[j[1]] = {}

                if week_num in weekwise_data[j[1]]:
                    weekwise_data[j[1]][week_num].append(j)
                else:
                    weekwise_data[j[1]][week_num] = [j]

    for i in weekwise_data.keys():
        for k in weekwise_data[i].keys():
            retail_and_recreation_percent_change_from_baseline = 0.0
            grocery_and_pharmacy_percent_change_from_baseline = 0.0
            parks_percent_change_from_baseline = 0.0
            transit_stations_percent_change_from_baseline = 0.0
            workplaces_percent_change_from_baseline = 0.0
            residential_percent_change_from_baseline = 0.0

            for j in range(len(weekwise_data[i][k])):
                row = weekwise_data[i][k][j]
                retail_and_recreation_percent_change_from_baseline += row[8]
                grocery_and_pharmacy_percent_change_from_baseline += row[9]
                parks_percent_change_from_baseline += row[10]
                transit_stations_percent_change_from_baseline += row[11]
                workplaces_percent_change_from_baseline += row[12]
                residential_percent_change_from_baseline += row[13]

            size = len(weekwise_data[i][k])

            retail_and_recreation_percent_change_from_baseline /= size
            grocery_and_pharmacy_percent_change_from_baseline /= size
            parks_percent_change_from_baseline /= size
            transit_stations_percent_change_from_baseline /= size
            workplaces_percent_change_from_baseline /= size
            residential_percent_change_from_baseline /= size

            our_row = None
            for x in range(len(data1)):
                row = data1[x]
                if (row['Country'].lower() == i.lower() or (row['Country'] == 'Korea' and i == 'South Korea') or ((row['Country'] == 'Viet nam' or row['Country'] == 'Viet Nam') and i == 'Vietnam')) and row['year_week'] == k:
                    our_row = row
                    break
            # print(our_row)
            # print(i)
            our_row['retail_and_recreation_percent_change_from_baseline'] = retail_and_recreation_percent_change_from_baseline
            our_row['grocery_and_pharmacy_percent_change_from_baseline'] = grocery_and_pharmacy_percent_change_from_baseline
            our_row['parks_percent_change_from_baseline'] = parks_percent_change_from_baseline
            our_row['transit_stations_percent_change_from_baseline'] = transit_stations_percent_change_from_baseline
            our_row['workplaces_percent_change_from_baseline'] = workplaces_percent_change_from_baseline
            our_row['residential_percent_change_from_baseline'] = residential_percent_change_from_baseline

    file = 'Stringency index (till Feb).xlsx'
    print(file)
    filepath = os.path.join(folder_name, file)
    readcsvfilejp = pd.read_excel(filepath, engine='openpyxl', sheet_name=None)
    sheets = readcsvfilejp.keys()
    for sheet in sheets:
        #print(sheet)
        sheetObj = pd.read_excel(filepath, engine='openpyxl', sheet_name= sheet)
        # print(sheetObj.keys())
        stingeny(sheetObj, sheet, data1)

    data2 = []
    counter2 = 0
    print(data1)
    for row in data1:
        print(row)
        if row['Country'].lower() == 'viet nam' or row['Country'].lower() == 'vietnam':
            row['Country'] = 'Vietnam'
        if row['Country'].lower() == 'japan':
            row['Country'] = 'Japan'
        new_row = {}
        new_row['Country'] = row['Country']
        new_row['Year_week'] = row['year_week']
        new_row['Year_month'] = row['year_month']
        new_row['Year'] = row['Year']
        new_row['Month'] = row['Month']
        new_row['Week'] = row['Week']
        new_row['Consumer Confidence'] = row['Consumer Confidence'] if 'Consumer Confidence' in row else ''
        new_row['New Cases'] = row['New Cases'] if 'New Cases' in row else ''
        new_row['Cumulative_cases'] = row['Cumulative_cases'] if 'Cumulative_cases' in row else ''
        new_row['New_deaths'] = row['New_deaths'] if 'New_deaths' in row else ''
        new_row['Cumulative_deaths'] = row['Cumulative_deaths'] if 'Cumulative_deaths' in row else ''
        new_row['stringency_index'] = row['stringency_index'] if 'stringency_index' in row else ''
        new_row['government_response_index'] = row['government_response_index'] if 'government_response_index' in row else ''
        new_row['containment_health_index'] = row['containment_health_index'] if 'containment_health_index' in row else ''
        new_row['economic_support_index'] = row['economic_support_index'] if 'economic_support_index' in row else ''
        new_row['c1_school_closing'] = row['c1_school_closing'] if 'c1_school_closing' in row else ''
        new_row['c2_workplace_closing'] = row['c2_workplace_closing'] if 'c2_workplace_closing' in row else ''
        new_row['c3_cancel_public_events'] = row['c3_cancel_public_events'] if 'c3_cancel_public_events' in row else ''
        new_row['c4_restrictions_on_gatherings'] = row['c4_restrictions_on_gatherings'] if 'c4_restrictions_on_gatherings' in row else ''
        new_row['c5_close_public_transport'] = row['c5_close_public_transport'] if 'c5_close_public_transport' in row else ''
        new_row['c6_stay_at_home_requirements'] = row['c6_stay_at_home_requirements'] if 'c6_stay_at_home_requirements' in row else ''
        new_row['c7_movementrestrictions'] = row['c7_movementrestrictions'] if 'c7_movementrestrictions' in row else ''
        new_row['c8_internationaltravel'] = row['c8_internationaltravel'] if 'c8_internationaltravel' in row else ''
        new_row['e1_income_support'] = row['e1_income_support'] if 'e1_income_support' in row else ''
        new_row['e2_debtrelief'] = row['e2_debtrelief'] if 'e2_debtrelief' in row else ''
        new_row['h1_public_information_campaigns'] = row['h1_public_information_campaigns'] if 'h1_public_information_campaigns' in row else ''
        new_row['h2_testing_policy'] = row['h2_testing_policy'] if 'h2_testing_policy' in row else ''
        new_row['h3_contact_tracing'] = row['h3_contact_tracing'] if 'h3_contact_tracing' in row else ''
        new_row['h6_facial_coverings'] = row['h6_facial_coverings'] if 'h6_facial_coverings' in row else ''
        new_row['h7_vaccination_policy'] = row['h7_vaccination_policy'] if 'h7_vaccination_policy' in row else ''
        new_row['confirmed_cases'] = row['confirmed_cases'] if 'confirmed_cases' in row else ''
        new_row['confirmed_deaths'] = row['confirmed_deaths'] if 'confirmed_deaths' in row else ''
        new_row['retail_and_recreation_percent_change_from_baseline'] = row['retail_and_recreation_percent_change_from_baseline'] if 'retail_and_recreation_percent_change_from_baseline' in row else ''
        new_row['grocery_and_pharmacy_percent_change_from_baseline'] = row['grocery_and_pharmacy_percent_change_from_baseline'] if 'grocery_and_pharmacy_percent_change_from_baseline' in row else ''
        new_row['parks_percent_change_from_baseline'] = row['parks_percent_change_from_baseline'] if 'parks_percent_change_from_baseline' in row else ''
        new_row['transit_stations_percent_change_from_baseline'] = row['transit_stations_percent_change_from_baseline'] if 'transit_stations_percent_change_from_baseline' in row else ''
        new_row['workplaces_percent_change_from_baseline'] = row['workplaces_percent_change_from_baseline'] if 'workplaces_percent_change_from_baseline' in row else ''
        new_row['residential_percent_change_from_baseline'] = row['residential_percent_change_from_baseline'] if 'residential_percent_change_from_baseline' in row else ''
        data2.append(new_row)
    df1 = pd.DataFrame(data2)
    df1.to_excel('Cleaned Dataset/Covid v4 .xlsx', index=False)

def stingeny(sheetName, column_name, data1):
    weekwise_data = {}
    #print(sheetName.iterrows())
    for i, k in sheetName.iterrows():
        print(k[1])
        if k[1] == 'Australia' or k[1] == 'South Korea' or k[1] == 'Singapore' or k[1] == 'Japan' or k[1] == 'Vietnam':
            count = 2
            for j in sheetName.keys()[2:403]:
                dt = datetime.strptime(j, '%d%b%Y')
                if not dt.year == 2020 and not dt.year == 2021:
                    continue

                d1 = dt.isocalendar()
                start_day = dt - timedelta(days=dt.weekday())
                year = start_day.year if start_day.year != 2019 else 2020
                week_num = str(year) + '_' + str(d1[1]).zfill(2)
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
                            row['year_week'] == k:
                        our_row = row
                        break
                our_row[column_name] = new_cases
            else:
                new_mode = my_mode(weekwise_data[i][k])
                our_row = None
                for x in range(len(data1)):
                    row = data1[x]
                    print(i, k)
                    if (row['Country'].lower() == i.lower() or (row['Country'] == 'Korea' and i == 'South Korea') or (
                            (row['Country'] == 'Viet nam' or row['Country'] == 'Viet Nam') and i == 'Vietnam')) and \
                            row['year_week'] == k:
                        our_row = row
                        break
                our_row[column_name] = new_mode


if __name__ == '__main__':
    main()