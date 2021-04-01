import pandas as pd
import os
import re
import datetime

from dateutil import rrule


def split_on_hypyhen_and_space(date):
    start = None
    end =None
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

def split_on_hypyhen_and_space_and_new_line(date):
    start = None
    end = None
    topic = None
    try:
        date1 = date.split('\n')[0]
        topic = date.split('\n')[1]
        start = date1.split(' - ')[0]
        end = date1.split(' - ')[1]
    except:
        try:
            start = date1.split('- ')[0]
            end = date1.split('- ')[1]
        except:
            topic = ''
            start = ''
            end = ''

    return start, end, topic

def main():
    folder_name = 'Dataset'
    file = os.path.join(folder_name, 'VN digital weekly_2018.xlsx')
    print(file)
    digital_data(file, folder_name)

def digital_data(file, folder_name):
    data = {}
    count = 1
    # Sheet 2018
    prevBOOSTEDBY = None
    readexcelfile = pd.read_excel(file, engine='openpyxl', sheet_name='VN', skiprows=18)
    for i, j in readexcelfile.iterrows():
        if i >= 40:
            print(j)
            continue
        if not pd.isna(j[2]):
            start, end = split_on_hypyhen_and_space(j[2])
            dt = None
            dt_end = None
            if start.find('Sept'):
                start = start.replace('Sept ', 'Sep')
                start = start.replace('Sept', 'Sep')
            dt = datetime.datetime.strptime(start + ' 2018', '%d %b %Y')
            if end:
                end = end.replace('Sept ', 'Sep')
                end = end.replace('Sept', 'Sep')
                dt_end = datetime.datetime.strptime(end + ' 2018', '%d %b %Y')
            else:
                dt_end = dt

            noOfWeeks = len(list(rrule.rrule(rrule.WEEKLY , dtstart=dt, until=dt_end)))

            for dt1 in rrule.rrule(rrule.WEEKLY , dtstart=dt, until=dt_end):
                topics_boosted = j[1] if j[1] != '' and not pd.isna(j[1]) else prevBOOSTEDBY
                month = dt1.month
                di = dt1.isocalendar()
                year = dt1.year
                week_num = di[1]
                year_month_week = '2018' + '_' + str(month) + str(week_num).zfill(2) + str(j[2])
                year_week = str(year) + '_' + str(week_num).zfill(2)

                if not year_month_week in data:
                    data[year_month_week] = {}
                data[year_month_week]['Year_Month'] = str(year) + '_' + str(month).zfill(2)
                data[year_month_week]['Year_Week'] = year_week
                data[year_month_week]['Year'] = year
                data[year_month_week]['Month'] = int(month)
                data[year_month_week]['Week'] = int(week_num)
                data[year_month_week]['Platform'] = 'Facebook'
                data[year_month_week]['Menu'] = 'FB'
                data[year_month_week]['Topics'] = topics_boosted
                data[year_month_week]['Date'] = str(dt1).split(' ')[0]
                data[year_month_week]['Engagement'] = j[4]/noOfWeeks
                data[year_month_week]['Engagement Rate'] = j[4] / j[6] * 100
                data[year_month_week]['Reach'] = j[5]/noOfWeeks
                data[year_month_week]['Impression'] = j[6]/noOfWeeks
                data[year_month_week]['Video Views'] = j[8]/noOfWeeks if j[8] != '-' else j[8]
                data[year_month_week]['View Through Rate'] = j[9]/noOfWeeks if j[9] != '-' else j[9]
                data[year_month_week]['30s View'] = j[10]/noOfWeeks if j[10] != '-' else j[10]
                data[year_month_week]['30s View Through Rate'] = j[11]*100/noOfWeeks if not j[11] == '-' else '-'
                data[year_month_week]['Video 25%'] = j[12]*100 if not j[12] == '-' else '-'
                data[year_month_week]['Video 50%'] = j[13]*100 if not j[13] == '-' else '-'
                data[year_month_week]['Video 75%'] = j[14]*100 if not j[14] == '-' else '-'
                data[year_month_week]['Video 100%'] = j[15]*100 if not j[15] == '-' else '-'
                data[year_month_week]['Page Like'] = j[16]/noOfWeeks if j[16] != '-' else j[16]
                data[year_month_week]['Photo View'] = j[17]/noOfWeeks if j[17] != '-' else j[17]
                data[year_month_week]['Post Comments'] = j[18]/noOfWeeks if j[18] != '-' else j[18]
                data[year_month_week]['Post Reactions'] = j[19]/noOfWeeks if j[19] != '-' else j[19]
                data[year_month_week]['Post Share'] = j[20]/noOfWeeks if j[20] != '-' else j[20]
                data[year_month_week]['All Clicks'] = "" #j[21]/noOfWeeks if j[21] != '-' else j[21]
                data[year_month_week]['Actual Spent'] = j[23]/noOfWeeks if j[23] != '-' else j[23]
                prevBOOSTEDBY = topics_boosted

    # Sheet 2019 Digital
    year1 = None
    year2 = None
    file = os.path.join(folder_name, 'VN digital weekly_2019.xlsx')
    print(file)
    readexcelfile = pd.read_excel(file, engine='openpyxl', sheet_name='VN', skiprows= 23 )
    for i, j in readexcelfile.iterrows():
        if (i > 8 and i <=12) or (i>=19 and i<=22 ) or i>=24:
            # print(j)
            continue

        start, end, topic = split_on_hypyhen_and_space_and_new_line(j[1])
        if start.find('Sept'):
            start = start.replace('Sept ', 'Sep')
            start = start.replace('Sept', 'Sep')
        if end.find('Sept'):
            end = end.replace('Sept ', 'Sep')
            end = end.replace('Sept', 'Sep')
        start = start.strip()
        end = end.strip()
        if (start.find('Jan') != -1) and (end.find('Jan') != -1) or (start.find('Feb') != -1) and (end.find('Feb') != -1):
            year1 = ' 2020'
            year2 = ' 2020'
        elif (start.find('Jan') == -1) and (end.find('Jan') != -1):
            year1 = ' 2019'
            year2 = ' 2020'
        else:
            year1 = ' 2019'
            year2 = ' 2019'

        dt = None
        dt_end = None

        dt = datetime.datetime.strptime(start + year1, '%d %b %Y')
        if end:
            dt_end = datetime.datetime.strptime(end + year2, '%d %b %Y')
        else:
            dt_end = dt
        # print(dt, dt_end)
        noOfWeeks = len(list(rrule.rrule(rrule.WEEKLY, dtstart=dt, until=dt_end)))
        for dt1 in rrule.rrule(rrule.WEEKLY, dtstart=dt, until=dt_end):
            month = dt1.month
            di = dt1.isocalendar()
            week_num = di[1]
            year = dt1.year
            if topic == None:
                topic = ''
            year_month_week = str(year) + '_' + str(month) + str(week_num).zfill(2) + topic
            # print(year_month_week)
            year_week = str(year) + '_' + str(week_num).zfill(2)
            if not year_month_week in data:
                data[year_month_week] = {}
            data[year_month_week]['Year_Month'] = str(year) + '_' + str(month).zfill(2)
            data[year_month_week]['Year_Week'] = year_week
            data[year_month_week]['Year'] = year
            data[year_month_week]['Month'] = int(month)
            data[year_month_week]['Week'] = int(week_num)
            data[year_month_week]['Platform'] = 'Facebook'
            data[year_month_week]['Menu'] = 'FB'
            data[year_month_week]['Topics'] = topic
            data[year_month_week]['Date'] = str(dt1).split(' ')[0]
            data[year_month_week]['Engagement'] = j[5] / noOfWeeks
            data[year_month_week]['Engagement Rate'] = j[5] / j[7] * 100
            data[year_month_week]['Reach'] = j[6] / noOfWeeks
            data[year_month_week]['Impression'] = j[7] / noOfWeeks
            data[year_month_week]['Video Views'] = j[10] / noOfWeeks if j[10] != '-' else j[10]
            data[year_month_week]['View Through Rate'] = j[11] / noOfWeeks if j[11] != '-' else j[11]
            data[year_month_week]['30s View'] = j[12] / noOfWeeks if j[12] != '-' else j[12]
            data[year_month_week]['30s View Through Rate'] = j[13]*100/ noOfWeeks if not j[13] == '-' else '-'
            data[year_month_week]['Video 25%'] = j[14]*100 if not j[14] == '-' else '-'
            data[year_month_week]['Video 50%'] = j[15]*100 if not j[15] == '-' else '-'
            data[year_month_week]['Video 75%'] = j[16]*100 if not j[16] == '-' else '-'
            data[year_month_week]['Video 100%'] = j[17]*100if not j[17] == '-' else '-'
            data[year_month_week]['Page Like'] = j[18]/ noOfWeeks if j[18] != '-' else j[18]
            data[year_month_week]['Photo View'] = j[19] / noOfWeeks if j[19] != '-' else j[19]
            data[year_month_week]['Post Comments'] = j[20] / noOfWeeks if j[20] != '-' else j[20]
            data[year_month_week]['Post Reactions'] = j[21] / noOfWeeks if j[21] != '-' else j[21]
            data[year_month_week]['Post Share'] = j[22] / noOfWeeks if j[22] != '-' else j[22]
            data[year_month_week]['All Clicks'] = '' # post clicks j[23]
            data[year_month_week]['Actual Spent'] = j[25] / noOfWeeks if j[25] != '-' else j[25]


    data2 = []
    for a in data.keys():
        data2.append(data[a])

    df1 = pd.DataFrame(data2)
    df1.to_excel('Cleaned Dataset/Vietnam weekly Digital Result v 1.xlsx', index=False)


if __name__ == '__main__':
    main()