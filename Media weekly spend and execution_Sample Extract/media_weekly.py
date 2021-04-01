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
    file = os.path.join(folder_name, 'Media weekly spend and execution_Sample Extract.xlsx')
    print(file)
    tv_data(file)
    digital_data(file)


def tv_data(file):
    data = {}
    readexcelfile = pd.read_excel(file, engine='openpyxl', sheet_name='2019 Zespri TV')
    for i, j in readexcelfile.iterrows():
        if i == 0:
            State = 'Sydney'
        if not (j[0] == 'Planned' or j[0] == 'Delivered' or pd.isna(j[0])):
            State = j[0]
        if not State in data:
            data[State] = {}
        for k in range(1, 11):
            dt = datetime.datetime.strptime('2019/5/19', '%Y/%m/%d')
            dt1 = dt + datetime.timedelta(days=(k - 1) * 7)
            d1 = dt1.isocalendar()
            week_num = d1[1]
            year_week = '2019' + '_' + str(week_num).zfill(2)
            if not year_week in data[State]:
                data[State][year_week] = {}
            data[State][year_week]['Year_Month'] = '2019' + '_' + str(dt.month).zfill(2)
            data[State][year_week]['Year_Week'] = year_week
            data[State][year_week]['Year'] = '2019'
            data[State][year_week]['Week'] = week_num
            data[State][year_week]['Date'] = dt1.date()
            data[State][year_week]['Platform'] = 'TV'
            data[State][year_week]['State'] = State
            if j[0] == 'Planned' and not pd.isna(j[0]):
                data[State][year_week]['Planned TRP'] = j[k]
            if j[0] == 'Delivered' and not pd.isna(j[0]):
                data[State][year_week]['Delivered TRP'] = j[k]

    readexcelfile = pd.read_excel(file, engine='openpyxl', sheet_name='2020 Zespri TV')
    for i, j in readexcelfile.iterrows():
        if i == 0:
            State = 'Sydney'
        if not (j[0] == 'Planned' or j[0] == 'Delivered' or pd.isna(j[0])):
            State = j[0]
        if not State in data:
            data[State] = {}
        for k in range(1, 15):
            dt = datetime.datetime.strptime('2020/5/17', '%Y/%m/%d')
            dt1 = dt + datetime.timedelta(days=(k - 1) * 7)
            d1 = dt1.isocalendar()
            week_num = d1[1]
            year_week = '2020' + '_' + str(week_num).zfill(2)
            if not year_week in data[State]:
                data[State][year_week] = {}
            data[State][year_week]['Year_Month'] = '2020' + '_' + str(dt.month).zfill(2)
            data[State][year_week]['Year_Week'] = year_week
            data[State][year_week]['Year'] = '2020'
            data[State][year_week]['Week'] = week_num
            data[State][year_week]['Date'] = dt1.date()
            data[State][year_week]['Platform'] = 'TV'
            data[State][year_week]['State'] = State
            if j[0] == 'Planned' and not pd.isna(j[0]):
                data[State][year_week]['Planned TRP'] = j[k]
            if j[0] == 'Delivered' and not pd.isna(j[0]):
                data[State][year_week]['Delivered TRP'] = j[k]

    data2 = []
    for a in data.keys():
        for b in data[a].keys():
            data2.append(data[a][b])

    df1 = pd.DataFrame(data2)
    df1.to_excel('Cleaned Dataset/Media weekly Tv Result v 2.xlsx', index=False)

def Data2020(sheetname, start, end, spent, file, data):
    readexcelfile = pd.read_excel(file, engine='openpyxl', sheet_name=sheetname, skiprows=11)
    for i, j in readexcelfile.iterrows():
        if (i >= 2):
            print(j)
            continue
        dt = datetime.datetime.strptime(start, '%d/%m/%Y')
        dt_end = datetime.datetime.strptime(end, '%d/%m/%Y')
        for dt1 in rrule.rrule(rrule.WEEKLY, dtstart=dt, until=dt_end):
            creative = j[1]
            month = dt1.month
            di = dt1.isocalendar()
            week_num = di[1]
            year = dt1.year
            year_month_week_placement = str(year) + '_' + str(month) + str(week_num).zfill(
                2) + '_' + creative
            year_week = str(year) + '_' + str(week_num).zfill(2)
            if not year_month_week_placement in data:
                data[year_month_week_placement] = {}
            data[year_month_week_placement]['Year_Month'] = str(year) + '_' + str(month).zfill(2)
            data[year_month_week_placement]['Year_Week'] = year_week
            data[year_month_week_placement]['Year'] = year
            data[year_month_week_placement]['Month'] = int(month)
            data[year_month_week_placement]['Week'] = int(week_num)
            data[year_month_week_placement]['Platform'] = 'Youtube'
            data[year_month_week_placement]['Menu'] = 'Youtube'
            data[year_month_week_placement]['Topics'] = creative
            data[year_month_week_placement]['Date'] = str(dt1).split(' ')[0]

            data[year_month_week_placement]['Engagement'] = ''
            data[year_month_week_placement]['Engagement Rate'] = ''
            data[year_month_week_placement]['Reach'] = j[7]
            data[year_month_week_placement]['Impression'] = j[3]
            data[year_month_week_placement]['Video Views'] = ''
            data[year_month_week_placement]['View Through Rate'] = ''
            data[year_month_week_placement]['30s View'] = ''
            data[year_month_week_placement]['30s View Through Rate'] = ''
            data[year_month_week_placement]['Video 25%'] = j[12] * 100
            data[year_month_week_placement]['Video 50%'] = j[14] * 100
            data[year_month_week_placement]['Video 75%'] = j[16] * 100
            data[year_month_week_placement]['Video 100%'] = j[18] * 100
            data[year_month_week_placement]['Page Like'] = ''
            data[year_month_week_placement]['Photo View'] = ''
            data[year_month_week_placement]['Post Comments'] = ''
            data[year_month_week_placement]['Post Reactions'] = ''
            data[year_month_week_placement]['Post Share'] = ''
            data[year_month_week_placement]['All Clicks'] = j[4]  # post clikcs j[23]
            data[year_month_week_placement]['Actual Spent'] = spent
    return data
def digital_data(file):
    data = {}
    count = 1
    # Sheet 2018 Digital
    prevBOOSTEDBY = None
    readexcelfile = pd.read_excel(file, engine='openpyxl', sheet_name='2018 Digital', skiprows=20)
    for i, j in readexcelfile.iterrows():
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
                year_month_week = '2018' + '_' + str(month) + str(week_num).zfill(2)
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
                data[year_month_week]['Topics'] = ''
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
                data[year_month_week]['All Clicks'] = j[21]/noOfWeeks if j[21] != '-' else j[21]
                data[year_month_week]['Actual Spent'] = j[23]/noOfWeeks if j[23] != '-' else j[23]
                prevBOOSTEDBY = topics_boosted

    # Sheet 2019 Digital
    year1 = None
    year2 = None
    readexcelfile = pd.read_excel(file, engine='openpyxl', sheet_name='2019 Digital', skiprows = 22 )
    for i, j in readexcelfile.iterrows():
        if (i >= 6 and i <=9) or (i>=14):
            continue
        # if i>=9 and i<=14:
        #     # print(j[1])
        start, end , topic = split_on_hypyhen_and_space_and_new_line(j[1])
        if start.find('Sept'):
            start = start.replace('Sept ', 'Sep')
            start = start.replace('Sept', 'Sep')
        if end.find('Sept'):
            end = end.replace('Sept ', 'Sep')
            end = end.replace('Sept', 'Sep')
        start = start.strip()
        end = end.strip()
        if (start.find('Jan') != -1) and (end.find('Jan') != -1):
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
            if i > 10:
                data[year_month_week]['View Through Rate'] = j[12]/noOfWeeks if j[12] != '-' else j[12]
                data[year_month_week]['30s View'] = j[13]/ noOfWeeks if j[13] != '-' else j[13]
                data[year_month_week]['30s View Through Rate'] = j[14]*100 /noOfWeeks if not j[14] == '-' else '-'
                data[year_month_week]['Video 25%'] = j[15]*100 if not j[15] == '-' else '-'
                data[year_month_week]['Video 50%'] = j[16]*100 if not j[16] == '-' else '-'
                data[year_month_week]['Video 75%'] = j[17]*100 if not j[17] == '-' else '-'
                data[year_month_week]['Video 100%'] =j[18]*100if not j[18] == '-' else '-'
                data[year_month_week]['Page Like'] = j[19] / noOfWeeks if j[19] != '-' else j[19]
                data[year_month_week]['Photo View'] = j[20] / noOfWeeks if j[20] != '-' else j[20]
                data[year_month_week]['Post Comments'] = j[21] / noOfWeeks if j[21] != '-' else j[21]
                data[year_month_week]['Post Reactions'] = j[22] / noOfWeeks if j[22] != '-' else j[22]
                data[year_month_week]['Post Share'] = j[23] / noOfWeeks if j[23] != '-' else j[23]
                data[year_month_week]['All Clicks'] = j[24] / noOfWeeks if j[24] != '-' else j[24]
                data[year_month_week]['Actual Spent'] = j[27] / noOfWeeks if j[27] != '-' else j[27]
            else:
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

    # Sheet 2020 Digital Apr
    prevPlacement = None
    readexcelfile = pd.read_excel(file, engine='openpyxl', sheet_name='2020 Digital APR', skiprows=17)
    for i, j in readexcelfile.iterrows():
        if (i <= 6) or (i > 14):
            # print(j)
            continue
        start = '24 Apr'
        end = '30 Apr'
        dt = None
        dt_end = None
        dt = datetime.datetime.strptime(start + ' 2020', '%d %b %Y')
        dt_end = datetime.datetime.strptime(end + ' 2020', '%d %b %Y')

        noOfWeeks = len(list(rrule.rrule(rrule.WEEKLY, dtstart=dt, until=dt_end)))
        for dt1 in rrule.rrule(rrule.WEEKLY, dtstart=dt, until=dt_end):
            placement = j[2] if j[2] != '' and not pd.isna(j[2]) else prevPlacement
            # print(placement)
            month = dt1.month
            di = dt1.isocalendar()
            week_num = di[1]
            year = dt1.year
            year_month_week_placement = str(year) + '_' + str(month) + str(week_num).zfill(2) + '_' + placement + str(j[3])
            year_week = str(year) + '_' + str(week_num).zfill(2)
            if not year_month_week_placement in data:
                data[year_month_week_placement] = {}
            data[year_month_week_placement]['Year_Month'] = str(year) + '_' + str(month).zfill(2)
            data[year_month_week_placement]['Year_Week'] = year_week
            data[year_month_week_placement]['Year'] = year
            data[year_month_week_placement]['Month'] = int(month)
            data[year_month_week_placement]['Week'] = int(week_num)
            platform1 = j[3]
            if platform1 == 'FB' or platform1 == 'FB News Feed' or platform1 == 'FB Stories':
                platform1 = 'Facebook'
            elif platform1 == 'IG ' or platform1 == 'IG News Feed' or platform1 == 'IG Stories':
                platform1 = 'Instagram'
            data[year_month_week_placement]['Platform'] = platform1
            data[year_month_week_placement]['Menu'] = j[3]
            data[year_month_week_placement]['Topics'] = placement
            data[year_month_week_placement]['Date'] = str(dt1).split(' ')[0]

            if i < 5:
                data[year_month_week_placement]['Engagement'] = j[7]/ noOfWeeks
                data[year_month_week_placement]['Engagement Rate'] = j[7] / j[9] *100 / noOfWeeks
                data[year_month_week_placement]['Reach'] = j[8] / noOfWeeks
                data[year_month_week_placement]['Impression'] = j[9] / noOfWeeks
                data[year_month_week_placement]['Video Views'] = j[12] / noOfWeeks
                data[year_month_week_placement]['View Through Rate'] = j[13] * 100 / noOfWeeks if j[13] == '-' else '-'
                data[year_month_week_placement]['30s View'] = ''
                data[year_month_week_placement]['30s View Through Rate'] = ''
                data[year_month_week_placement]['Video 25%'] = j[14]
                data[year_month_week_placement]['Video 50%'] = j[15]
                data[year_month_week_placement]['Video 75%'] = j[16]
                data[year_month_week_placement]['Video 100%'] = j[17]
                data[year_month_week_placement]['Page Like'] = j[18] / noOfWeeks
                data[year_month_week_placement]['Photo View'] = j[19] / noOfWeeks
                data[year_month_week_placement]['Post Comments'] = j[20] / noOfWeeks
                data[year_month_week_placement]['Post Reactions'] = j[21] / noOfWeeks
                data[year_month_week_placement]['Post Share'] = j[22] / noOfWeeks
                data[year_month_week_placement]['All Clicks'] = '' #post clikcs j[23]
                data[year_month_week_placement]['Actual Spent'] = j[25] / noOfWeeks
            else:
                data[year_month_week_placement]['Engagement'] = j[4] / noOfWeeks
                data[year_month_week_placement]['Engagement Rate'] = j[4] / j[6] *100 / noOfWeeks
                data[year_month_week_placement]['Reach'] = j[5] / noOfWeeks
                data[year_month_week_placement]['Impression'] = j[6] / noOfWeeks
                data[year_month_week_placement]['Video Views'] = j[9] / noOfWeeks
                data[year_month_week_placement]['View Through Rate'] = j[10]*100 / noOfWeeks if not j[10] == '-' else '-'
                data[year_month_week_placement]['30s View'] = ''
                data[year_month_week_placement]['30s View Through Rate'] = ''
                data[year_month_week_placement]['Video 25%'] = j[11]
                data[year_month_week_placement]['Video 50%'] = j[12]
                data[year_month_week_placement]['Video 75%'] = j[13]
                data[year_month_week_placement]['Video 100%'] = j[14]
                data[year_month_week_placement]['Page Like'] = j[15] / noOfWeeks if j[15] != '-' else j[15]
                data[year_month_week_placement]['Photo View'] = j[16] / noOfWeeks if j[16] != '-' else j[16]
                data[year_month_week_placement]['Post Comments'] = j[17] / noOfWeeks if j[17] != '-' else j[17]
                data[year_month_week_placement]['Post Reactions'] = j[18] / noOfWeeks if j[18] != '-' else j[18]
                data[year_month_week_placement]['Post Share'] = j[19] / noOfWeeks if j[19] != '-' else j[19]
                data[year_month_week_placement]['All Clicks'] = '' #j[20]
                data[year_month_week_placement]['Actual Spent'] = j[21] / noOfWeeks if j[21] != '-' else j[21]
                prevPlacement = placement

    # Sheet 2020 Digital Jul
    prevCreative = None
    prevActualSpent = None
    readexcelfile = pd.read_excel(file, engine='openpyxl', sheet_name='2020 Digital Jul', skiprows=15)
    for i, j in readexcelfile.iterrows():
        if (i == 5) or (i == 11) or i>=17:
            # print(j)
            continue
        dt = datetime.datetime.strptime('17/07/2020', '%d/%m/%Y')
        dt_end = datetime.datetime.strptime('27/07/2020', '%d/%m/%Y')
        print(dt, dt_end)
        noOfWeeks = len(list(rrule.rrule(rrule.WEEKLY, dtstart=dt, until=dt_end)))
        for dt1 in rrule.rrule(rrule.WEEKLY, dtstart=dt, until=dt_end):
            creative = j[2] if j[2] != '' and not pd.isna(j[2]) else prevCreative
            # print(placement)
            month = dt1.month
            di = dt1.isocalendar()
            week_num = di[1]
            year = dt1.year
            year_month_week_placement = str(year) + '_' + str(month) + str(week_num).zfill(2) + '_' + creative + str(j[3])
            # print(year_month_week_placement)
            year_week = str(year) + '_' + str(week_num).zfill(2)
            if not year_month_week_placement in data:
                data[year_month_week_placement] = {}
            data[year_month_week_placement]['Year_Month'] = str(year) + '_' + str(month).zfill(2)
            data[year_month_week_placement]['Year_Week'] = year_week
            data[year_month_week_placement]['Year'] = year
            data[year_month_week_placement]['Month'] = int(month)
            data[year_month_week_placement]['Week'] = int(week_num)
            platform1 = j[3]
            if platform1 == 'FB instant Article' or platform1 == 'FB Video Feeds ' or platform1 == 'FB Feeds':
                platform1 = 'Facebook'
            elif platform1 == 'Instagram Feeds' or platform1 == 'Instagram Explore':
                platform1 = 'Instagram'
            actual_spent = j[28] if j[28] != '' or pd.isna(j[28]) else prevActualSpent
            data[year_month_week_placement]['Platform'] = platform1
            data[year_month_week_placement]['Menu'] = j[3]
            data[year_month_week_placement]['Topics'] = creative
            data[year_month_week_placement]['Date'] = str(dt1).split(' ')[0]

            data[year_month_week_placement]['Engagement'] = j[8] / noOfWeeks
            data[year_month_week_placement]['Engagement Rate'] = j[8] / j[10] * 100 / noOfWeeks
            data[year_month_week_placement]['Reach'] = j[9] / noOfWeeks
            data[year_month_week_placement]['Impression'] = j[10] / noOfWeeks
            data[year_month_week_placement]['Video Views'] = j[13] / noOfWeeks
            data[year_month_week_placement]['View Through Rate'] = j[14] * 100 / noOfWeeks
            data[year_month_week_placement]['30s View'] = ''
            data[year_month_week_placement]['30s View Through Rate'] = ''
            data[year_month_week_placement]['Video 25%'] = j[16] * 100
            data[year_month_week_placement]['Video 50%'] = j[18] * 100
            data[year_month_week_placement]['Video 75%'] = j[20] * 100
            data[year_month_week_placement]['Video 100%'] = j[22] * 100
            data[year_month_week_placement]['Page Like'] = ''
            data[year_month_week_placement]['Photo View'] = ''
            data[year_month_week_placement]['Post Comments'] = j[23] / noOfWeeks
            data[year_month_week_placement]['Post Reactions'] = j[24] / noOfWeeks
            data[year_month_week_placement]['Post Share'] = j[25] / noOfWeeks
            data[year_month_week_placement]['All Clicks'] = j[26] / noOfWeeks # post clikcs j[23]
            data[year_month_week_placement]['Actual Spent'] = actual_spent / noOfWeeks

            prevCreative = creative
            prevActualSpent = actual_spent

    # Sheet 2020 Digital Youtube Jun-Aug and Sep-oct
    
    data = Data2020('2020 Digital Youtube Jun-Aug', '23/06/2020', '20/08/2020', 950000, file, data)
    data = Data2020('2020 Digital Youtube Sep-Oct', '09/09/2020', '31/10/2020', 54941.83, file, data)

    data2 = []
    for a in data.keys():
        data2.append(data[a])

    df1 = pd.DataFrame(data2)
    df1.to_excel('Cleaned Dataset/Media weekly Digital Result v 2.xlsx', index=False)


if __name__ == '__main__':
    main()