import pandas as pd
import os
import re
import datetime

from dateutil import rrule


def split_on_hypyhen_and_space(date):
    data1 =None
    try:
        data1 = date.split(' - ')[0]
    except:
        data1 = str(date).split(' ')[0]

    return data1

def main():
    folder_name = 'Dataset'
    file = os.path.join(folder_name, 'Media 17 29.xlsx')
    print(file)
    data = {}

    metric = [
        ("Terrestrial TV - GRPs17-20", "Terrestrial TV"),
        ("Terrestrial TV - Spend17-20", "Terrestrial TV"),
        ("CATV - GRPs17-20", "CATV"),
        ("CATV - Spend17-20", "CATV"),
        ("Terrestrial TV - GRPs20", "Terrestrial TV"),
        ("Terrestrial TV - Spend 2020", "Terrestrial TV"),
        ("CATV - GRPs20", "CATV"),
        ("CATV - Spend 2020", "CATV"),
        ("OOH - Reach", "OOH"),
        ("OOH - Spend", "OOh"),
        ("Instagram - Impressions", "Instagram"),
        ("Instagram - Clicks", "Instagram"),
        ("Instagram - Views", "Instagram"),
        ("Instagram Spend", "Instagram"),
        ("SMR Impressions", "SMR"),
        ("SMR Clicks", "SMR"),
        ("SMR views", "SMR"),
        ("SMR spend", "SMR"),
        ("Youtube - Impressions", "Youtube"),
        ("Youtube - Clicks", "Youtube"),
        ("Youtube - views", "Youtube"),
        ("Youtube Spend", "Youtube"),
        ("Tving - Impression", "Tving"),
        ("Tving - Clicks", "Tving"),
        ("Tving - Views", "Tving"),
        ("Tving - Spend", "Tving"),
        ("Kakao - Impression", "kakao"),
        ("Kakao - Clicks", "kakao"),
        ("Kakao - Views", "kakao"),
        ("Kakao - Spend", "kakao"),
        ("Naver Brandsearch (PC) - Impression", "Naver Brandsearch (PC)"),
        ("Naver Brandsearch (PC) - Click", "Naver Brandsearch (PC)"),
        ("Naver Brandsearch (PC) - Spend", "Naver Brandsearch (PC)"),
        ("Naver Brandsearch (MO) - Impression", "Naver Brandsearch (MO)"),
        ("Naver Brandsearch (MO) - Click", "Naver Brandsearch (MO)"),
        ("Naver Brandsearch (MO) - Spend", "Naver Brandsearch (MO)"),
        ("Facebook - Impressions", "Facebook"),
        ("Facebook - Clicks", "Facebook"),
        ("Facebook - Engagement", "Facebook"),
        ("Facebook Spend" , "Facebook"),
        ("Online Search Impressions", "Online Search"),
        ("Online Search Clicks", "Online Search"),
        ("Online Search Spend", "Online Search"),
        ("Naver Main - ImpressionMasked", "Naver Main"),
        ("Naver Main -  ClicksMasked", "Naver Main"),
        ("Naver Main - ViewsMasked", "Naver Main"),
        ("Naver Main - SpendMasked", "Naver Main"),
        ("Insight - ImpressionMasked", "Insight"),
        ("Insight - ClicksMasked", "Insight"),
        ("Insight  - ViewsMasked", "Insight"),
        ("Insight  - SpendMasked", "Insight"),
        ("Blind - ImpressionMasked", "Blind"),
        ("Blind - ClicksMasked", "Blind"),
        ("Blind - ViewsMasked", "Blind"),
        ("Blind  - SpendMasked", "Blind"),
        ("Kakao + Naver Manin", "kakao")
    ]
    readexcelfile = pd.read_excel(file, engine='openpyxl', sheet_name='Media 17 29')
    for i, j in readexcelfile.iterrows():
        retailer_name = j[1]
        dt = datetime.datetime.strptime(str(j[0]), '%Y-%m-%d %H:%M:%S')
        year = dt.year
        month = dt.month
        d1 = dt.isocalendar()
        week_num = d1[1]
        year_week = str(year) + '_' + str(week_num).zfill(2)
        for k in range(1,57):
            year_week_month = str(year) + '_' + str(week_num).zfill(2) + str(month).zfill(2) + metric[k-1][0]
            if not year_week_month in data:
                data[year_week_month] = {}
            data[year_week_month]['Year_Month'] = str(year) + '_' + str(month).zfill(2)
            data[year_week_month]['Year_Week'] = year_week
            data[year_week_month]['Year'] = str(year)
            data[year_week_month]['Month'] = str(month)
            data[year_week_month]['Week'] = int(week_num)
            data[year_week_month]['Date'] = dt.date()
            data[year_week_month]['Media'] = metric[k-1][0]
            data[year_week_month]['Value'] = j[k]
            data[year_week_month]['Media Channel'] = metric[k-1][1]

    data2 = []
    for a in data.keys():
        data2.append(data[a])


    df1 = pd.DataFrame(data2)
    df1.to_excel('Cleaned Datset/Keorea Media Result1 v 1.xlsx', index=False)


if __name__ == '__main__':
    main()