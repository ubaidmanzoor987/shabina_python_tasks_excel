import pandas as pd
import os
import re
import datetime
from googletrans import Translator

def date_format_check(date):
    d = None
    try:
        d = datetime.datetime.strptime(date, '%d.%m.%Y')
    except:
        try:
            d = datetime.datetime.strptime(date, '%d.%m..%Y')
        except:
            d = datetime.datetime.strptime(date, '%m.%d.%Y')
    return d

def main():
    folder_name = 'Data'
    folders = os.listdir(folder_name)
    data = []
    for subfolder in folders:
        files = os.listdir(os.path.join(folder_name, subfolder))
        for excelFile in files:
            print(excelFile)
            filepath = os.path.join(folder_name, subfolder, excelFile)
            wb = pd.read_excel(filepath, engine='openpyxl' , sheet_name = None)
            sheets = wb.keys()
            for sheet in sheets:
                sheetObj = pd.read_excel(filepath, sheet)
                for i, j in sheetObj.iterrows():
                    if not pd.isna(j[3]):
                        prefectureName = j[0]
                        if prefectureName == '北海道':
                            prefectureName = 'Hokkaido'
                        elif prefectureName == '名古屋':
                            prefectureName = 'Nagoya'
                        elif prefectureName == '大阪':
                            prefectureName = 'Osaka'
                        elif prefectureName == '宮城':
                            prefectureName = 'Miyagi'
                        elif prefectureName == '広島':
                            prefectureName = 'Hiroshima'
                        elif prefectureName == '東京':
                            prefectureName = 'Tokyo'
                        elif prefectureName == '福岡':
                            prefectureName = 'Fukuoka'
                        Countryname = j[1]
                        Season = j[2]
                        date = j[3]
                        d = date_format_check(date)
                        di = date_format_check(date).isocalendar()
                        month = datetime.date.strftime(d, "%m")
                        full_year = datetime.date.strftime(d, "%Y")
                        year = di[0]
                        week_num = di[1]
                        year_week = str(year) + '_' + str(week_num)
                        year_month = str(full_year) + '_' + str(month)
                        Chainname = j[4]
                        アオキスーパー
                        if Chainname == 'サニー':
                            Chainname = 'Sunny'
                        elif Chainname == 'ダイキョーバリュー':
                            Chainname = 'Daikyo Value'
                        elif Chainname == 'ニューヨークストア':
                            Chainname = 'New York store'
                        elif Chainname == 'ハローデイ':
                            Chainname = 'Hello day'
                        elif Chainname == 'ボンラパス':
                            Chainname = 'Bon Repas'
                        elif Chainname == 'マックスバリュ':
                            Chainname = 'Maxvalue'
                        elif Chainname == 'マミーズ':
                            Chainname = "Mommy's"
                        elif Chainname == 'マルキョウ':
                            Chainname = "Marukyo"
                        elif Chainname == 'マルショク':
                            Chainname = "Marushoku"
                        elif Chainname == 'レッドキャベツ':
                            Chainname = "Red cabbage"
                        elif Chainname == '西鉄ストア':
                            Chainname = "Nishitetsu store"

                        Storename = j[5]
                        if Storename == '高宮':
                            Storename = 'Takamiya'

                        Storetype = j[6]
                        if Storetype == '食べ頃店':
                            Storetype = 'When to eat'

                        variety = j[7]
                        if variety == 'その他グリーン':
                            variety = 'Other green'
                        elif variety == 'その他ゴールド':
                            variety = 'Other gold'
                        elif variety == 'その他レッド':
                            variety = 'Other red'

                        cutivation_method = j[8]
                        brandname = j[9]
                        if brandname == 'JA全農ふくれん':
                            brandname = 'JA Zen-Noh Fukuren'
                        elif brandname == 'ＪＡ八女':
                            brandname = 'JA Yame'
                        elif brandname == 'キラキラ':
                            brandname = 'Glitter'
                        elif brandname ==  'その他':
                            brandname = 'Other'

                        Placeoforigin = j[10]
                        if Placeoforigin == 'アメリカ':
                            Placeoforigin = 'America'
                        elif Placeoforigin == 'ﾆｭｰｼﾞｰﾗﾝﾄﾞ':
                            Placeoforigin = 'New Zealand'
                        elif Placeoforigin == '愛媛':
                            Placeoforigin = 'Ehime'
                        elif Placeoforigin == '福岡':
                            Placeoforigin = 'Fukuoka'

                        Samplenumber = j[11]
                        Producercode = j[12]
                        SalesunitQuantity = j[13]
                        hardness1 = j[14]
                        hardness2 = j[15]
                        Average_hardness = j[16]
                        Sugar_Content = j[17]
                        Weight = j[18]
                        Size = j[19]
                        UPJPY = j[20]
                        Tray_Equivalent = j[21]
                        KiwiofSalesfloorareaNumberoftrays = j[22]
                        Shelftype = j[23]
                        if Shelftype == '保冷あり':
                            Shelftype = 'With cold storage'

                        Mostpopularfruit = j[24]
                        if Mostpopularfruit == 'いちご':
                            Mostpopularfruit = 'Strawberry'
                        elif Mostpopularfruit == 'バナナ':
                            Mostpopularfruit = 'Banana'
                        elif Mostpopularfruit == 'みかん':
                            Mostpopularfruit = 'Mandarin orange'
                        elif Mostpopularfruit == 'りんご':
                            Mostpopularfruit = 'Apple'

                        mostpopularfruit2 = j[25]
                        if mostpopularfruit2 == 'みかん':
                            mostpopularfruit2 = 'Mandarin orange'

                        mostpopularfruit3= j[26]
                        if mostpopularfruit3 == 'いちご':
                            mostpopularfruit3 = 'Strawberry'
                        Currency = j[27]

                        file = re.sub(r"[^a-zA-Z0-9.]", " ", excelFile)
                        fileName = file.strip()[0:-5]
                        data_obj = [year_week, year_month, full_year, month, week_num, fileName,
                                    date, sheet, variety, cutivation_method,
                                    Average_hardness, Sugar_Content, Weight, Size,
                                    Tray_Equivalent, prefectureName, Countryname,
                                    Season, Chainname, Storename, Storetype, brandname,
                                    Placeoforigin, Samplenumber, Producercode, SalesunitQuantity,
                                    hardness1, hardness2, UPJPY, KiwiofSalesfloorareaNumberoftrays,
                                    Shelftype, Mostpopularfruit, mostpopularfruit2, mostpopularfruit3,
                                    Currency
                                    ]
                        data.append(data_obj)

    df1 = pd.DataFrame(data, columns=['year_week', 'year_month', 'Year', 'Month', 'Week',
                                     'File Name', 'date', 'Prefectures', 'Variety', 'Cultivation Method',
                                      'Average Hardness', 'Sugar Content', 	'Weight',
                                      'Size', 'Tray Equivalent Price (JPY)', 'Perfecture Name', 'Country Name',
                                      'Season', 'Chain Name', 'Store Name', 'Store Type', 'Brand name',
                                      'Place of Origin', 'Sample Number', 'Producer Code', 'Sales unit Quantity',
                                      'Hardness 1', 'Hardness 2', 'U/P JPY', 'Kiwi of Sales floor area (Numberoftrays)',
                                      'Shelf Type', 'Most Popular Fruit', 'Most Popular Fruit 2', 'Most Popular Fruit 3',
                                      'Currency'
                                      ])
    df1.to_excel('Results.xlsx')

if __name__ == '__main__':
    main()