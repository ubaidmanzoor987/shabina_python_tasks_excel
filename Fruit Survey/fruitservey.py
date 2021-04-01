import pandas as pd
import os
import re
import datetime

folder_name = '../Translation Data'
file = 'English Translations Dictionary.xlsx'
filepath = os.path.join(folder_name, file)
readexcelfile12 = pd.read_excel(filepath, engine='openpyxl', sheet_name='Translation Dictionary')
translation_dict = {}
for i, j in readexcelfile12.iterrows():
    japanese_language = j[0]
    english_trans = j[1]
    translation_dict[japanese_language] = english_trans

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

def google_translator_my(value):

    if value in translation_dict:
        return translation_dict[value]
    return value

def main():
    folder_name = 'Dataset'
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
                        prefectureName = google_translator_my(str(j[0]))
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
                        month = str(datetime.date.strftime(d, "%m")).zfill(2)
                        full_year = datetime.date.strftime(d, "%Y")
                        year = di[0]
                        week_num = str(di[1]).zfill(2)
                        year_week = str(year) + '_' + week_num
                        year_month = str(full_year) + '_' + month
                        Chainname = google_translator_my(j[4])
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

                        Storename = google_translator_my(j[5])
                        if Storename == '高宮':
                            Storename = 'Takamiya'

                        Storetype = google_translator_my(j[6])
                        if Storetype == '食べ頃店':
                            Storetype = 'When to eat'

                        variety = google_translator_my(j[7])
                        if variety == 'その他グリーン':
                            variety = 'Other green'
                        elif variety == 'その他ゴールド':
                            variety = 'Other gold'
                        elif variety == 'その他レッド':
                            variety = 'Other red'

                        cutivation_method = google_translator_my(j[8])
                        brandname = google_translator_my(j[9])
                        if brandname == 'JA全農ふくれん':
                            brandname = 'JA Zen-Noh Fukuren'
                        elif brandname == 'ＪＡ八女':
                            brandname = 'JA Yame'
                        elif brandname == 'キラキラ':
                            brandname = 'Glitter'
                        elif brandname ==  'その他':
                            brandname = 'Other'

                        Placeoforigin = google_translator_my(j[10])
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
                        Shelftype = google_translator_my(j[23])
                        if Shelftype == '保冷あり':
                            Shelftype = 'With cold storage'

                        Mostpopularfruit = google_translator_my(j[24])
                        if Mostpopularfruit == 'いちご':
                            Mostpopularfruit = 'Strawberry'
                        elif Mostpopularfruit == 'バナナ':
                            Mostpopularfruit = 'Banana'
                        elif Mostpopularfruit == 'みかん':
                            Mostpopularfruit = 'Mandarin orange'
                        elif Mostpopularfruit == 'りんご':
                            Mostpopularfruit = 'Apple'

                        mostpopularfruit2 = google_translator_my(j[25])
                        if mostpopularfruit2 == 'みかん':
                            mostpopularfruit2 = 'Mandarin orange'

                        mostpopularfruit3= google_translator_my(j[26])
                        if mostpopularfruit3 == 'いちご':
                            mostpopularfruit3 = 'Strawberry'
                        Currency = j[27]

                        file = re.sub(r"[^a-zA-Z0-9.]", " ", excelFile)
                        fileName = file.strip()[0:-5]
                        data_obj = [year_week, year_month, full_year, int(month), int(week_num), fileName,
                                    date, sheet, variety, cutivation_method,
                                    Average_hardness, Sugar_Content, Weight, Size,
                                    Tray_Equivalent, prefectureName, Countryname,
                                    Season, Chainname, Storename, Storetype, brandname,
                                    Placeoforigin, Samplenumber, Producercode, SalesunitQuantity,
                                    hardness1, hardness2, UPJPY, KiwiofSalesfloorareaNumberoftrays,
                                    Shelftype, Mostpopularfruit, mostpopularfruit2, mostpopularfruit3,
                                    Currency, variety + cutivation_method
                                    ]
                        data.append(data_obj)

    cols = ['Year_Week', 'Year_Month', 'Year', 'Month', 'Week',
                                     'File Name', 'date', 'Tagged Prefectures', 'Variety', 'Cultivation Method',
                                      'Average Hardness', 'Sugar Content', 	'Weight',
                                      'Size', 'Tray Equivalent Price (JPY)', 'Perfecture', 'Country Name',
                                      'Season', 'Chain Name', 'Store Name', 'Store Type', 'Brand name',
                                      'Place of Origin', 'Sample Number', 'Producer Code', 'Sales unit Quantity',
                                      'Hardness 1', 'Hardness 2', 'U/P JPY', 'Kiwi of Sales floor area (Numberoftrays)',
                                      'Shelf Type', 'Most Popular Fruit', 'Most Popular Fruit 2', 'Most Popular Fruit 3',
                                      'Currency', 'FruitGroup',
                                      ]
    data1 = []
    for d in data:
        row = {}
        for i in range(len(cols)):
            col = cols[i]
            row[col] = d[i]
        data1.append(row)

    df1 = pd.DataFrame(data1)
    df1.to_excel('Cleaned Dataset/FruitServey v 3.xlsx')

if __name__ == '__main__':
    main()