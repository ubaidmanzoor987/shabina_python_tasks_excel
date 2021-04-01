import datetime

import pandas as pd
import ssl
ssl._create_default_https_context = ssl._create_unverified_context
import os

destination_file = './Cleaned Dataset/stringencyIndex.csv'
source_file = "https://raw.githubusercontent.com/OxCGRT/covid-policy-tracker/master/data/timeseries/stringency_index.csv"

conntry_dict={'JPN': 'JP', 'AUS': 'AU', 'KOR': 'KR', 'SGP': 'SG', 'VNM': 'VN'}

def get_daily_covid_cases(country_codes):
    daily_cases = pd.read_csv(source_file)
    data = {}
    startFrom = None
    if os.path.exists(destination_file):
        existing_data = pd.read_csv(destination_file)
        data1 = existing_data.tail(1)
        dateKey = data1['dateKey']
        dateKey = str(dateKey[dateKey.keys()[0]])
        print(dateKey)
        startFrom = datetime.datetime.strptime(dateKey, '%Y%m%d')
        startFrom = startFrom.replace(month=startFrom.month + 1)

    for i, j in daily_cases.iterrows():
        if j['country_code'] in country_codes:
            country_code = j['country_code']
            for k in j[3:].keys():
                if not country_code in data:
                    data[country_code] = {}

                dateObj = datetime.datetime.strptime(k, '%d%b%Y')
                if(startFrom and dateObj < startFrom):
                    continue
                dateObj = dateObj.replace(day=1)
                dateStr = dateObj.strftime('%Y%m%d')
                if not dateStr in data[country_code]:
                    print(conntry_dict[country_code])
                    data[country_code][dateStr] = {
                        "dateKey": dateStr,
                        "country_code": conntry_dict[country_code],
                        "measure": "Stringency of Measures",
                        "value": 0,
                        "count": 0,
                        "units": "Index"
                    }
                jK = j[k] if not pd.isna(j[k]) else 0
                data[country_code][dateStr]["value"] += jK
                data[country_code][dateStr]["count"] += 1

    data2 = []
    for s in data.keys():
        for a in data[s].keys():
            data[s][a]["value"] /= data[s][a]["count"]
            del data[s][a]["count"]
            data2.append(data[s][a])
    df = pd.DataFrame(data2)
    if startFrom:
        df.to_csv(destination_file, mode='a', header=False, index=False)
    else:
        df.to_csv(destination_file, mode='a', index=False)


if __name__ == '__main__':
    country_codes = ['JPN', 'AUS', 'KOR', 'SGP', 'VNM']
    daily_covid_cases = get_daily_covid_cases(country_codes)
