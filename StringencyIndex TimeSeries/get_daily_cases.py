import pandas as pd
import ssl
ssl._create_default_https_context = ssl._create_unverified_context


def get_daily_covid_cases(country_codes):
    daily_cases = pd.read_csv("https://covid19.who.int/WHO-COVID-19-global-data.csv")
    daily_cases = daily_cases[['Date_reported', 'Country_code', 'New_cases']]
    daily_cases.columns = ['date', 'country_code', 'new_cases']
    daily_cases = daily_cases[daily_cases['country_code'].isin(country_codes)]
    daily_cases['date'] = pd.to_datetime(daily_cases['date'])
    daily_cases['year'] = daily_cases['date'].dt.year
    daily_cases['month'] = daily_cases['date'].dt.month
    daily_cases_agg = daily_cases.groupby(['country_code', 'year', 'month'], as_index=False)['new_cases'].sum()
    daily_cases_agg['datekey'] = daily_cases_agg['year'].astype(str) + daily_cases_agg['month'].astype(str) + '01'
    daily_cases_agg['measure'] = 'daily cases'
    daily_cases_agg['units'] = 'count'
    daily_cases_agg = daily_cases_agg[['datekey', 'country_code', 'measure', 'new_cases', 'units']]
    daily_cases_agg.columns = ['datekey', 'country_code', 'measure', 'value', 'units']
    # daily_cases_agg.to_csv('covid_daily_cases.csv', index=0)
    return daily_cases_agg


# def get_daily_mobility_data(country_codes):
#     mobility = pd.read_csv("https://www.gstatic.com/covid19/mobility/Global_Mobility_Report.csv")


if __name__ == '__main__':
    country_codes = ['JP', 'AU', 'KR', 'SG', 'VN']
    daily_covid_cases = get_daily_covid_cases(country_codes)
