

# Импорт библиотек
import pandas as pd
import numpy as np
from fuzzywuzzy import process

# Загрузка данных из Excel

df = pd.read_excel(r'C:\Users\Лада\Downloads\docs_final\docs\python\hruschev-pandas.xlsx')
df.head()
# Добавление столбца total (сумма продаж за январь, февраль и март)
df["total"] = df["Jan"] + df["Feb"] + df["Mar"]+ df["April"]

# Словарь соответствия штатов и их аббревиатур
state_to_code = {
    "Rostov.obl": "RO","VERMONT": "VT", "GEORGIA": "GA", "IOWA": "IA", "Armed Forces Pacific": "AP", "GUAM": "GU",
    "KANSAS": "KS", "FLORIDA": "FL", "AMERICAN SAMOA": "AS", "NORTH CAROLINA": "NC", "HAWAII": "HI",
    "NEW YORK": "NY", "CALIFORNIA": "CA", "ALABAMA": "AL", "IDAHO": "ID", "FEDERATED STATES OF MICRONESIA": "FM",
    "Armed Forces Americas": "AA", "DELAWARE": "DE", "ALASKA": "AK", "ILLINOIS": "IL",
    "Armed Forces Africa": "AE", "SOUTH DAKOTA": "SD", "CONNECTICUT": "CT", "MONTANA": "MT", "MASSACHUSETTS": "MA",
    "PUERTO RICO": "PR", "Armed Forces Canada": "AE", "NEW HAMPSHIRE": "NH", "MARYLAND": "MD", "NEW MEXICO": "NM",
    "MISSISSIPPI": "MS", "TENNESSEE": "TN", "PALAU": "PW", "COLORADO": "CO", "Armed Forces Middle East": "AE",
    "NEW JERSEY": "NJ", "UTAH": "UT", "MICHIGAN": "MI", "WEST VIRGINIA": "WV", "WASHINGTON": "WA",
    "MINNESOTA": "MN", "OREGON": "OR", "VIRGINIA": "VA", "VIRGIN ISLANDS": "VI", "MARSHALL ISLANDS": "MH",
    "WYOMING": "WY", "OHIO": "OH", "SOUTH CAROLINA": "SC", "INDIANA": "IN", "NEVADA": "NV", "LOUISIANA": "LA",
    "NORTHERN MARIANA ISLANDS": "MP", "NEBRASKA": "NE", "ARIZONA": "AZ", "WISCONSIN": "WI", "NORTH DAKOTA": "ND",
    "Armed Forces Europe": "AE", "PENNSYLVANIA": "PA", "OKLAHOMA": "OK", "KENTUCKY": "KY", "RHODE ISLAND": "RI",
    "DISTRICT OF COLUMBIA": "DC", "ARKANSAS": "AR", "MISSOURI": "MO", "TEXAS": "TX", "MAINE": "ME"
}

# Функция для нечеткого сопоставления штатов
def convert_state(row):
    if pd.notnull(row['state']):
        state_name = row['state'].upper()
        abbrev = process.extractOne(state_name, choices=state_to_code.keys(), score_cutoff=80)
        if abbrev:
            return state_to_code[abbrev[0]]
    return np.nan

# Добавление столбца abbrev
df.insert(6, "abbrev", np.nan)
df['abbrev'] = df.apply(convert_state, axis=1)

# Группировка данных по столбцу abbrev
df_sub = df[["abbrev", "Jan", "Feb","Mar", "April", "total"]].groupby('abbrev').sum()

# Создание строки с итогами
sum_row = df_sub[["Jan", "Feb", "Mar","April", "total"]].sum()
df_sub_sum = pd.DataFrame(data=sum_row).T

# Переименование индекса для строки с итогами
df_sub_sum.index = ["Total"]

# Объединение данных
final_table = pd.concat([df_sub, df_sub_sum])

# Функция форматирования чисел
def money(x):
    return "${:,.0f}".format(x)

# Применение форматирования
formatted_df = final_table.apply(lambda x: x.map(money) if x.dtype == 'float64' else x)

# Вывод финальной таблицы
print(formatted_df)