import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

df = pd.read_csv('weather_data.csv')
df = df.dropna()

cities = list(set(df['Location']))
cities.sort()

months = {
    1: 'january',
    2: 'february',
    3: 'march',
    4: 'april',
    5: 'may',
    6: 'june',
    7: 'july',
    8: 'august',
    9: 'september',
    10: 'october',
    11: 'november',
    12: 'december'
}

df['Date_Time'] = pd.to_datetime(df['Date'] + ' ' + df['Time'])

wb = Workbook()

default_sheet = "Sheet"
if default_sheet in wb.sheetnames:
    sheet_to_remove = wb[default_sheet]
    wb.remove(sheet_to_remove)

wb.create_sheet('All month long')
data_sheet = wb['All month long']

for city in cities:
    Path(f'Analyzed/{city}').mkdir(exist_ok=True)

    filtered_df = df[df['Location'] == city]
    filtered_df = filtered_df.sort_values(by='Date')
    
    for i in range(1, 13):
        Path(f'Analyzed/{city}/{months[i].capitalize()}').mkdir(exist_ok=True)
        
        data_sheet['A1'].value = 'Average temperature for the month (°C)'
        data_sheet['A1'].font = Font(bold=True)
        data_sheet['A1'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
        data_sheet['A2'].value = 'Average humidity for the month (%)'
        data_sheet['A2'].font = Font(bold=True)
        data_sheet['A2'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
        data_sheet['A3'].value = 'Average monthly precipitation (mm)'
        data_sheet['A3'].font = Font(bold=True)
        data_sheet['A3'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
        data_sheet['A4'].value = 'Average wind speed for the month (km/h)'
        data_sheet['A4'].font = Font(bold=True)
        data_sheet['A4'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
        filtered_df = df[(df['Location'] == city) & (df['Date_Time'].dt.month == i)]
        
        temperature_avg = filtered_df['Temperature_C'].mean()
        data_sheet['B1'] = temperature_avg
        data_sheet['B1'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
        humididty_avg = filtered_df['Humidity_pct'].mean()
        data_sheet['B2'] = humididty_avg
        data_sheet['B2'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
        precipitation_avg = filtered_df['Precipitation_mm'].mean()
        data_sheet['B3'] = precipitation_avg
        data_sheet['B3'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
        wind_speed_avg = filtered_df['Wind_Speed_kmh'].mean()
        data_sheet['B4'] = wind_speed_avg
        data_sheet['B4'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
        unique_dates = list(set(filtered_df['Date']))
        unique_dates.sort()
        
        j = 1
        
        for date in unique_dates:
            data_sheet['A6'].value = 'Date'
            data_sheet['A6'].font = Font(bold=True)
            data_sheet['A6'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            
            data_sheet['B6'].value = 'Average temperature (°C)'
            data_sheet['B6'].font = Font(bold=True)
            data_sheet['B6'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            
            data_sheet['C6'].value = 'Average humidity (%)'
            data_sheet['C6'].font = Font(bold=True)
            data_sheet['C6'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            
            data_sheet['D6'].value = 'Average precipitation (mm)'
            data_sheet['D6'].font = Font(bold=True)
            data_sheet['D6'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            
            data_sheet['E6'].value = 'Average wind speed (km/h)'
            data_sheet['E6'].font = Font(bold=True)
            data_sheet['E6'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)