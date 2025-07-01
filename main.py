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