import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

df = pd.read_csv('weather_data.csv')
df = df.dropna()

cities = list(set(df['Location']))
cities.sort()

df['Date_Time'] = pd.to_datetime(df['Date'] + ' ' + df['Time'])