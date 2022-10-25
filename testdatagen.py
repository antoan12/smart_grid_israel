
import csv
import random
import time
import pandas as pd
import datetime as dt


excel_data_generation = 'Dist_Generating.xlsx' ## the uploaded file must be in the same directory
df = pd.read_excel(excel_data_generation)

excel_data_consumption = 'Consumption.xlsx' ## the uploaded file must be in the same directory
df1 = pd.read_excel(excel_data_consumption)

#real time sim
now = dt.datetime.now()
date_time = now.strftime("%m/%d/2018 %H:00:00")
date = pd.to_datetime(date_time)
inx = df[df['Date'] == date].index.item()


x_value = inx
dates = 0
solargen = 0
gasgen = 0
coalgen = 0
windgen = 0
totgen = 0
household = 0
commercial = 0
factories = 0
totcons = 0

fieldnames = ["x_value", "dates", "solargen", "gasgen", "coalgen", "windgen", "totgen", "household", "commercial", "factories", "totcons"]

with open('data_gen.csv', 'w') as csv_file:
    csv_writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
    csv_writer.writeheader()

while True:

    with open('data_gen.csv', 'a') as csv_file:
        csv_writer = csv.DictWriter(csv_file, fieldnames=fieldnames)

        info = {
            "x_value": x_value,
            "dates": dates,
            "solargen": solargen,
            "gasgen": gasgen,
            "coalgen": coalgen,
            "windgen": windgen,
            "totgen": totgen,
            "household": household,
            "commercial": commercial,
            "factories": factories,
            "totcons": totcons
        }

        csv_writer.writerow(info)
        print(x_value, dates, solargen, gasgen, coalgen, windgen, totgen, household, commercial, factories, totcons)

        x_value += 1
        dates = df['Date'][x_value]
        solargen = df['Solar'][x_value]
        gasgen = df['Gas'][x_value]
        coalgen = df['Coal'][x_value]
        windgen = df['Wind'][x_value]
        totgen = solargen + gasgen + coalgen + windgen
        household = df1['Household'][x_value]
        commercial = df1['Commercial'][x_value]
        factories = df1['Factories'][x_value]
        totcons = household + commercial + factories


    time.sleep(0.01)
