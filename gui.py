import numpy as np
import pandas as pd
import csv
import time


excel_data_generation = 'dataoutputpy.xlsx' ## the uploaded file must be in the same directory
df = pd.read_excel(excel_data_generation)

x_value = 0
date = 0
generation = 0
consumption = 0
storedenergy = 0
newload = 0
newoverload = 0

fieldnames = ["x_value", "date", "generation", "consumption", "storedenergy", "newload", "newoverload"]

with open('data_to_plot.csv', 'w') as csv_file:
    csv_writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
    csv_writer.writeheader()

while True:

    with open('data_to_plot.csv', 'a') as csv_file:
        csv_writer = csv.DictWriter(csv_file, fieldnames=fieldnames)

        info = {
            "x_value": x_value,
            "date": date,
            "generation": generation,
            "consumption": consumption,
            "storedenergy": storedenergy,
            "newload":newload,
            "newoverload":newoverload
        }

        csv_writer.writerow(info)
        print(x_value, date, generation, consumption,storedenergy,newload,newoverload)

        x_value += 1
        date = df['date'][x_value]
        generation = df['generation'][x_value]
        consumption = df['consumption'][x_value]
        storedenergy = df['storedenergy'][x_value]
        newload = df['newload'][x_value]
        newoverload = df['newoverload'][x_value]

    time.sleep(1)
