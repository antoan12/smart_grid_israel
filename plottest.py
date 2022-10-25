import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.animation import FuncAnimation
import numpy as np


plt.style.use('fivethirtyeight')

fig, ax = plt.subplots(2, 2, figsize=(14, 8))
fig.autofmt_xdate(rotation=35)

plt.rc('xtick', labelsize=2) #fontsize of the x tick labels

plt.suptitle('Tips Dataset')

def animate(frame):
    data = pd.read_csv('data_to_plot.csv')
    x = data['date']
    y1 = data['generation']
    y2 = data['consumption']
    overload = y2-y1
    u = overload.copy()
    l = overload.copy()
    u[u <= -0.01] = np.nan
    l[l >= 0.01] = np.nan
    y3 = data['storedenergy']
    y4 = data['newload']
    y5 = data['newoverload']

    ax[0][0].plot(x, y1, lw=2, color='blue', label='Generation')
    ax[0][0].plot(x, y2, lw=2, color='red', label='Consumption')
    ax[0][0].set(xlabel="Date", ylabel="Gen VS Cons [MWh]")

    ax[0][1].plot(x, y3, lw=2, color='blue', label='Stored Energy')
    ax[0][1].set(xlabel="Date", ylabel="Stored Energy [MWh]")

    ax[1][0].plot(x, l, lw=2, color='blue', label='Overload')
    ax[1][0].plot(x, u, lw=2, color='red')
    ax[1][0].set(xlabel="Date", ylabel="Overload [MWh]")

    ax[1][1].plot(x, y2, lw=2, color='red', label='Prev Load')
    ax[1][1].plot(x, y2-y5, lw=2, color='green', label='New Load')
    ax[1][1].set(xlabel="Date", ylabel="Prev Load Vs New Load [MWh]")

    plt.tight_layout()
    ax[0][0].set_xlim(left=frame-12, right=frame)
    ax[0][1].set_xlim(left=frame-12, right=frame)
    ax[1][0].set_xlim(left=frame-12, right=frame)
    ax[1][1].set_xlim(left=frame-12, right=frame)
    plt.gcf().autofmt_xdate()
    # plt.cla()

ani = FuncAnimation(plt.gcf(), animate, interval=800, frames=10000)

plt.tight_layout()
plt.show()
