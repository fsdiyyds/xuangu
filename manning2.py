# -*- coding: utf-8 -*-

'''
@Time    : 2023/3/29 10:35
@Author  : fenglei
@FileName: manning2.py
@Software: PyCharm
 
'''
import numpy as np
import pandas as pd
from scipy.integrate import simps
from scipy.interpolate import interp1d
import tkinter as tk
from tkinter import filedialog
import matplotlib.pyplot as plt

# 设置GUI窗口
root = tk.Tk()
root.title('水位-流量计算')

# 定义全局变量
distance = []
elevation = []
roughness_coeff = 0.03
slope = 0.001
table = pd.DataFrame()

# 定义函数实现文件选择，并更新distance和elevation全局变量
def open_file():
    global distance, elevation
    file_path = filedialog.askopenfilename()
    print('Selected file:', file_path)
    try:
        df = pd.read_csv(file_path, header=None, delimiter=r",")
        distance = df.iloc[:, 0].tolist()
        elevation = df.iloc[:, 1].tolist()
        print('Distance:', distance)
        print('Elevation:', elevation)
    except Exception as e:
        print('Error:', e)
        distance = []
        elevation = []

# 定义函数实现水位-流量曲线绘制


def manning_flow_rate(distance, elevation, water_level, roughness_coeff, slope):
    min_elevation = min(elevation)
    interp_func = interp1d(elevation, distance, kind="linear", fill_value="extrapolate")
    try:
        left_distance = interp_func(water_level)
        right_distance = interp_func(water_level + 0.01)
    except ValueError:
        return 0
    width = abs(right_distance - left_distance)
    depth = water_level - min_elevation
    flow_area = 0
    wetted_perimeter = 0
    prev_dist = interp_func(min_elevation)
    prev_elev = min_elevation
    for elev, dist in zip(elevation, distance):
        if elev > water_level:
            break
        # 计算该段折线下端的切线长度
        dx = abs(prev_dist - dist)
        dh = abs(prev_elev - elev)
        L = (dx ** 2 + dh ** 2) ** 0.5
        prev_dist = dist
        prev_elev = elev
        # 计算湿周和过水面积
        wetted_perimeter += L
        if elev >= water_level - depth:
            dy = abs(prev_elev - water_level + depth)
            flow_area += dx * dy
        else:
            flow_area += (water_level - elev) * dx
    if wetted_perimeter == 0:
        return 0
    hydraulic_radius = flow_area / wetted_perimeter
    velocity = (1 / roughness_coeff) * hydraulic_radius ** (2 / 3) * slope ** (1 / 2)
    flow_rate = velocity * flow_area
    return flow_rate


def plot_rating_curve(table):
    fig, ax1 = plt.subplots()
    ax2 = ax1.twinx()
    ax1.set_xlabel("Water Level (m)")
    ax1.set_ylabel("Flow Rate (m^3/s)")
    ax2.set_ylabel("Wetted Perimeter (m), Flow Area (m^2)")
    ax1.set_title("Rating Curve")
    p1 = ax1.plot(table["water_level"], table["flow_rate"], color="blue", label="Flow Rate")
    p2 = ax2.plot(table["water_level"], table["wetted_perimeter"], color="green", label="Wetted Perimeter")
    p3 = ax2.plot(table["water_level"], table["flow_area"], color="red", label="Flow Area")
    ps = p1 + p2 + p3
    labs = [p.get_label() for p in ps]
    ax1.legend(ps, labs, loc="upper left")
    plt.show()

# 定义函数实现流量计算
def calc_flow():
    global table
    flow = float(flow_entry.get())
    interp_func = interp1d(table['流量（立方米/秒）'], table['水位（米）'], kind='linear', fill_value="extrapolate")
    head = np.around(interp_func(flow), 2)
    head_label.config(text='压力高度为：{}米'.format(head))
    print('Given flow rate: {}m³/s, the head is {}m'.format(flow, head))

# 定义函数实现水位计算
def calc_head():
    global table
    head = float(head_entry.get())
    interp_func = interp1d(table['水位（米）'], table['流量（立方米/秒）'], kind='linear', fill_value="extrapolate")
    flow = np.around(interp_func(head), 2)
    flow_label.config(text='流量为：{}立方米/秒'.format(flow))
    print('Given head: {}m, the flow rate is {}m³/s'.format(head, flow))


def calculate_wetted_perimeter_and_flow_area(distance, elevation, water_level):
    wetted_perimeter = 0
    flow_area = 0
    min_elevation = min(elevation)
    interp_func = interp1d(elevation, distance, kind="linear", fill_value="extrapolate")
    prev_dist = interp_func(min_elevation)
    prev_elev = min_elevation
    for elev, dist in zip(elevation, distance):
        if elev > water_level:
            break
        # 计算该段折线下端的切线长度
        dx = abs(prev_dist - dist)
        dh = abs(prev_elev - elev)
        L = (dx ** 2 + dh ** 2) ** 0.5
        prev_dist = dist
        prev_elev = elev
        # 计算湿周和过水面积
        wetted_perimeter += L
        if elev >= water_level:
            dy = abs(prev_elev - water_level)
            flow_area += (dx * dy)
        else:
            flow_area += ((water_level - elev) * dx)
    return wetted_perimeter, flow_area


# 定义函数实现生成水位-流量表格
def generate_table(distance, elevation, roughness_coeff, slope, min_water_level, max_water_level, step=0.1,
                   dropna=True):
    water_levels = np.arange(min_water_level, max_water_level + step, step)
    flow_rates = []
    wetted_perimeters = []
    flow_areas = []
    for water_level in water_levels:
        flow_rate = manning_flow_rate(distance, elevation, water_level, roughness_coeff, slope)
        flow_rates.append(flow_rate)
        wetted_perimeter, flow_area = calculate_wetted_perimeter_and_flow_area(distance, elevation, water_level)
        wetted_perimeters.append(wetted_perimeter)
        flow_areas.append(flow_area)

    table = pd.DataFrame({
        "water_level": water_levels,
        "flow_rate": flow_rates,
        "wetted_perimeter": wetted_perimeters,
        "flow_area": flow_areas
    })
    if dropna:
        table.dropna(inplace=True)
    table = table.round({"water_level": 2, "flow_rate": 2, "wetted_perimeter": 2, "flow_area": 2})
    return table

# 设置GUI窗口布局
frame1 = tk.Frame(root)
frame1.pack(side=tk.TOP)
file_button = tk.Button(frame1, text='选择文件', command=open_file)
file_button.grid(row=0, column=0)
param_label = tk.Label(frame1, text='河道糙率:')
param_label.grid(row=0, column=1, padx=10)
roughness_entry = tk.Entry(frame1)
roughness_entry.grid(row=0, column=2)
roughness_entry.insert(0, '0.03')
slope_label = tk.Label(frame1, text='河道比降:')
slope_label.grid(row=0, column=3, padx=10)
slope_entry = tk.Entry(frame1)
slope_entry.grid(row=0, column=4)
slope_entry.insert(0, '0.001')
generate_button = tk.Button(frame1, text='计算', command=generate_table)
generate_button.grid(row=0, column=5)

frame2 = tk.Frame(root)
frame2.pack(side=tk.TOP)
head_label = tk.Label(frame2, text='输入流量来计算水位', font=("Arial", 12))
head_label.pack(pady=10)
flow_entry = tk.Entry(frame2)
flow_entry.pack(pady=10)
flow_button = tk.Button(frame2, text='计算水位', command=calc_flow)
flow_button.pack(pady=10)
flow_label = tk.Label(frame2, text='')
flow_label.pack()

frame3 = tk.Frame(root)
frame3.pack(side=tk.TOP)
flow_label = tk.Label(frame3, text='输入水位来计算流量', font=("Arial", 12))
flow_label.pack(pady=10)
head_entry = tk.Entry(frame3)
head_entry.pack(pady=10)
head_button = tk.Button(frame3, text='计算流量', command=calc_head)
head_button.pack(pady=10)
head_label = tk.Label(frame3, text='')
head_label.pack()

root.mainloop()