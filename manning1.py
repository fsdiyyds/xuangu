# -*- coding: utf-8 -*-

'''
@Time    : 2023/3/29 10:04
@Author  : fenglei
@FileName: manning1.py
@Software: PyCharm
 
'''
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from scipy.integrate import simps
from scipy.interpolate import interp1d

def manning_flow_rate(distance, elevation, water_level, roughness_coeff=0.03, slope=0.001, n_steps=100):
    """
    计算天然河道断面中的水位下的流量
    :param distance: 距离数据（米）
    :param elevation: 高程数据（米）
    :param water_level: 水位高程（米）
    :param roughness_coeff: Manning摩擦系数
    :param slope: 河道比降，米/米为单位
    :param n_steps: 积分时的分数（默认为100）
    :return: 流量（立方米/秒）
    """
    # 计算断面面积
    elevation_water = np.minimum(elevation, water_level)
    section_area = simps(elevation_water, distance)
    # 计算水力半径
    section_perimeter = np.sqrt(np.power(np.diff(distance), 2) + np.power(np.diff(elevation_water), 2))
    section_perimeter = np.append(section_perimeter, section_perimeter[-1])
    section_wetted_perimeter = simps(1 / section_perimeter, distance)
    wetted_radius = section_area / section_wetted_perimeter
    # 计算流量
    flow_rate = roughness_coeff * section_area * pow(wetted_radius, 2 / 3) * pow(slope, 0.5)
    return flow_rate

def generate_table(distance, elevation, roughness_coeff, slope, round_decimals=2):
    """
    计算一系列水位下的流量
    :param distance: 距离数据（米）
    :param elevation: 高程数据（米）
    :param roughness_coeff: Mannning摩擦系数
    :param slope: 河道比降，以米/米为单位
    :param round_decimals: 保留的小数点位数（默认为2）
    :return: 水位-流量表格（Pandas DataFrame格式），包含两列：水位（米）、流量（立方米/秒）
    """
    head_range = np.arange(0, round(max(elevation))+0.05, 0.05)  # 水位范围
    flow_rates = [manning_flow_rate(distance, elevation, h, roughness_coeff, slope) for h in head_range]  # 计算每个水位对应的流量
    # 将水位和流量列表转换为DataFrame格式
    df = pd.DataFrame({'水位（米）': np.round(head_range, round_decimals), '流量（立方米/秒）': np.round(flow_rates, round_decimals)})
    return df

def plot_rating_curve(df):
    """
    绘制水位-流量曲线
    :param df: 水位-流量表格（Pandas DataFrame格式），包含两列：水位（米）、流量（立方米/秒）
    """
    # 创建子图和画布
    fig, ax1 = plt.subplots(figsize=(8, 6))

    # 绘制流量-水位曲线
    ax1.plot(df['流量（立方米/秒）'], df['水位（米）'], 'b-', label='流量-水位')
    ax1.set_xlabel('流量（立方米/秒）', color='blue')
    ax1.set_ylabel('水位（米）', color='blue')
    ax1.tick_params(axis='x', labelcolor='blue')
    ax1.tick_params(axis='y', labelcolor='blue')

    # 添加坐标轴指南和图例
    ax1.grid(True)
    ax1.legend(loc='upper left')


    # 计算给定流量对应的水位值
    def calc_head(flow):
        interp_func = interp1d(df['流量（立方米/秒）'], df['水位（米）'], kind='linear', fill_value="extrapolate")
        return interp_func(flow)

    # 计算给定水位对应的流量值
    def calc_flow(head):
        interp_func = interp1d(df['水位（米）'], df['流量（立方米/秒）'], kind='linear', fill_value="extrapolate")
        return interp_func(head)

    # 打印流量-水位和水位-流量的计算结果
    print('流量为{}m³/s时，水位为{}m'.format(20, calc_head(20)))
    print('水位为{}m时，流量为{}m³/s'.format(15, calc_flow(15)))

    plt.title("水位-流量曲线")
    plt.show()


# 测试
distance = [0, 50, 100, 150, 200, 250, 300, 350, 400]  # 距离数据
elevation = [20, 20, 19, 16, 14, 25, 30, 35, 37]  # 高程数据
roughness_coeff = 0.03  # Manning摩擦系数
slope = 0.001  # 河道比降

# 生成水位-流量表格
table = generate_table(distance, elevation, roughness_coeff, slope)
print(table)

# 绘制水位-流量曲线
plot_rating_curve(table)