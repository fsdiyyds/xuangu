# -*- coding: utf-8 -*-

'''
@Time    : 2023/3/29 9:53
@Author  : fenglei
@FileName: manning.py
@Software: PyCharm

'''
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from scipy.integrate import trapz
from scipy.interpolate import interp1d

def manning_flow_rate(distance, elevation, water_level, roughness_coeff=0.03, slope=0.001, n_steps=100):
    """
    计算天然河道断面中的水位流量关系
    :param distance: 距离数据，以米为单位
    :param elevation: 高程数据，以米为单位
    :param water_level: 水位高程，以米为单位
    :param roughness_coeff: Mannning摩擦系数
    :param slope: 河道比降，以米/米为单位
    :param n_steps: 积分时的分数，越大越准确（默认为100）
    :return: 流量（单位：立方米/秒）
    """
    # 将距离分段平均化
    distance_avg = np.linspace(distance[0], distance[-1], n_steps)
    # 使用线性插值将高程插值到新的距离列表中
    elevation_func = interp1d(distance, elevation, kind='linear')
    elevation_avg = elevation_func(distance_avg)
    # 计算河道横截面面积
    section_area = np.trapz(y=elevation_avg-water_level, x=distance_avg)
    # 计算水力半径
    section_wetted_perimeter = np.trapz(y=1.0/(elevation_avg-water_level), x=distance_avg)
    wetted_radius = section_area / section_wetted_perimeter
    # 计算流量
    flow_rate = roughness_coeff * section_area * pow(wetted_radius, 2/3) * pow(slope, 0.5)
    return flow_rate

def generate_table(distance, elevation, roughness_coeff, slope, round_decimals=2):
    """
    生成水位-流量表格
    :param distance: 距离数据，以米为单位
    :param elevation: 高程数据，以米为单位
    :param roughness_coeff: Manning摩擦系数
    :param slope: 河道比降，以米/米为单位
    :param round_decimals: 保留的小数点位数（默认为2）
    :return: 水位-流量表格（DataFrame格式）
    """
    # 计算最低和最高水位
    min_head = min(elevation)
    max_head = max(elevation) + 5
    # 初始化结果列表
    head = []
    flow = []
    # 从最低水位开始计算，直到最高水位为止
    for h in np.arange(min_head, max_head, 0.05):
        flow_rate = manning_flow_rate(distance, elevation, h, roughness_coeff, slope)
        head.append(round(h, round_decimals))
        flow.append(round(flow_rate, round_decimals))
    # 将水位和流量列表转换为DataFrame格式
    df = pd.DataFrame({'水位（米）': head, '流量（立方米/秒）': flow})
    return df

def plot_rating_curve(df):
    """
    绘制水位-流量曲线图
    :param df: 水位-流量表格（DataFrame格式）
    """
    plt.plot(df['流量（立方米/秒）'], df['水位（米）'])

# 测试例子
distance = [0, 50, 100, 150, 200, 250, 300, 350, 400]  # 距离数据
elevation = [20, 20, 19, 16, 14, 25, 30, 35, 37]  # 高程数据
roughness_coeff = 0.03  # Mannning摩擦系数
slope = 0.001  # 河道比降

# 生成水位-流量表格
table = generate_table(distance, elevation, roughness_coeff, slope)
print(table)

# 绘制水位-流量曲线图
plot_rating_curve(table)
plt.show()

#
# import numpy as np
# import pandas as pd
# import matplotlib.pyplot as plt
# from scipy.integrate import simps
# from scipy.interpolate import interp1d
#
# def manning_flow_rate(distance, elevation, water_level, roughness_coeff=0.03, slope=0.001, n_steps=100):
#     """
#     计算天然河道断面中的水位下的流量
#     :param distance: 距离数据（米）
#     :param elevation: 高程数据（米）
#     :param water_level: 水位高程（米）
#     :param roughness_coeff: Manning摩擦系数
#     :param slope: 河道比降，米/米为单位
#     :param n_steps: 积分时的分数（默认为100）
#     :return: 流量（立方米/秒）
#     """
#     # 计算断面面积
#     elevation_water = np.minimum(elevation, water_level)
#     section_area = simps(elevation_water, distance)
#     # 计算水力半径
#     section_perimeter = np.sqrt(np.power(np.diff(distance), 2) + np.power(np.diff(elevation_water), 2))
#     section_perimeter = np.append(section_perimeter, section_perimeter[-1])
#     section_wetted_perimeter = simps(1 / section_perimeter, distance)
#     wetted_radius = section_area / section_wetted_perimeter
#     # 计算流量
#     flow_rate = roughness_coeff * section_area * pow(wetted_radius, 2 / 3) * pow(slope, 0.5)
#     return flow_rate
#
# def generate_table(distance, elevation, roughness_coeff, slope, round_decimals=2):
#     """
#     计算一系列水位下的流量
#     :param distance: 距离数据（米）
#     :param elevation: 高程数据（米）
#     :param roughness_coeff: Mannning摩擦系数
#     :param slope: 河道比降，以米/米为单位
#     :param round_decimals: 保留的小数点位数（默认为2）
#     :return: 水位-流量表格（Pandas DataFrame格式），包含两列：水位（米）、流量（立方米/秒）
#     """
#     head_range = np.arange(0, round(max(elevation))+0.05, 0.05)  # 水位范围
#     flow_rates = [manning_flow_rate(distance, elevation, h, roughness_coeff, slope) for h in head_range]  # 计算每个水位对应的流量
#     # 将水位和流量列表转换为DataFrame格式
#     df = pd.DataFrame({'水位（米）': np.round(head_range, round_decimals), '流量（立方米/秒）': np.round(flow_rates, round_decimals)})
#     return df
#
# def plot_rating_curve(df):
#     """
#     绘制水位-流量曲线
#     :param df: 水位-流量表格（Pandas DataFrame格式），包含两列：水位（米）、流量（立方米/秒）
#     """
#     fig = plt.figure(figsize=(8, 6))
#     plt.plot(df['流量（立方米/秒）'], df['水位（米）'])
#     plt.xlabel('流量（立方米/秒）')
#     plt.ylabel('水位（米）')
#     plt.title("水位-流量曲线")
#     plt.show()
#
# # 测试
# distance = [0, 50, 100, 150, 200, 250, 300, 350, 400]  # 距离数据
# elevation = [20, 20, 19, 16, 14, 25, 30, 35, 37]  # 高程数据
# roughness_coeff = 0.03  # Manning摩擦系数
# slope = 0.001  # 河道比降
#
# # 生成水位-流量表格
# table = generate_table(distance, elevation, roughness_coeff, slope)
# print(table)
#
# # 绘制水位-流量曲线
# plot_rating_curve(table)