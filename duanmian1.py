

import win32com.client
import os

# 创建 AutoCAD 应用程序对象
acad = win32com.client.Dispatch("AutoCAD.Application.22")
acad.Visible = True
# 获取当前文档对象
doc = acad.ActiveDocument

try:
    doc.SelectionSets.Item("SS1").Delete()
except:
    print("Delete selection failed")
ss = doc.SelectionSets.Add("SS1")
doc.Utility.Prompt("请选择多段线，右键结束\n")

# 获取选择集
ss.SelectOnScreen()


def cal_mid(lst):
    x1 = (lst[0] + lst[2]) / 2
    y1 = (lst[1] + lst[3]) / 2
    return x1, y1


points = []

if len(ss) == 0:
    doc.Utility.Prompt("未选择对象！\n")
else:
    for line in ss:
        name = line.EntityName
        cord_lst = [line.Coordinates[i] for i in range(len(line.Coordinates))]
        x, y = cal_mid(cord_lst)
        z = line.Elevation
        points.append([x, y, z])

print(points)
file_path = os.path.join("G:/目标文件夹/", "elevation_points.txt")
with open(file_path, 'w') as f:
    # 遍历 distances 数组，将点坐标和距离保存到文本文件中
    for point in points:
        f.write("{:.2f},{:.2f},{:.2f}\n".format(point[0], point[1], point[2]))  # {:.2f}, {:.2f},

# 按选择序列号排序
# objs.sort()
#
# # 创建一个用于存储点坐标和距离的数组
# points = sorted(points, key=lambda x: x[0])
# print("new points:", points)
#
# distances = [(0.00, points[0][2])]
# # 遍历点数组，计算距离和高程，并将其保存到 distances 数组中
# for i in range(1, len(points)):
#     pt1 = points[0]
#     pt2 = points[i]
#     dx = pt2[0] - pt1[0]
#     dy = pt2[1] - pt1[1]
#     dz = pt2[2] - pt1[2]
#     distance = math.sqrt(dx * dx + dy * dy)
#     distances.append((distance, pt1[2]))
#     print(distances)
#
# # 打开一个文本文件用于保存横断面
# file_path = os.path.join("G:/2023年工作/09布置断面/", "cross_section.txt")
# with open(file_path, 'w') as f:
#     # 遍历 distances 数组，将点坐标和距离保存到文本文件中
#     for i, dist in enumerate(distances):
#         x, y, z = points[i]
#         f.write("{:.2f}, {:.2f}\n".format(dist[0], z))  # {:.2f}, {:.2f},
#
# # 删除选择集
ss.Delete()

# # 创建 AutoCAD 应用程序对象
# acad = win32com.client.Dispatch("AutoCAD.Application.22")
# acad.Visible = True
# # 获取当前文档对象
# doc = acad.ActiveDocument
#
#
# try:
#     doc.SelectionSets.Item("SS1").Delete()
# except:
#     print("Delete selection failed")
# ss = doc.SelectionSets.Add("SS1")
# doc.Utility.Prompt("请选择多段线，右键结束\n")
#
# # 获取选择集
# ss.SelectOnScreen()
# def cal_mid(lst):
#     x1 = (lst[0] + lst[2]) / 2
#     y1 = (lst[1] + lst[3]) / 2
#     return x1, y1
#
# points = []
#
# if len(ss) == 0:
#     doc.Utility.Prompt("未选择对象！\n")
# else:
#     for line in ss:
#         name = line.EntityName
#         cord_lst =[line.Coordinates[i] for i in range(len(line.Coordinates))]
#         x, y= cal_mid(cord_lst)
#         z = line.Elevation
#         points.append([x,y,z])
#
# print(points)
#
#
# # 按选择序列号排序
# # objs.sort()
#
# # 创建一个用于存储点坐标和距离的数组
# points = sorted(points, key=lambda x: x[0])
# print("new points:", points)
#
# distances = [(0.00, points[0][2])]
# # 遍历点数组，计算距离和高程，并将其保存到 distances 数组中
# for i in range(1, len(points)):
#     pt1 = points[0]
#     pt2 = points[i]
#     dx = pt2[0] - pt1[0]
#     dy = pt2[1] - pt1[1]
#     dz = pt2[2] - pt1[2]
#     distance = math.sqrt(dx * dx + dy * dy)
#     distances.append((distance, pt1[2]))
#     print(distances)
#
# # 打开一个文本文件用于保存横断面
# file_path = os.path.join("G:/2023年工作/09布置断面/", "cross_section.txt")
# with open(file_path, 'w') as f:
#     # 遍历 distances 数组，将点坐标和距离保存到文本文件中
#     for i, dist in enumerate(distances):
#         x, y, z = points[i]
#         f.write("{:.2f}, {:.2f}\n".format(dist[0], z))  # {:.2f}, {:.2f},
#
# # 删除选择集
# ss.Delete()
# -*- coding: utf-8 -*-

'''
@Time    : 2023/3/30 14:58
@Author  : fenglei
@FileName: duanmian1.py
@Software: PyCharm
 
'''
