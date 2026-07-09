# -*- coding: utf-8 -*-

'''
@Time    : 2023/3/30 14:58
@Author  : fenglei
@FileName: dem1.py
@Software: PyCharm
 
'''
import numpy as np
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D

data = np.genfromtxt('F:/2023年工作/09布置断面/elevation_points.csv', delimiter=',')

xmin = np.min(data[:,0])
xmax = np.max(data[:,0])
ymin = np.min(data[:,1])
ymax = np.max(data[:,1])
resolution = 1
x = np.arange(xmin, xmax, resolution)
y = np.arange(ymin, ymax, resolution)
X, Y = np.meshgrid(x, y)

Z = np.zeros_like(X)
for i in range(len(x)):
    for j in range(len(y)):
        dist = np.sqrt((data[:,0]-x[i])**2 + (data[:,1]-y[j])**2)
        weights = 1/dist
        Z[j,i] = np.sum(weights*data[:,2])/np.sum(weights)

fig = plt.figure()
ax = fig.add_subplot(111, projection='3d')
ax.plot_surface(X, Y, Z, cmap='terrain')
plt.show()