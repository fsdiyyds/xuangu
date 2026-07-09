import math

from numpy import mean
from osgeo import gdal


def readfile(textFile):
    file = open(textFile, "r")
    row = file.readlines()

    list_text = []
    for line in row:
        line = list(line.strip().split(' '))
        s = []
        for i in line:
            s.append(int(i))
        list_text.append(s)
    return list_text


def read_img(filename):
    dataset = gdal.Open(filename)  # 打开文件
    im_width = dataset.RasterXSize  # 栅格矩阵的列数
    im_height = dataset.RasterYSize  # 栅格矩阵的行数
    # im_bands = dataset.RasterCount  # 波段数
    im_geotrans = dataset.GetGeoTransform()  # 仿射矩阵，左上角像素的大地坐标和像素分辨率
    im_proj = dataset.GetProjection()  # 地图投影信息，字符串表示
    im_data = dataset.ReadAsArray(0, 0, im_width, im_height)

    del dataset  # 关闭对象dataset，释放内存
    # return im_width, im_height, im_proj, im_geotrans, im_data,im_bands
    return [im_proj, im_geotrans, im_data, im_width, im_height]


# 遥感影像的存储# 写GeoTiff文件


def write_img(filename, im_proj, im_geotrans, im_data):
    # 判断栅格数据的数据类型
    if 'int8' in im_data.dtype.name:
        datatype = gdal.GDT_Byte
    elif 'int16' in im_data.dtype.name:
        datatype = gdal.GDT_UInt16
    else:
        datatype = gdal.GDT_Float32

    # 判读数组维数
    if len(im_data.shape) == 3:
        # 注意数据的存储波段顺序：im_bands, im_height, im_width
        im_bands, im_height, im_width = im_data.shape
    else:
        im_bands, (im_height, im_width) = 1, im_data.shape  # 没看懂

    # 创建文件时 driver = gdal.GetDriverByName("GTiff")，数据类型必须要指定，因为要计算需要多大内存空间。
    driver = gdal.GetDriverByName("GTiff")
    dataset = driver.Create(filename, im_width, im_height, im_bands, datatype)

    dataset.SetGeoTransform(im_geotrans)  # 写入仿射变换参数
    dataset.SetProjection(im_proj)  # 写入投影

    if im_bands == 1:
        dataset.GetRasterBand(1).WriteArray(im_data)  # 写入数组数据
    else:
        for i in range(im_bands):
            dataset.GetRasterBand(i + 1).WriteArray(im_data[i])

    del dataset


def DirectionMatrix(im_data):  # 计算方向矩阵的函数

    col = len(im_data[0])
    row = len(im_data)
    sqrt2 = math.sqrt(2)

    DirectionMatrix = [[0] * col for i in range(row)]

    for i in range(row):
        for j in range(col):
            M = im_data[i][j]
            E = (M - im_data[i][j + 1]) / 1 if j != col - 1 else -1
            SE = (M - im_data[i + 1][j + 1]) / sqrt2 if i != row - 1 and j != col - 1 else -1
            S = (M - im_data[i + 1][j]) / 1 if i != row - 1 else -1
            SW = (M - im_data[i + 1][j - 1]) / sqrt2 if i != row - 1 and j != 0 else -1
            W = (M - im_data[i][j - 1]) / 1 if j != 0 else -1
            NW = (M - im_data[i - 1][j - 1]) / sqrt2 if j != 0 and i != 0 else -1
            N = (M - im_data[i - 1][j]) / 1 if i != 0 else -1
            NE = (M - im_data[i - 1][j + 1]) / sqrt2 if i != 0 and j != col - 1 else -1

            if M == min(M, E, SE, SW, W, NW, N, NE):
                M = (E + SE + M + SW + W + NW + N + NE) / 7

            M = 0
            M = M if M > S else S
            M = M if M > SE else SE
            M = M if M > N else N
            M = M if M > E else E
            M = M if M > NE else NE
            M = M if M > NW else M
            M = M if M > W else W
            M = M if M > SW else SW

            # print(M)
            if M > 0:
                if M == S:
                    DirectionMatrix[i][j] = 4
                elif M == SE:
                    DirectionMatrix[i][j] = 2
                elif M == N:
                    DirectionMatrix[i][j] = 64
                elif M == E:
                    DirectionMatrix[i][j] = 1
                elif M == NE:
                    DirectionMatrix[i][j] = 128
                elif M == NW:
                    DirectionMatrix[i][j] = 32
                elif M == W:
                    DirectionMatrix[i][j] = 16
                elif M == SW:
                    DirectionMatrix[i][j] = 8
                else:
                    DirectionMatrix[i][j] = 0
            else:
                if i == 0:
                    DirectionMatrix[i][j] = 64
                if i == row - 1:
                    DirectionMatrix[i][j] = 4
                if j == 0:
                    DirectionMatrix[i][j] = 16
                if j == col - 1:
                    DirectionMatrix[i][j] = 1

    return DirectionMatrix


# 计算水流出点的函数
def CalOut(Direction_Matrix):
    col = len(Direction_Matrix[0])
    row = len(Direction_Matrix)

    flag = [[False] * col for i in range(row)]
    for j in range(col):
        i = 0
        if Direction_Matrix[i][j] >= 32:
            flag[i][j] = True
        else:
            flag[i][j] = False
        i = row - 1
        if Direction_Matrix[i][j] >= 2 and Direction_Matrix[i][j] <= 8:
            flag[i][j] = True
        else:
            flag[i][j] = False
    for i in range(0, row):
        j = 0
        if Direction_Matrix[i][j] >= 8 and Direction_Matrix[i][j] <= 32:
            flag[i][j] = True
        else:
            flag[i][j] = False

        j = col - 1
        if Direction_Matrix[i][j] == 1 or Direction_Matrix[i][j] == 2 or Direction_Matrix[i][j] == 128:
            flag[i][j] = True
        else:
            flag[i][j] = False

    return flag


# 判断水流源点的函数
def CalOrigin(Direcion_matrix):
    row = len(Direcion_matrix)
    col = len(Direcion_matrix[0])

    flag = [[False] * col for i in range(row)]
    for j in range(col):
        i = 0
        if Direcion_matrix[i][j] >= 1 and Direcion_matrix[i][j] <= 16:
            flag[i][j] = True
        else:
            flag[i][j] = False
        i = row - 1
        if Direcion_matrix[i][j] >= 16 and Direcion_matrix[i][j] <= 128 or Direcion_matrix[i][j] == 1:
            flag[i][j] = True
        else:
            flag[i][j] = False
    for i in range(row):
        j = 0
        if Direcion_matrix[i][j] == 1 or Direcion_matrix[i][j] == 2 or Direcion_matrix[i][j] == 4 or Direcion_matrix[i][
            j] == 128:
            flag[i][j] = True
        else:
            flag[i][j] = False
        j = col - 1
        if Direcion_matrix[i][j] >= 8 and Direcion_matrix[i][j] <= 32:
            flag[i][j] = True
        else:
            flag[i][j] = False
    return flag


def Add(Direction_Matrix, Origin, Result, i, j) -> int:
    col = len(Direction_Matrix[0])
    row = len(Direction_Matrix)
    if j >= 1:
        if Direction_Matrix[i][j - 1] == 1:
            Result[i][j] += Add(Direction_Matrix, Origin, Result, i, j - 1)
            Result[i][j] += 1
    if i >= 1 and j >= 1:
        if Direction_Matrix[i - 1][j - 1] == 2:
            Result[i][j] += Add(Direction_Matrix, Origin, Result, i - 1, j - 1)
            Result[i][j] += 1
    if i >= 1:
        if Direction_Matrix[i - 1][j] == 4:
            Result[i][j] += Add(Direction_Matrix, Origin, Result, i - 1, j)
            Result[i][j] += 1
    if i >= 1 and j < col - 1:
        if Direction_Matrix[i - 1][j + 1] == 8:
            Result[i][j] += Add(Direction_Matrix, Origin, Result, i - 1, j + 1)
            Result[i][j] += 1
    if j < col - 1:
        if Direction_Matrix[i][j + 1] == 16:
            Result[i][j] += Add(Direction_Matrix, Origin, Result, i, j + 1)
            Result[i][j] += 1
    if i < row - 1 and j < col - 1:
        if Direction_Matrix[i + 1][j + 1] == 32:
            Result[i][j] += Add(Direction_Matrix, Origin, Result, i + 1, j + 1)
            Result[i][j] += 1
    if i < row - 1:
        if Direction_Matrix[i + 1][j] == 64:
            Result[i][j] += Add(Direction_Matrix, Origin, Result, i + 1, j)
            Result[i][j] += 1
    if i < row - 1 and j >= 1:
        if Direction_Matrix[i + 1][j - 1] == 128:
            Result[i][j] += Add(Direction_Matrix, Origin, Result, i + 1, j - 1)
            Result[i][j] += 1
    return Result[i][j]


def Result(Direction_Matrix, ResultMatrix):
    row = len(Direction_Matrix)
    col = len(Direction_Matrix[0])

    origin = [[False] * col for i in range(row)]
    outFlow = [[False] * col for i in range(row)]
    origin = CalOrigin(Direction_Matrix)
    outFlow = CalOut(Direction_Matrix)
    # print(origin)
    for i in range(row):
        for j in range(col):
            if outFlow[i][j] == True:
                Add(Direction_Matrix, origin, ResultMatrix, i, j)
    return ResultMatrix


if __name__ == "__main__":
    # 读取文件
    tif_path = 'AP_12976_FBD_F0690_RT1_HH.tif'
    im_proj, im_geotrans, im_data, im_width, im_height = read_img(tif_path)
    imData = readfile('DEM')
    # 流向矩阵
    DMatrix = DirectionMatrix(imData)
    print(DMatrix)
    # 定义流量矩阵
    row = len(DMatrix)
    col = len(DMatrix[0])
    ResultMatrix = [[0] * col for i in range(row)]
    # print(ResultMatrix)
    Result(DMatrix, ResultMatrix)
    print(ResultMatrix)


