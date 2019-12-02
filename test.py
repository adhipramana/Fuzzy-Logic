# -*- coding: utf-8 -*-
"""
Created on Sun Dec  1 14:24:13 2019

@author: Adhip
"""
import xlrd
import math
import numpy as np
from numpy.linalg import inv

#rule
w = 2
MaxIter = 5
x = math.pow(10,(-5))

file = (r'logfuz.xlsx')
excel = xlrd.open_workbook(file, encoding_override='cp1252')

dataset = excel.sheet_by_name('Sheet4')
partisi = excel.sheet_by_name('Sheet3')

datasetM = []
partisiM = []

for i in range(5):
    hasil = []
    for j in range(4):
        hasil.append(dataset.cell(i,j).value)
    datasetM.append(hasil)

for i in range(5):
    hasil = []
    for j in range(2):
        hasil.append(partisi.cell(i,j).value)
    partisiM.append(hasil)

cluster1 = []
cluster2 = []

for i in range(len(partisiM)):
    cluster1.append(partisiM[i][0])
    cluster2.append(partisiM[i][1])

for i in range(MaxIter):
#    cluster1
    miw2c1 = []
    miwX1c1 = []
    miwX2c1 = []
    miwX3c1 = []
    miwX4c1 = []
#    cluster2
    miw2c2 = []
    miwX1c2 = []
    miwX2c2 = []
    miwX3c2 = []
    miwX4c2 = []
       
    for j in range(len(cluster1)):
        miw2c1.append(math.pow(cluster1[j], 2))
        miw2c2.append(math.pow(cluster2[j], 2))
        
    for j in range(len(cluster1)):
        miwX1c1.append(miw2c1[j]*datasetM[j][0])
        miwX2c1.append(miw2c1[j]*datasetM[j][1])
        miwX3c1.append(miw2c1[j]*datasetM[j][2])
        miwX4c1.append(miw2c1[j]*datasetM[j][3])
        
        miwX1c2.append(miw2c2[j]*datasetM[j][0])
        miwX2c2.append(miw2c2[j]*datasetM[j][1])
        miwX3c2.append(miw2c2[j]*datasetM[j][2])
        miwX4c2.append(miw2c2[j]*datasetM[j][3])
    
    Hasilc1 = [sum(miw2c1),sum(miwX1c1), sum(miwX2c1), sum(miwX3c1), sum(miwX4c1)]
    Hasilc2 = [sum(miw2c2),sum(miwX1c2), sum(miwX2c2), sum(miwX3c2), sum(miwX4c2)]
    
    Vc1 = []
    Vc2 = []
    x = 1
    
    for j in range(len(Hasilc1)-1):
        Vc1.append(Hasilc1[j+1]/Hasilc1[0])
        Vc2.append(Hasilc2[j+1]/Hasilc2[0])
    
    XkminV1 = []
    XkminV2 = []
    
    for j in range(len(datasetM)):
        hasil1 = []
        hasil2 = []
        for k in range(len(datasetM[j])):
            hasil1.append(datasetM[j][k] - Vc1[k])
            hasil2.append(datasetM[j][k] - Vc2[k])
        XkminV1.append(hasil1)
        XkminV2.append(hasil2)
    
    XkminV1T = np.transpose(XkminV1)
    XkminV2T = np.transpose(XkminV2)
    
    XkminV1kaliXkminV1T = np.matmul(XkminV1, XkminV1T)
    XkminV2kaliXkminV2T = np.matmul(XkminV2, XkminV2T)
    
    Fbelumdibagic1 = []
    Fbelumdibagic2 = []
    
    for j in range(len(XkminV1kaliXkminV1T)):
        hasil1 = []
        hasil2 = []
        for k in range(len(XkminV1kaliXkminV1T[k])):
            temp1 = 0
            temp2 = 0
            for l in range(len(miw2c1)):
                temp1 += (XkminV1kaliXkminV1T[j][k] * miw2c1[l])
                temp2 += (XkminV2kaliXkminV2T[j][k] * miw2c2[l])
            hasil1.append(temp1)
            hasil2.append(temp2)
        Fbelumdibagic1.append(hasil1)
        Fbelumdibagic2.append(hasil2)
    
    Fc1 = []
    Fc2 = []
    
    for j in range(len(Fbelumdibagic1)):
        hasil1 = []
        hasil2 = []
        for k in range(len(Fbelumdibagic1[j])):
            hasil1.append(Fbelumdibagic1[j][k]/(sum(miw2c1)))
            hasil2.append(Fbelumdibagic2[j][k]/(sum(miw2c2)))
        Fc1.append(hasil1)
        Fc2.append(hasil2)
        
    FInversec1 = inv(np.matrix(Fc1))
    FInversec2 = inv(np.matrix(Fc2))
    
    print(FInversec1)
    
    
    
    
    
    
    
