# coding: utf-8

# Python2.X encoding wrapper (Windows dedicated processing)
import codecs
import sys
import time
import random
from xlwings import Workbook, Sheet, Range, Chart,RgbColor
sys.stdout = codecs.getwriter('cp932')(sys.stdout)

x=100
y=100

wb = Workbook()
cell=[[random.randint(0,1) for i in xrange(x) ]for j in xrange(y)]
life=[[0 for i in xrange(x) ]for j in xrange(y)]

for i in xrange(x):
    for j in xrange(y):
        cell[0][j]=0
        cell[i][0]=0
        cell[x-1][j]=0
        cell[i][y-1]=0

while(True):

    for i in xrange(1,x):
        for j in xrange(1,y):
            if cell[i][j]==0:
                Range('Sheet1', (i+1,j+1)).color=(255,255,255)   
            else:
                Range('Sheet1', (i+1,j+1)).color=(0,0,0)

    for i in xrange(1,x):
        for j in xrange(1,y):
            if cell[i][j]==1:
                if cell[i-1][j-1]==1:
                    life[i][j]+=1
                if cell[i][j-1]==1:
                    life[i][j]+=1
                if cell[i+1][j-1]==1:
                    life[i][j]+=1
                if cell[i-1][j]==1:
                    life[i][j]+=1
                if cell[i+1][j]==1:
                    life[i][j]+=1
                if cell[i-1][j+1]==1:
                    life[i][j]+=1
                if cell[i][j+1]==1:
                    life[i][j]+=1   
                if cell[i+1][j+1]==1:
                    life[i][j]+=1

    for i in xrange(1,x):
        for j in xrange(1,y):
            
            if life[i][j]==2 or life[i][j]==3:
                cell[i][j]=1
            else:
                cell[i][j]=0
            life[i][j]=0

    
    time.sleep(.1)
