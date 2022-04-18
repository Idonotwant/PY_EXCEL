'''
Author: Wei-Ju-Chen(Idonotwant)
ProjectName: scatter
Description: scatter
Brief_Log:
2022/04/16 Start
2022/04/18 v1 packed
Note: 
python setting.py to make reference file
python scatter.py to make scatter with log x axis(is usually used in frequence response)
'''
from openpyxl import load_workbook,drawing
from os import system,sep
from sys import exit
from openpyxl.chart import (
    ScatterChart,
    Reference,
    Series,
    marker
)

class scatter():
    def __init__(self,**kwgs) -> None:
        self.titlek = kwgs['title']
        self.titlev = None
    def get(title):
        pass

bookpass = input('workbook position(absolute): ')
bookpass = bookpass[1:-1]

def find_col(col): #ex: "A"->1
    col = col.upper()
    col = "".join(reversed(col))
    base = 1
    colValue = 0
    for i in range (len(col)):
        colValue += base*(ord(col[i])-64)
        base *= 26
    return colValue

def splitter(target): #ex: "BA123"->53,123
    i = 0
    while target[i].isalpha():
        i+=1
    return find_col(target[:i]) ,int(target[i:])



# referrence
title = 'B1'
xlabel = 'B2'
xend = 'B3'
ylabel = 'B4'
yend = 'B5'
chartPos = 'B6'

wb = load_workbook(bookpass)
ws = wb.active

blue = drawing.colors.ColorChoice(prstClr='blue')

chart = ScatterChart(scatterStyle='marker')
#scatterStyle=None,varyColors=NONE,dLbls=NONE,extLst=None
chart.title = ws[ws[title].value].value
chart.x_axis.title = ws[ws[xlabel].value].value
chart.y_axis.title = ws[ws[ylabel].value].value
chart.x_axis.scaling.logBase = 10

x_min_col,x_min_row = splitter(ws[xlabel].value)
x_min_row+=1
x_max_col,x_max_row = splitter(ws[xend].value)

y_min_col,y_min_row = splitter(ws[ylabel].value)
y_min_row+=1
y_max_col,y_max_row = splitter(ws[yend].value)

xvalues = Reference(ws, min_col=x_min_col, min_row=x_min_row, max_row=x_max_row)
values = Reference(ws, min_col=y_min_col, min_row=y_min_row, max_row=y_max_row)
series = Series(values, xvalues, title_from_data=False)

chartPosV = ws[chartPos].value

lineProp = drawing.line.LineProperties(solidFill = blue)
series.graphicalProperties.line = lineProp
chart.series.append(series)



chart.style = 25
ws.add_chart(chart,chartPosV)

tmp = bookpass.split('.')
bookoutpass = tmp[0]+"_out."+tmp[1]
wb.save(bookoutpass)
system(bookoutpass)