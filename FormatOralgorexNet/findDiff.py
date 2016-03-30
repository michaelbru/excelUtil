import re,os
import openpyxl
from openpyxl.styles import Font, Fill,Color
from openpyxl.cell import get_column_letter
from openpyxl import Workbook
from openpyxl.styles.colors import RED , BLUE , GREEN  
import numpy as np
from  datetime import timedelta
path = r'F:\xyz\projects\Plasim\YahaliAtp\ResultsAtp\Splash-Proof\pendantSplashProof.xlsx'


def openXl(bookname):

            #try: 
            #    os.remove(bookname) 
                
            #except: 
            #    pass
# Create a workbook and add a worksheet.
        if os.path.exists( bookname ):
           return openpyxl.load_workbook(bookname,guess_types=False)
        else:
            raise ValueError

def saveXl(wb,bookname):
    wb.save(bookname)


#open map xml in MALE worksheet
wbMap =  openXl(path)
wsMap = wbMap.active
# assign location of line to start from it 
startLine = 2
# assign location of columns
timeColumn  = 'A'
pressureColumn = 'B'
 
#create new list with PCB pinout - column C   
mapArr1 = []
mapArr2 = []
x= []
for i in range(1,40):
    mapArr1.append( wsMap['A%s'%(i+startLine)].value)  #time
    mapArr2.append( wsMap['B%s'%(i+startLine)].value)  #pressure
for i in range(38):
    t1 = mapArr1[i+1]
    t2 = mapArr1[i]
    #tm1 = timedelta(seconds = t1.second,minutes = t1.minute)
    #tm2 = timedelta(seconds = t2.second,minutes = t2.minute)
    #tm = tm1-tm2
    print(t1,t1.hour,t1.minute,'=>',t2)
    print('~~~~~~~~~~~~')
    t1 = t1.hour*60+t1.minute
    t2 = t2.hour*60+t2.minute
    print(t1,t2)
    print('~~~~~~~~~~~~')
    #print(mapArr1[i+1].minute)
    x.append(t1-t2)
#

#x =np.array(x) 
#xd = np.diff(x)
#print(xd)
print(x)
y =np.array(mapArr2) 
#print(y)
yd = np.diff(y)
#print(yd)
r = yd/x
print(r)
#print(r[r!= np.inf])
avr = np.average(r)
print(avr,"BAR/SEC")