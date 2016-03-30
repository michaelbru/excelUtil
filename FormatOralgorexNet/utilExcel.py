import re,os
import openpyxl
from openpyxl.styles import Font, Fill,Color
from openpyxl.cell import get_column_letter
from openpyxl import Workbook
from openpyxl.styles.colors import RED , BLUE , GREEN  


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



def getNet(path):   
    #open netlist
    data = open(path,'r').read()
    dataList = data.split('\n')
    return dataList

def getFormattedNetList(dataList):
     #format netlist as dictionary
    newData = {}
    for i in range(len(dataList)):
        if re.search('^\S',dataList[i]):
            j = i+1
            newData[dataList[i]] =[]
            while  j < len(dataList): 
                if re.search('^\s',dataList[j]):
                    res = dataList[j].lstrip()
                    if res!='':
                        if res[-1]==',':
                         res = res[:res.find('(')-1]+res[res.find(')')+1:-1]
                        else:
                            res = res[:res.find('(')-1]+res[res.find(')')+1:]
                        newData[dataList[i]].append(res)
                    j+=1
                else:
                    break

    return newData

