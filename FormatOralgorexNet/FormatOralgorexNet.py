import re
import pandas
tst = '''DISC2
         R789 (0603)-2,
         R782 (0805)-1,
         R779 (0805)-1,
         J1 (3110-50SG0BK00A1)-15
DISC2_CURR
         R784 (0402)-2,
         C355 (0402)-1,
         U1 (BGA_100)-G4
DISC2_INH+
         D20 (SMB)-K,
         J22 (3110-40SG0BK00A1)-17,
         J9 (3110-50SG0BK00A1)-33
'''
path =  r'F:\XYZ\PROJECTS\SAC\SACSIMULATOR\PIT292_X1_004.NET'
data = open(path,'r').read()
#print(data)
path = r'F:\xyz\projects\Sac\SACJIG.NET'
dataToFind = open(path,'r').read()
#print(dataToFind)

# retrieve all netlists 
dataToFindNets = [ s for s in dataToFind.split('\n') if re.search('^\S',s) ] # true if start with non-whitespace character
oldData = {}
#print(dataToFindNets)

listDataToFind = dataToFind.split('\n')
for i in range(len(listDataToFind)):
    if re.search('^\S',listDataToFind[i]):
        j = i+1
        oldData[listDataToFind[i]] =[]
        while  re.search('^\s',listDataToFind[j]) and j < len(listDataToFind):
            res = listDataToFind[j].lstrip()
            res = res[:res.find('(')-1]+res[res.find(')')+1:]
            oldData[listDataToFind[i]].append(res)
            j+=1
#print(oldData)

# new data~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
dataList = data.split('\n')
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
#print(newData)


j9 = {}
for k in newData:
    for v in newData[k]:
        if v.find('J9')!=-1:
            j9[k] = v
#j9 = [ k  for x in newData ]
print(j9)

merge ={}
for net in dataToFindNets:
    try:
        merge[net]=  (newData[net], oldData[net])
    except:
            pass
#print (dataToFindNets)
#print(merge)


#for net in dataToFindNets:
#    if 

#dataIndex = [None]*len(dataToFindNets)
#i=0
#for net in dataToFindNets:
#    try:
#        dataIndex[i] = dataList.index(net)
#        i+=1
#    except:
#        pass


#print(dataList[dataIndex[0]])