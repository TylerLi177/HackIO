import pandas as pd
import statistics
import numpy as np

letterRange = list(map(chr, range(97, 107)))
sheetsToRead = ['WTG-A','WTG-B','WTG-C','WTG-D','WTG-E','WTG-F','WTG-G','WTG-H','WTG-I','WTG-J']
columnsToRead = ['H:X', 'AA:AB', 'AD:AX', 'AZ', 'BC:CI', 'CK:DD', 'DF:DO', 'DS:DZ']
dataDictionary = {}


def getDataFromSheet(column, sheet):
    na_values = ["Bad", "Pt Created", "No Data", "Calc Failed", "Tag not Found"]  # add na values if necessary
    data = pd.read_excel(r'C:\\Users\\tydud\\Downloads\\OSU Hack Ohio Fall 2022 Data_Final.xlsx', sheet_name=sheet,
                         na_values=na_values, skiprows=range(0, 2), usecols=column)

    return data


def findMean(data):
    return statistics.mean(data)


# TurbineA = open('TurbineA.txt', 'w')
# TurbineB = open('TurbineB.txt', 'w')
# TurbineC = open('TurbineC.txt', 'w')
# TurbineD = open('TurbineD.txt', 'w')
# TurbineE = open('TurbineE.txt', 'w')
# TurbineF = open('TurbineF.txt', 'w')
# TurbineG = open('TurbineG.txt', 'w')
# TurbineH = open('TurbineH.txt', 'w')
# TurbineI = open('TurbineI.txt', 'w')
# TurbineJ = open('TurbineJ.txt', 'w')
indexLetter = 0
for sheet in sheetsToRead:
    dataFile = 'Turbine{letter}.txt'.format(letter=letterRange[indexLetter])
    print(sheet)
    for item in columnsToRead:
        print(item)
        dataDictionary[item] = getDataFromSheet(item, sheet)
        print(dataDictionary[item])
        columnsInDataFrame = len(dataDictionary[item].columns)
        numberOfColumns = list(range(1, int(columnsInDataFrame - 1)))
        print(numberOfColumns)
        print("Number of Columns = {numberOfColumns}".format(numberOfColumns=columnsInDataFrame))
        for index in numberOfColumns:
            print(index)
            # dataFile.write(str(dataDictionary[item].iloc[:, index]))
            np.savetxt(dataFile, dataDictionary[item].iloc[:, index], fmt='%d')
    # dataFile.close()
    indexLetter += 1
        

# TurbineA.close()
# TurbineB.close()
# TurbineC.close()
# TurbineD.close()
# TurbineE.close()
# TurbineF.close()
# TurbineG.close()
# TurbineH.close()
# TurbineI.close()
# TurbineJ.close()
