import pandas as pd
import statistics
from math import isnan
import csv

letterRange = list(map(chr, range(97, 107)))
sheetsToRead = ['WTG-A', 'WTG-B', 'WTG-C', 'WTG-D', 'WTG-E', 'WTG-F', 'WTG-G', 'WTG-H', 'WTG-I', 'WTG-J']
columnsToRead = ['H:X', 'AA:AB', 'AD:AX', 'AZ', 'BC:CI', 'CK:DD', 'DF:DO', 'DS:DZ']
dataDictionary = {}
momentDictionary = {}


def getDataFromSheet(column, sheet):
    na_values = ["Bad", "Pt Created", "No Data", "Calc Failed", "Tag not Found"]  # add na values if necessary
    data = pd.read_excel('C:\\Users\\tydud\\Downloads\\OSU Hack Ohio Fall 2022 Data_Final.xlsx', sheet_name=sheet,
                         na_values=na_values, skiprows=range(0, 2), usecols=column)
    return data


def findMean(data):
    return statistics.mean(data)


indexLetter = 0
for sheet in sheetsToRead:
    print("Sheet " + sheet)
    for item in columnsToRead:
        print("Column " + item)
        dataDictionary[item] = getDataFromSheet(item, sheet)
        print(dataDictionary[item])
        columnsInDataFrame = len(dataDictionary[item].columns)
        numberOfColumns = list(range(1, columnsInDataFrame - 1))
        print(numberOfColumns)
        print("Number of Columns = {numberOfColumns}".format(numberOfColumns=columnsInDataFrame))
        for index in numberOfColumns:
            dataFile = pd.DataFrame(dataDictionary[item].iloc[:, index])
            momentDictionary[index] = pd.DataFrame(dataDictionary[item].iloc[:, index])
            dataFile.to_csv('Turbine{letter}.csv'.format(letter=letterRange[indexLetter].upper()), index=False)
        with open('Turbine{letter}.txt'.format(letter=letterRange[indexLetter].upper()), 'w') as f:
            writer = csv.writer(f, delimiter='\t')
            writer.writerows(zip(momentDictionary[range(1, 16)]))
    indexLetter += 1

for letter in letterRange:
    data = pd.read_csv('Turbine{letter}.csv'.format(letter=letter.upper()))
    means = open('Turbine{letter}.txt'.format(letter=letter.upper()))
    for column in data:
        average = findMean(list(column))
        data.write(str(average) + "\n")
    data.close()

