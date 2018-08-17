##############################################################################
#
# Funciones Performance Report
#
#

import json
import os
import csv
import re

# Process Graphs Category

def ProcessCategory(workbook, category):
    basePath = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    fileName = os.path.join(basePath, 'data', category["file"])

    try:
        fd = open(fileName)
        rd = csv.reader(fd, delimiter="\t", quotechar='"')
    except:
        return workbook

    rowCount = 1
    worksheet = workbook.add_worksheet("Data-" + category["name"])
    for row in rd:
        seriesCount = 1
        timeVal = row[int(category["timeCol"])]

        if (rowCount > 1):
            date, time = timeFormat(timeVal)
        else:
            date = "Date"
            time = "Time"

        posName = "A" + str(rowCount)
        worksheet.write(posName, date)
        posName = "B" + str(rowCount)
        worksheet.write(posName, time)             
        for graph in category["graphs"]:       
            for serie in graph["series"]:
                if (rowCount > 1):
                    value = row[serie["col"]]
                    if value != " ":
                        serValue = serieOperator(float(value), serie["op"], serie["decimals"])
                    else:
                        serValue = 0
                else:
                    serValue = serie["name"]
                
                posName = getColumnName(seriesCount,rowCount)

                worksheet.write(posName[0], serValue)                     
                seriesCount = seriesCount + 1            
        rowCount = rowCount + 1

    worksheet = workbook.add_worksheet("Graphs-" + category["name"])
    seriesCount = 1
    graphPosCount = 1
    for graph in category["graphs"]:        
        chart = workbook.add_chart({'type': 'line'}) 
        chart.set_title ({'name': graph["name"]}) 
          
        for serie in graph["series"]:  
            lastPost, posName = getColumnName(seriesCount,rowCount)              
            categoryName = "=" + "'Data-" + category["name"] + "'!$B$1:$B$" + str(rowCount)     
            valuesName = "=" + "'Data-" + category["name"] + "'!$" + posName + "$1:$" + lastPost
            
            chart.add_series({
                'name':       serie["name"],
                'categories': categoryName,
                'values':     valuesName,
            })

            seriesCount = seriesCount + 1

        graphPos = "A" + str(graphPosCount)
        worksheet.insert_chart(graphPos, chart, {'x_offset': 25, 'y_offset': 30})
        graphPosCount = graphPosCount + 15

    return workbook

# Get Excel Column

def getColumnName(seriesCount,rowCount):
    loopSeriesCount = int(seriesCount / 25) + 1

    nLetter = 66
    posName = ""
    letterName = ""
    for letterCount in range(loopSeriesCount):
        if letterCount == loopSeriesCount - 1:
            nLetter = nLetter + seriesCount - (25*(loopSeriesCount-1))
            letter = chr(nLetter)
            letterName = posName + letter
            posName = posName + letter + str(rowCount)               
        else:                        
            nLetter = 64 + (loopSeriesCount-1)
            letter = chr(nLetter)
            posName = posName + letter

    return posName, letterName


# Time Format

def timeFormat(dateTime):
    regExpr = "(\d\d/\d\d/\d\d\d\d) (\d\d:\d\d:\d\d).\d\d\d"

    match = re.search(regExpr, dateTime)
    date = match.group(1)
    time = match.group(2)

    return date, time

# Series Operator

def serieOperator(value, operation, decimals):
    regExpr = "^(m|d):([0-9]+)"

    match = re.search(regExpr, operation)

    operator = "d"
    factor = 1

    if match is not None:
        if match.group(1):
            operator = match.group(1)

        if match.group(2):
            factor = int(match.group(2))

        if operator == "d":
            value = value / factor

        if operator == "m":
            value = value * factor

    return round(value, decimals)

# Read Config File

def loadConfig(configPath):
    pathFile = os.path.join(configPath, 'report.json')

    try:
        with open(pathFile) as f:
            config = json.load(f)
    except:
        return None

    return config
