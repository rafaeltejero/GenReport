##############################################################################
#
# Generate Performance Report
#
# https://xlsxwriter.readthedocs.io
#

import datetime
import os
import xlsxwriter
import funcReport

if __name__ == '__main__':
    startTime = datetime.datetime.now()

    print(startTime)

    basePath = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    configPath = os.path.join(basePath, 'config')
    config = funcReport.loadConfig(configPath)

    outFile = os.path.join(basePath, 'output', 'report.' + startTime.strftime("%Y-%m-%d.%H-%M-%S") + '.xlsx')
    #outFile = os.path.join(basePath, 'output', 'report.' +  startTime.strftime("%Y-%m-%d") + '.xlsx')

    if (not config):
        print("Config file not exist")
        exit()

    workbook = xlsxwriter.Workbook(outFile)

    for category in config["category"]:
        workbook = funcReport.ProcessCategory(workbook, category)

    workbook.close()

    print(datetime.datetime.now())

