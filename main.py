import json
import requests
import xlsxwriter
#import openpyxl
#from pandas.io.json import json_normalize
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime, date, timedelta
import numpy as np
# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

def getToken():
    try:
        tk="1000.478c0d7fd353e2078ba91d8be04e01dc.e4c290f9a2b472931d56c134454d008b";


        url = "https://accounts.zoho.com/oauth/v2/token?client_id=1000.YWUJ04M5SN4ZX5EUPCHXJH76VTOGMU&client_secret=a6f27cfa1efd4c0b2a22576db6aa0afb01aca1ea7f&refresh_token="+str(tk)+"&grant_type=refresh_token"
        payload={}
        headers = {
            'Cookie': '_zcsr_tmp=b3834764-4840-43b8-9e37-75bd649b1200; b266a5bf57=57c7a14afabcac9a0b9dfc64b3542b70; iamcsr=b3834764-4840-43b8-9e37-75bd649b1200'
        }
        response = requests.request("POST", url, headers=headers, data=payload)

        print(response.text)
        return response.text
    except Exception as e:
        print("error login",e)
def getReportZeus():
    try:
        valueToken = getToken()
        sp = valueToken.split('"')
        sp2 = "Zoho-oauthtoken "+str(sp[3])
        urlReportZeus = "https://www.site24x7.com/api/reports/performance/455061000000166065?period=50&start_date=2023-01-05T00:00:00-0500&end_date=2023-02-01T00:00:00-0500"
        payloadReportZeus = {}
        headersReportZeus = {
            'Content-Type': 'application/json;charset=UTF-8',
            'Accept': 'application/json; version=2.0',
            #'Authorization': 'Zoho-oauthtoken 1000.b62bfc12c2f344c282489244d69ef450.0710acaceb56344c19a9e14ef3793572',
            'Authorization': sp2,
            'Cookie': 'zaaid=789662975; JSESSIONID=A56F8A04D2DB789D4CEDE4E1923E235F; _zcsr_tmp=380adf49-d7e1-4c16-89aa-760d8a8bc625; aeefa57a7c=38ab671b121c441c11386ff4ef020ae5; s247cname=380adf49-d7e1-4c16-89aa-760d8a8bc625'
        }
        responseReportZeus = requests.request("GET", urlReportZeus, headers=headersReportZeus, data=payloadReportZeus)
        #fileReportZeus = open("reportZeus.json", "x")
        with open("reportZeus.json", "w") as myReportZeus:
            myReportZeus.write(str(responseReportZeus.text))
            print(myReportZeus)
        #print(responseReportZeus.text)
        return responseReportZeus
    except Exception as e:
        print("Error en reporte Zeus" , e)




def jsonToExcel():
    v1=getReportZeus().json()
    print("Test: ", v1)
    v2=v1['data']['chart_data'][0]['0']['OverallCPUChart']['chart_data']
    workbook = xlsxwriter.Workbook('ReportZeusCpu.xlsx')
    cell_format = workbook.add_format({'num_format': '#,##'})
    worksheet = workbook.add_worksheet("CPU")
    v3=v1['data']['chart_data'][1]['0']['OverallMemoryChart']['chart_data']
    worksheet1 = workbook.add_worksheet("Memoria")
    v4=v1['data']['chart_data'][2]['0']['OverallDiskUtilization']['chart_data']
    print('disk: ', v4)
    worksheet2 = workbook.add_worksheet("Disco")
    row = 0
    column = 0
    worksheet.write(row, column, 'Fecha')
    worksheet.write(row, column + 1, 'Porcentaje')
    worksheet1.write(row, column, 'Fecha')
    worksheet1.write(row, column + 1, 'Porcentaje')
    worksheet2.write(row, column, 'Fecha')
    worksheet2.write(row, column + 1, 'Porcentaje Minimo')
    worksheet2.write(row, column + 2, 'Porcentaje Maximo')
    for i in range(len(v2)):
        CpuD,CpuP=v2[i]
        MemD,MemP=v3[i]
        DiskD,DiskP,DiskP2=v4[i]
        #cpuReplace=str(CpuD)
        #print(cpuReplace.find('0500'))
        #cpuReplace.replace("0500",'sapoperro')
        #print(type(cpuReplace))
        CpuF = pd.to_datetime(CpuD, format='%Y/%m/%d')

        MemF = pd.to_datetime(MemD)
        DiskF = pd.to_datetime(DiskD)
        worksheet.write(row+1, column, str(CpuF))
        worksheet.write_number(row+1, column+1,(CpuP))
        worksheet1.write(row+1, column, str(MemF))
        worksheet1.write_number(row + 1, column + 1, (MemP))
        worksheet2.write(row + 1, column, str(DiskF))
        worksheet2.write_number(row + 1, column + 1, (DiskP))
        worksheet2.write(row + 1, column + 2, (DiskP2))
        print(CpuF,CpuP,MemF,MemP)
        row += 1
        rango=row
    chart = workbook.add_chart({'type': 'area'})

    chart.add_series({
        'categories':'=CPU!$A$2:$A$27',#fecha
        'values':'=CPU!$B$2:$B$27',#porcentaje
    })
    chart.set_x_axis({'name':'fechas','date_axis':True,'num_format':'dd/mm/yyyy'})
    chart.set_size({'width': 720, 'height': 576})
    #chart.set_size({'x_scale': 1.5, 'y_scale': 2})
    chart.set_title({'name': 'CPU Utilization'})
    #chart.add_series({'values': '=worksheet!$B$1:$B$'+ str(rango)})
    #chart.add_series({'values': '=Sheet1!$C$1:$C$5'})
    worksheet.insert_chart('D5', chart)

    workbook.close()



#def Graphycal():


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    getToken()
    #getReportZeus()
    jsonToExcel()
 #   Graphycal()
    #print("lol")
# See PyCharm help at https://www.jetbrains.com/help/pycharm/
