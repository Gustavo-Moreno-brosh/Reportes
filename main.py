import json
import requests
import xlsxwriter
import win32com.client
#import openpyxl
#from pandas.io.json import json_normalize
import pandas as pd
import matplotlib
from datetime import datetime, date, timedelta
import numpy as np
import psutil
# This is a sample Python script.
from win32com import client

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

def getToken():
    try:
        tk="1000.17def166de7b00b2dc0a2c7bfb2832b5.765d4bc77f0c39dad717fff84507f31d";


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
        monitorsDict = {"455061000000172003":"Antares SERVER SENCICO",
            "455061000000172015":"Asbanc SERVER SENCICO",
            "455061000000172027":"Kailash SERVER SENCICO",
            "455061000006172007":"Azure - sencicoprod-12.sanborja.sencico.gob.pe",
            "455061000008627019":"BIM Sencico Server",
            "455061000006172019":"Azure - sencicoprod-0.sanborja.sencico.gob.pe",
            "455061000006172031":"Azure - SIAF-PRI-AWS.sanborja.sencico.gob.pe",
            "455061000006232005":"Azure - sencicoprd-0.sanborja.sencico.gob.pe",
            "455061000000159075":"VeeamBackup SERVER SENCICO",
            "455061000006232017":"Azure - sencicoprod-4.sanborja.sencico.gob.pe",
            "455061000000159089":"Kronos Server SENCICO",
            "455061000006232029":"Azure - sencicoprod-7.sanborja.sencico.gob.pe",
            "455061000006167009":"Azure - sencicoprod-2.sanborja.sencico.gob.pe",
            "455061000006232041":"Azure - sencicoprod-5.sanborja.sencico.gob.pe",
            "455061000006167021":"Azure - sencicoprod-9.sanborja.sencico.gob.pe",
            "455061000000159051":"Oppweb SERVER SENCICO",
            "455061000000159063":"Pegasus SERVER SENCICO",
            "455061000006166007":"Azure - sencicoprod-10.sanborja.sencico.gob.pe",
            "455061000006167033":"Azure - sencicoprod-6.sanborja.sencico.gob.pe",
            "455061000000159015":"Apolo SERVER SENCICO",
            "455061000006172043":"Azure - siaf-trans-aws.sanborja.sencico.gob.pe",
            "455061000000159027":"Fenix SERVER SENCICO",
            "455061000000159039":"Intranet SERVER SENCICO",
            "455061000011629005":"Data Sunrise New 2 Server Sencico",
            "455061000000159003":"AD Administrativo SERVER SENCICO",
            "455061000012478005":"Helpdesk SERVER SENCICO",
            "455061000012350005":"Helpdesk SERVER SENCICO old 2",
            "455061000000166053":"Qlick SERVER SENCICO",
            "455061000006166019":"Azure - sencicoprod-3.sanborja.sencico.gob.pe",
            "455061000006167045":"Azure - sencicoprod-8.sanborja.sencico.gob.pe",
            "455061000000166065":"Zeus SERVER SENCICO",
            "455061000000166077":"PVVM SERVER SENCICO",
            "455061000000166017":"Fsecure SERVER SENCICO",
            "455061000000166029":"Gescont SERVER SENCICO",
            "455061000000166041":"Helpdesk SERVER SENCICO old",
            "455061000000166003":"AD Educativo SERVER SENCICO",
            "455061000000172099":"PPSH-i- SERVER SENCICO",
            "455061000011428007":"Data Sunrise New 1 Server Sencico",
            "455061000000172075":"Kailash-Dev SERVER SENCICO",
            "455061000000172039":"Pcsistel SERVER SENCICO",
            "455061000000149003":"Ultron SERVER SENCICO",
            "455061000000172051":"Pumakana SERVER SENCICO",
            "455061000000172063":"Sextans SERVER SENCICO"
        }
        valueToken = getToken()
        for i in monitorsDict.keys():
            print (i, monitorsDict[i])

            sp = valueToken.split('"')
            sp2 = "Zoho-oauthtoken "+str(sp[3])
            urlReportZeus = "https://www.site24x7.com/api/reports/performance/"+i+"?period=50&start_date=2023-01-05T00:00:00-0500&end_date=2023-02-01T00:00:00-0500"
            payloadReportZeus = {}
            headersReportZeus = {
                'Content-Type': 'application/json;charset=UTF-8',
                'Accept': 'application/json; version=2.0',
                #'Authorization': 'Zoho-oauthtoken 1000.95dfce45c8c1a6b07aeb76d006bbc14e.b1faf931b13c99321a3f7c51e6381ada',
                'Authorization': sp2,
                'Cookie': 'zaaid=789662975; JSESSIONID=A56F8A04D2DB789D4CEDE4E1923E235F; _zcsr_tmp=380adf49-d7e1-4c16-89aa-760d8a8bc625; aeefa57a7c=38ab671b121c441c11386ff4ef020ae5; s247cname=380adf49-d7e1-4c16-89aa-760d8a8bc625'
            }
            responseReportZeus = requests.request("GET", urlReportZeus, headers=headersReportZeus, data=payloadReportZeus)
            #fileReportZeus = open("reportZeus.json", "x")
            with open(monitorsDict[i]+".json", "w") as myReportZeus:
                myReportZeus.write(str(responseReportZeus.text))
                print(myReportZeus)
            #print(responseReportZeus.text)
            print (responseReportZeus)
            nomArchivo = monitorsDict[i]
            #return responseReportZeus
            jsonToExcel(responseReportZeus, nomArchivo)
    except Exception as e:
        print("Error en reporte Zeus" , e)




def jsonToExcel(responseReportZeus,nomArchivo):
    #v1=getReportZeus().json()
    v1=responseReportZeus.json()
    print("Test: ", v1)
    v2=v1['data']['chart_data'][0]['0']['OverallCPUChart']['chart_data']
    workbook = xlsxwriter.Workbook(nomArchivo+'.xlsx')
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
    worksheet2.write(row, column + 1, 'Espacio Libre')
    worksheet2.write(row, column + 2, 'Espacio Usado')
    x = 0
    x2 = 0
    for i in range(len(v2)):

        CpuD,CpuP=v2[i]
        MemD,MemP=v3[i]
        DiskD,DiskP,DiskP2=v4[i]
        disksum1 = v1['data']['chart_data'][2]['0']['OverallDiskUtilization']['chart_data'][i][1]
        disksum2 = v1['data']['chart_data'][2]['0']['OverallDiskUtilization']['chart_data'][i][2]

        print(disksum1, disksum2)
        #cpuReplace=str(CpuD)
        #print(cpuReplace.find('0500'))
        #cpuReplace.replace("0500",'sapoperro')
        #print(type(cpuReplace))
        CpuPiv = CpuD.split('T')
        CpuPiv2 = str(CpuPiv[0])
        CpuF = pd.to_datetime(CpuPiv2, format='%Y%m%d', errors='ignore')
        MemPiv = MemD.split('T')
        MemPiv2 = str(MemPiv[0])
        MemF = pd.to_datetime(MemPiv2, format='%Y%m%d', errors='ignore')
        DiskPiv = DiskD.split('T')
        DiskPiv2 = str(DiskPiv[0])
        DiskF = pd.to_datetime(DiskPiv2, format='%Y%m%d', errors='ignore')
        worksheet.write(row+1, column, str(CpuF))
        worksheet.write_number(row+1, column+1,(CpuP))
        worksheet1.write(row+1, column, str(MemF))
        worksheet1.write_number(row +1, column + 1, (MemP))
        worksheet2.write(row +1, column, str(DiskF))
        worksheet2.write_number(row +1, column + 1, (DiskP))

        worksheet2.write(row +1, column + 2, (DiskP2))
        print(CpuF,CpuP,MemF,MemP)

        x=disksum1+x
        x2=disksum2+x2

        row += 1
    porMin = x/row
    porMax = x2/row
    worksheet2.write(row +1, column + 1, (porMin))
    worksheet2.write(row +1, column + 2, (porMax))
    workbook.close()
    #Inicio Creacion GRAFICO CPU
    #chart = workbook.add_chart({'type': 'area',})

    #chart.add_series({
    #   'categories':'=CPU!$A$31:$A$57',#fecha
    #   'values':'=CPU!$B$31:$B$57',#porcentaje
    # })
    #chart.set_x_axis({'name':'fechas','date_axis':True})
    #chart.set_y_axis({'name':'porcentaje','date_axis':True})
    #chart.set_legend({'none': True})
    #chart.set_size({'width': 620, 'height': 476})
    #chart.set_size({'x_scale': 1.5, 'y_scale': 2})
    #chart.set_title({'name': 'CPU Utilization'})
    #chart.add_series({'values': '=worksheet!$B$1:$B$'+ str(rango)})
    #chart.add_series({'values': '=Sheet1!$C$1:$C$5'})
    #worksheet.insert_chart('A1', chart)
    #Fin Creacion GRAFICO CPU

    # Inicio Creacion GRAFICO Memoria
    #chart = workbook.add_chart({'type': 'area', })
    #chart.add_series({
    #   'categories': '=Memoria!$A$31:$A$57',  # fecha
    #   'values': '=Memoria!$B$31:$B$57',  # porcentaje
    #})
    #chart.set_x_axis({'name': 'fechas', 'date_axis': True})
    #chart.set_y_axis({'name': 'porcentaje', 'date_axis': True})
    # chart.set_legend({'none': True})
    #chart.set_size({'width': 620, 'height': 476})
    # chart.set_size({'x_scale': 1.5, 'y_scale': 2})
    #chart.set_title({'name': 'Memoria Utilization'})
    # chart.add_series({'values': '=worksheet!$B$1:$B$'+ str(rango)})
    # chart.add_series({'values': '=Sheet1!$C$1:$C$5'})
    """# worksheet1.insert_chart('A1', chart)
    # Fin Creacion GRAFICO Memoria

    # Inicio Creacion GRAFICO Disco
    chart2 = workbook.add_chart({'type': 'pie', })
    chart2.add_series({
        'name': 'visajes',
        'categories': '=Disco!$B$58:$C$58' , # media libre
        'values': ['Disco', 57, 1, 57, 2]  # media usado
    })


    # chart.set_size({'x_scale': 1.5, 'y_scale': 2})

    # chart.add_series({'values': '=worksheet!$B$1:$B$'+ str(rango)})
    # chart.add_series({'values': '=Sheet1!$C$1:$C$5'})
    worksheet2.insert_chart('A1', chart2)
    # Fin Creacion GRAFICO Disco
    workbook.close()"""
def excelToPdf():
    try:
        excel = client.Dispatch("Excel.Application")
        #Read Excel File
        sheets = excel.Workbooks.Open('C:/Users/WILDC/PycharmProjects/ReportsSite24x7/ReportZeusCpu.xlsx')
        work_sheets = sheets.Worksheets[0]
        # Convert into PDF File
        work_sheets.ExportAsFixedFormat(0, "C:/Users/WILDC/PycharmProjects/ReportsSite24x7/reporteZeus.pdf")
        #matacion()
    except Exception as e:
        print("Error creando pdf", e)

def matacion():
    for proc in psutil.process_iter():
        if proc.name() == "EXCEL.EXE":
            proc.kill()
            print("dddddddd")
            f="muerto"
#def Graphycal():


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    #getToken()
    getReportZeus()
    #jsonToExcel()
    #excelToPdf()
 #   Graphycal()
    #print("lol")
# See PyCharm help at https://www.jetbrains.com/help/pycharm/
