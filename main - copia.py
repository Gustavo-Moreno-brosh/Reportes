import json
import requests
import xlsxwriter
#import openpyxl
#from pandas.io.json import json_normalize
import pandas as pd
from datetime import datetime, date, timedelta

# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

def getToken():
    try:
        tk="1000.e979ba3497641da97ba882d001f23bc9.6b43108ac62cf8aaa0338feb0d97d3ab";


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
    worksheet = workbook.add_worksheet("CPU")
    row = 0
    column = 0
    worksheet.write(row, column, 'Fecha')
    worksheet.write(row, column + 1, 'Porcentaje')
    for i in range(len(v2)):
        f1,v1=v2[i]
        f2 = pd.to_datetime(f1)
        worksheet.write(row+1, column, str(f2))
        worksheet.write(row+1, column+1, (str(v1)))
        print(f2,v1)
        row += 1
    workbook.close()



# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    getToken()
    #getReportZeus()
    jsonToExcel()
    #print("lol")
# See PyCharm help at https://www.jetbrains.com/help/pycharm/
