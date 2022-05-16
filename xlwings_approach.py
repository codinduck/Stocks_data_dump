import requests
import json
import pandas as pd
import xlwings as xw
import time

symbol = "NIFTY"
exp_date = '26-May-2022'
file = xw.Book(r"C:\Users\Rajat Malhotra\Desktop\teststrike2.xlsx")
shee1 = file.sheets('Sheet1')
shee2 = file.sheets('Sheet2')

def oc(sym,exp_date ):
    c1=0
    c2=0
    url = "https://www.nseindia.com/api/option-chain-indices?symbol="+sym
    headers = {"Accept-Encoding" : "gzip, deflate, br",
    "accept-language" : "en-US,en;q=0.9",
    "referer" : "https://www.nseindia.com/get-quotes/derivatives?symbol="+sym,
    "user-agent" : "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/101.0.4951.54 Safari/537.36"}
    response = requests.get(url, headers = headers).text
    data = json.loads(response)
    date_list = data['records']['expiryDates']
    dict1 = {}
    dict2 = {}
    for i in data['records']['data'] :
        if i['expiryDate'] == exp_date:

                dict1[c1] = i['CE']
                c1=c1+1
                dict2[c2] = i['PE']
                c2 = c2+1

    ce_df = pd.DataFrame.from_dict(dict1).transpose()
    ce_df.columns += "_CE"
    pe_df = pd.DataFrame.from_dict(dict2).transpose()
    pe_df.columns += "_PE"
    df = pd.concat([ce_df, pe_df], axis = 1)    
    return (date_list, df)
while True:
     try:
        data = oc(symbol,exp_date)
        print(data)
        shee1.range('A1').value = data[1]
        shee2.range('A1').options(transpose=True).value = data[0]
        time.sleep(10)
     except:
         print("retrying")
         time.sleep(5)

