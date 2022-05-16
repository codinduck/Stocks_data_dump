import requests
import json
import pandas as pd
import xlwings as xw
import time


sym = "NIFTY"
exp_date = '26-May-2022'

#wb= xw.Book("teststrike.xlsx")

#file = xw.Book('')

#sh1 = file.sheets('Sheet1')

#sh2 = file.sheets('Sheet2')


#def oc(sym,exp_date ):
    
url = "https://www.nseindia.com/api/option-chain-indices?symbol=NIFTY"
print(url)
headers = {"Accept-Encoding" : "gzip, deflate, br",
"accept-language" : "en-US,en;q=0.9",
"referer" : "https://www.nseindia.com/get-quotes/derivatives?symbol=NIFTY",
"user-agent" : "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/101.0.4951.54 Safari/537.36"}
print(headers)
response = requests.get(url, headers = headers).text
data = json.loads(response)
print(type(data))
exp_list = data['records']['expiryDates']
ce = {}
pe = {}
n=0
m=0
for i in data['records']['data'] :          
    if i['expiryDate'] == '26-May-2022':
        try:
            ce[n] = i['CE']
            n=n+1
            
        except:
            pass
        try:
            pe[m] = i['PE']
            m = m+1
            
        except:
            pass
ce_df = pd.DataFrame.from_dict(ce).transpose()
ce_df.columns += "_CE"
pe_df = pd.DataFrame.from_dict(pe).transpose()
pe_df.columns += "_PE"
df = pd.concat([ce_df, pe_df], axis = 1)
#df.to_excel()
#print(df)
df.to_excel(r"C:\Users\Rajat Malhotra\Desktop\teststrike.xlsx")
#return (exp_list, df)
#while True:
#data = exp_list,df
#print(data)
#sh1.range('A1').value = data[1]
#data = oc(sym,exp_date)
#sh2.range('A1').options(transpose=True).value = data[0]
#time.sleep(60)

        
        
