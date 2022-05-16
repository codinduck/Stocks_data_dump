import requests
import json
import pandas as pd




sym = "NIFTY"
exp_date = '26-May-2022'



def oc(sym,exp_date ):
    url = "https://www.nseindia.com/api/option-chain-indices?symbol="+sym
    headers = {"Accept-Encoding" : "gzip, deflate, br",
    "accept-language" : "en-US,en;q=0.9",
    "referer" : "https://www.nseindia.com/get-quotes/derivatives?symbol="+sym,
    "user-agent" : "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/101.0.4951.54 Safari/537.36"}
    response = requests.get(url, headers = headers).text
    data = json.loads(response)
    return data
d = oc(sym,exp_date)
testlist = []
identset = set()


for i in d['records']['data']:
        for j, k in i.items():
            if(j == 'CE'):
            
            
                for l,m in k.items():
                
                    identset.add(l)

#for n in identset:
df = pd.DataFrame()


def create(x):
    for i in d['records']['data']:
        for j, k in i.items():
            if(j == 'CE'):
            
            
                for l,m in k.items():
                
                   # identset.add(l)
                    if(l == x):
                        print(x)
                        testlist.append(m)
                        print(testlist)
                        dframe = pd.DataFrame(index = testlist, columns = x)
                        return dframe


data = create("strikePrice") 

                            
























df = pd.DataFrame()
dframe = pd.DataFrame(index = testlist, columns = identset)
data.to_excel(r"C:\Users\Rajat Malhotra\Desktop\teststrike.xlsx")
