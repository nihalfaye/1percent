import sys
import requests
import json
import pandas as pd
import xlwings as xw
import time

exp = "28-Dec-2023"

def fetch_OC(sym):
    url = "https://www.nseindia.com/api/option-chain-indices?symbol=" + sym
    headers = {
        "Accept-Encoding": "gzip, deflate, br",
        "Accept-Language": "en-US,en;q=0.9",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
    }
    session = requests.Session()
    response = session.get(url, headers=headers).text
    data = json.loads(response)

    ce = []
    pe = []

    OC = data['records']['data']

    for i in OC:
        if i["expiryDate"] == exp:
            ce[0+1] = i['CE']
            pe[0+1] = i['PE']
    
    ce_df = pd.DataFrame.from_dict(ce).transpose()
    ce_df.columns += "_CE"
    pe_df = pd.DataFrame.from_dict(pe).transpose()
    pe_df.columns += "_PE"

    df = pd.concat([ce_df, pe_df], axis=1)

    return df


data1 = fetch_OC("NIFTY")
print(data1)

# data2 = fetch_OC("BANKNIFTY")
# print(data2)

# data3 = fetch_OC("FINNIFTY")
# print(data3)

# data4 = fetch_OC("MIDCPNIFTY")
# print(data4)