import sys
import requests
import json
import pandas as pd
import xlwings as xw
import time

# open Excel workbook mentioned and perform action in it 
file = xw.Book("OC_NIFTY.xlsm")
main_sheet = file.sheets("Main Sheet")
exp_sheet = file.sheets["EXP Lists"]
BNF = file.sheets("BANKNIFTY")
NF = file.sheets("NIFTY")
FNF = file.sheets("FINNIFTY")
MDNF = file.sheets("MIDCPNIFTY")

def explist(sym):
    url = "https://www.nseindia.com/api/option-chain-indices?symbol=" + sym
    headers = {
        "Accept-Encoding": "gzip, deflate, br",
        "Accept-Language": "en-US,en;q=0.9",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
    }
    session = requests.Session()
    response = session.get(url, headers=headers).text
    data = json.loads(response)

    exp_list = data['records']['expiryDates']

    return exp_list

def oc(sym,exp):
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

    n = 0
    m = 0

    for i in data['records']['data']:
        if i['expiryDate'] == exp:
            try:
                ce[n] = i['CE']
                n = n + 1
            except:
                pass
            try:
                pe[m] = i['PE']
                m = m + 1
            except:
                pass

    ce_df = pd.DataFrame.from_dict(ce).transpose()
    ce_df.columns += "_CE"
    pe_df = pd.DataFrame.from_dict(pe).transpose()
    pe_df.columns += "_PE"

    df = pd.concat([ce_df, pe_df], axis=1)

    return df

while True:

    #to read specific cell value as Date required for input 
    expdate1 = main_sheet.range("B2").raw_value
    expdate2 = main_sheet.range("B3").raw_value
    expdate3 = main_sheet.range("B4").raw_value
    expdate4 = main_sheet.range("B5").raw_value

    exp1 = explist("BANKNIFTY")
    # exp2 = explist("NIFTY")
    # exp3 = explist("FINNIFTY")
    # exp4 = explist("MIDCPNIFTY")

    exp_sheet.range("A2").options(transpose=True).value = exp1
    # exp_sheet.range("B2").options(transpose=True).value = exp2
    # exp_sheet.range("C2").options(transpose=True).value = exp3Ṣ
    # exp_sheet.range("D2").options(transpose=True).value = exp4


    data1 = oc("BANKNIFTY",expdate1.strftime("%d%#b%#Y"))
    # data2 = oc("NIFTY",expdate2.strftime("%d-%b-%Y"))
    # data3 = oc("FINNIFTY",expdate3.strftime("%d-%b-%Y"))
    # data4 = oc("MIDCPNIFTY",expdate4.strftime("%d-%b-%Y"))

    BNF.range("A1").value = data1
    # NF.range("A1").value = data2
    # FNF.range("A1").value = data3
    # MDNF.range("A1").value = data4            

    time.sleep(60)
