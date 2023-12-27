import requests
import json
import pandas as pd
import xlwings as xw
from datetime import datetime, timedelta

# open excel file and locate tap name Historical data 
file = xw.Book("OC_NIFTY.xlsm")
histdata_sheet = file.sheets("Historical Data")

# track todays time and exact 1 year before time so we can fetch one year data as per NSE website interface
current_datetime = datetime.now() 
OneYearBackDate = current_datetime - timedelta(days=365)

# setup formate for date to feed in url
today = current_datetime.strftime("%d-%m-%Y")
OneYearAgo = OneYearBackDate.strftime("%d-%m-%Y")

def histdata(sym):
    url = sym
    headers = {"Accept-Encoding": "gzip, deflate, br","Accept-Language": "en-US,en;q=0.9","User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    }
    session = requests.Session()
    response = session.get(url, headers=headers).text
    data = json.loads(response)

    records = data['data']['indexCloseOnlineRecords']
    df1 = pd.DataFrame(records)

    # Convert TIMESTAMP column to a more readable format "dd-mm-yyyy" only
    df1['TIMESTAMP'] = pd.to_datetime(df1['TIMESTAMP']).dt.strftime('%d-%m-%Y')
    df = df1[['TIMESTAMP', 'EOD_OPEN_INDEX_VAL', 'EOD_HIGH_INDEX_VAL', 'EOD_CLOSE_INDEX_VAL', 'EOD_LOW_INDEX_VAL']]
    
    return df

try:
    historical_data1 = histdata("https://www.nseindia.com/api/historical/indicesHistory?indexType=NIFTY 50&from="+OneYearAgo+"&to=" + today)
    historical_data2 = histdata("https://www.nseindia.com/api/historical/indicesHistory?indexType=NIFTY BANK&from="+OneYearAgo+"&to=" + today)

    histdata_sheet.range("A3").value = historical_data1.values
    histdata_sheet.range("G3").value = historical_data2.values
except Exception as e:
    print(e)