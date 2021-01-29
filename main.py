import datetime as dt
import os
import pandas_datareader.data as web
from openpyxl import load_workbook


filename = 'Investment.xlsx'
csv_name_ticker_prices = 'tickerprices.csv'
path = os.getcwd()
wb_path = os.path.join(os.getcwd(), filename)
time = dt.datetime.combine(dt.date.today(), dt.datetime.min.time())

# read excel file from path
wb = load_workbook(wb_path)

# first sheet of file
sheet_ranges = (wb[wb.sheetnames[0]])

i = 3
first = True
while True:
    ticker = sheet_ranges['B' + str(i)].value
    if ticker is not None:
        print(ticker)
        try:
            if first:
                df = web.DataReader(ticker, 'yahoo', time)
                df['ticker'] = ticker
                first = False
            else:
                df_i = web.DataReader(ticker, 'yahoo', time)
                df_i['ticker'] = ticker
                df = df.append(df_i)
        except:
            continue
        i += 1
    else:
        break
print(df)
df.to_csv(os.path.join(path, csv_name_ticker_prices))




