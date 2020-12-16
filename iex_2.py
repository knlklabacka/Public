# Momentum Strategy
import numpy as np
import pandas as pd
import requests
import math
import xlsxwriter
from scipy.stats import percentileofscore as score
from secrets import IEX_CLOUD_API_TOKEN
from statistics import mean


def chunks(lst, n):
    for i in range(0, len(lst), n):
        yield lst[i:i + n]


def portfolio_input():
    # TODO get user input
    #  portfolio_size = input('Enter the size of your portfolio:')
    #  I'm hard coding the portfolio size for testing purposes
    return 10000


stocks = pd.read_csv('sp_500_stocks.csv')
iex_base_url = 'https://cloud.iexapis.com'
sandbox_base_url = 'https://sandbox.iexapis.com/stable'
symbol_groups = list(chunks(stocks['Ticker'], 100))
symbol_strings = []
my_columns = ['Tickers', 'Stock Price', 'One Year Price Return', 'Number of Shares to Buy']
hqm_columns = [
                'Ticker',
                'Price',
                'Number of Shares to Buy',
                'One-Year Price Return',
                'One-Year Return Percentile',
                'Six-Month Price Return',
                'Six-Month Return Percentile',
                'Three-Month Price Return',
                'Three-Month Return Percentile',
                'One-Month Price Return',
                'One-Month Return Percentile',
                'HQM Score'
                ]
time_periods = ['One-Year', 'Six-Month', 'Three-Month', 'One-Month']

for i in range(0, len(symbol_groups)):
    symbol_strings.append(','.join(symbol_groups[i]))

final_dataframe = pd.DataFrame(columns=my_columns)
hqm_dataframe = pd.DataFrame(columns=hqm_columns)
pd.set_option("display.max_rows", None, "display.max_columns", None)
hqm_dataframe.replace(to_replace=None, value=0, inplace=True)

# Start for For Loop
for symbol_string in symbol_strings:
    batch_api_call_url = f'{sandbox_base_url}/stock/market/batch/?types=stats,quote&symbols={symbol_string}&token={IEX_CLOUD_API_TOKEN}'
    data = requests.get(batch_api_call_url).json()
    for symbol in symbol_string.split(','):
        hqm_dataframe = hqm_dataframe.append(
            pd.Series([symbol,
                       data[symbol]['quote']['latestPrice'],
                       'N/A',
                       data[symbol]['stats']['year1ChangePercent'],
                       'N/A',
                       data[symbol]['stats']['month6ChangePercent'],
                       'N/A',
                       data[symbol]['stats']['month3ChangePercent'],
                       'N/A',
                       data[symbol]['stats']['month1ChangePercent'],
                       'N/A',
                       'N/A'
                       ],
                      index=hqm_columns),
            ignore_index=True)
# End of For Loop
        
# Replace [None] values in DataFrame with zero
hqm_dataframe.replace(to_replace=[None], value=0, inplace=True)

for row in hqm_dataframe.index:
    for time_period in time_periods:
        change_col = f'{time_period} Price Return'
        percentile_col = f'{time_period} Return Percentile'
        hqm_dataframe.loc[row, percentile_col] = score(hqm_dataframe[change_col],
                                                           hqm_dataframe.loc[row, change_col]) / 100


for row in hqm_dataframe.index:
    momentum_percentiles = []
    for time_period in time_periods:
        momentum_percentiles.append(hqm_dataframe.loc[row, f'{time_period} Return Percentile'])

    hqm_dataframe.loc[row, 'HQM Score'] = mean(momentum_percentiles)


hqm_dataframe.sort_values('HQM Score', ascending=False, inplace=True)
hqm_dataframe = hqm_dataframe[:50]
hqm_dataframe.reset_index(inplace=True, drop=True)
position_size = float(portfolio_input()/len(hqm_dataframe.index))

for i in hqm_dataframe.index:
    hqm_dataframe.loc[i, 'Number of Shares to Buy'] = math.floor(position_size/hqm_dataframe.loc[i, 'Price'])

writer = pd.ExcelWriter('recommended_trades.xlsx', engine='xlsxwriter')
hqm_dataframe.to_excel(writer, 'Momentum Strategy', index=False)

background_color = '#0a0a23'
font_color = '#ffffff'

string_format = writer.book.add_format(
    {
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1
    }
)
dollar_format = writer.book.add_format(
    {
        'num_format':'$0.00',
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1
    }
)
integer_format = writer.book.add_format(
    {
        'num_format':'0',
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1
    }
)
percent_format = writer.book.add_format(
    {
        'num_format':'0.0%',
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1
    }
)

column_formats = {
    'A': ['Ticker', string_format],
    'B': ['Price', dollar_format],
    'C': ['Number of Shares to Buy', integer_format],
    'D': ['One-Year Price Return', percent_format],
    'E': ['One-Year Return Percentile', percent_format],
    'F': ['Six-Month Price Return', percent_format],
    'G': ['Six-Month Return Percentile', percent_format],
    'H': ['Three-Month Price Return', percent_format],
    'I': ['Three-Month Return Percentile', percent_format],
    'J': ['One-Month Price Return', percent_format],
    'K': ['One-Month Return Percentile', percent_format],
    'L': ['HQM Score', integer_format]
    }

for column in column_formats.keys():
    writer.sheets['Momentum Strategy'].set_column(f'{column}:{column}', 25, column_formats[column][1])
    writer.sheets['Momentum Strategy'].write(f'{column}1', column_formats[column][0], string_format)

writer.save()



