# Start For loop

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
# End For Loop

# Replace [None] Values with a zero
hqm_dataframe.replace(to_replace=[None], value=0, inplace=True)
