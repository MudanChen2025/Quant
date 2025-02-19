import pandas as pd
import sys


pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', 1200)


########## read trimmed raw excel data ##########
df = pd.read_excel('raw_data_for_python.xlsx')
df['交易日期'] = pd.to_datetime(df['交易日期'])

########## read trimmed raw excel data ##########

########## set output file name ##########
last_idx = df.index[-1]
start_date = df.at[0, '交易日期'].strftime('%Y%m%d')
end_date = df.at[last_idx, '交易日期'].strftime('%Y%m%d')
start_trade_id = df.at[0, '流水号']
end_trade_id = df.at[last_idx, '流水号']
output_file_name = f'{start_date}_{end_date}_{start_trade_id}_{end_trade_id}_one_unit_per_row.xlsx'
print(df)
print(df.info())

########## set output file name ##########

########## breakdown trades into one unit per row ##########
unit_row_df = pd.DataFrame({
    'trade_id': [],
    'date': [],
    'tool': [],
    'direction': [],
    'price': [],
    'unit': [],
    'commission': []
})

for i in range(len(df)):
    trade_id = df.at[i, '流水号']
    date = df.at[i, '交易日期']
    tool = df.at[i, '品种编号']
    direction = df.at[i, '买卖方向']
    price = df.at[i, '成交价']
    unit = df.at[i, '成交量']
    commission = df.at[i, '客户手续费']
    unit_commission = commission / unit

    for j in range(unit):
        unit_row = pd.DataFrame({
            'trade_id': [trade_id],
            'date': [date],
            'tool': [tool],
            'direction': [direction],
            'price': [price],
            'unit': [1],
            'commission': [unit_commission]
        })
        unit_row_df = pd.concat([unit_row_df, unit_row], ignore_index=True)

print(unit_row_df)
print(unit_row_df.info())
sys.exit()
# validation
num_rows = len(unit_row_df)
num_units = df['成交量'].sum()
if num_rows == num_units:
    print(unit_row_df)
    unit_row_df.to_excel(output_file_name, index=False)
    print(f'unit_row_df has been saved as {output_file_name}.')
else:
    print('Number of rows doesn\'t match number of units traded. Please double check.')
