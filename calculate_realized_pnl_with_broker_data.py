import pandas as pd
import numpy as np
import sys


pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', 1200)


########## input ##########
file_name = '20230104_20231213_3888836_5274511_one_unit_per_row.xlsx'
tool = 'CN'
tool_list = ['HSI', 'HHI', 'CN']
multiplier_dict = {'HSI': 50, 'HHI': 50, 'CN': 1}

########## input ##########

########## read file ##########
unit_row_df = pd.read_excel(file_name)
unit_row_df['date'] = pd.to_datetime(unit_row_df['date'])

########## read file ##########

########## get realized pnl by tool ##########
for tool in tool_list:
    print(f'#########################################################################################################\n'
          f'Below is section for {tool}\n')

    ##### find the last zero-position point and unit rows with the same trade_id #####
    tool_specific_df = unit_row_df.loc[unit_row_df['tool'] == tool]
    tool_specific_df = tool_specific_df.reset_index()  # retain the original index column
    tool_specific_df['unit_change'] = np.where(tool_specific_df['direction'] == 'B', 1, -1)
    tool_specific_df['cumulative_position'] = tool_specific_df['unit_change'].cumsum()
    zero_position_points = tool_specific_df[tool_specific_df['cumulative_position'] == 0]

    last_idx = zero_position_points.index[-1]  # index of tool_specific df
    original_idx = zero_position_points.at[last_idx, 'index']  # index of unit_row_df
    date = zero_position_points.at[last_idx, 'date']
    trade_id = zero_position_points.at[last_idx, 'trade_id']

    print(f'date: {date}')
    print(f'trade id: {trade_id}')
    print(f'index in the original unit_row_df: {original_idx}')  # helps identify which unit to be exact within the trade id
    print(f'all unit rows of the same trade_id:\n'
          f'{tool_specific_df.loc[tool_specific_df["trade_id"] == trade_id]}\n')

    ##### find the last zero-position point and unit rows with the same trade_id #####

    ##### calculate realized pnl up to the last zero-position cutoff point #####
    unit_buy_before_cutoff = tool_specific_df.loc[(tool_specific_df['direction'] == 'B') &
                                                  (tool_specific_df.index <= last_idx)]
    unit_sell_before_cutoff = tool_specific_df.loc[(tool_specific_df['direction'] == 'S') &
                                                   (tool_specific_df.index <= last_idx)]
    num_buy_units = unit_buy_before_cutoff['unit'].sum()
    num_sell_units = unit_sell_before_cutoff['unit'].sum()

    if num_buy_units == num_sell_units:
        avg_buy_price = unit_buy_before_cutoff['price'].mean()
        avg_sell_price = unit_sell_before_cutoff['price'].mean()
        realized_pnl = (avg_sell_price - avg_buy_price) * num_buy_units * multiplier_dict[tool]
        commission = unit_buy_before_cutoff['commission'].sum() + unit_sell_before_cutoff['commission'].sum()
        unit_commission = commission / (num_buy_units + num_sell_units)

        print(f'weighted average buy price: {avg_buy_price}')
        print(f'weighted average sell price: {avg_sell_price}')
        print(f'number of units realized: {num_buy_units}')
        print(f'realized pnl as of {date}: {realized_pnl}')
        print(f'total commission paid: {commission}')
        print(f'unit commission rate: {unit_commission}')

    else:
        print('Number of buys is not equal to number of sells')

    print(f'End of section for {tool}\n'
          f'#########################################################################################################\n'
          )
    ##### calculate realized pnl up to the last zero-position point #####

########## get realized pnl by tool ##########

