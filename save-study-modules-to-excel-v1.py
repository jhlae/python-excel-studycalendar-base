import pandas as pd
from datetime import datetime
from openpyxl import Workbook


# Define a function which adds an empty row after each 60th row
def insert_empty_rows(df, interval=60):
    n = len(df)
    # check total number of empty rows needed
    num_empty_rows = n // interval

    # create a new dataframe with empty rows
    new_df = pd.DataFrame()
    for i in range(num_empty_rows):
        new_df = new_df._append(df[i*interval:(i+1)*interval])
        # add empty row
        new_df = new_df._append(pd.Series(), ignore_index=True)
    # after that, continue adding rows
    new_df = new_df._append(df[num_empty_rows*interval:])
    return new_df.reset_index(drop=True)


#
# Start from 08.12.2023 and end on 31.05.2024
# 

# Define a starting date
start_date = datetime(2023, 12, 8)
# Define an ending date
end_date = datetime(2024, 5, 31)
# Define excluded dates
excluded_dates = [
    "6.12.2023", 
    "25.12.2023", 
    "26.12.2023", 
    "27.12.2023", 
    "28.12.2023", 
    "29.12.2023", 
    "1.1.2024",
    "29.3.2024",
    "1.4.2024",
    "1.5.2024",
    "9.5.2024"
    ]

# set pandas date_range
dates = pd.date_range(start_date, end_date, freq='B').strftime('%d.%m.%Y')

# add a list of dates within a specified range, convert that with to_datetime
dates = pd.date_range(start_date, end_date, freq='B').strftime('%d.%m.%Y')
dates = pd.to_datetime(dates, dayfirst=True)
dates = dates[~dates.isin(pd.to_datetime(excluded_dates, dayfirst=True))]

# create a dataframe
df = pd.DataFrame({'PVM': dates})
df['VKONRO'] = df['PVM'].dt.isocalendar().week
df['VIIKONPV'] = df['PVM'].dt.day_name(locale='fi_FI').str.lower()

# add counter for modules that resets after sixty days
df['OPISKELUPVÄ'] = (df.index % 60) + 1

# add counter like cell for study weeks that uses cumulative sum function
# only needs to increment when a week changes
df['OPISKELUVKO'] = df['VKONRO'].diff().ne(0).cumsum() + 1

# add counter for module number plus a few empty cells
df['MOD'] = (df.index // 60) + 1
df['STATUS'] = ''
df['KPL'] = ''

# put all the created dataframe cells in a custom 
df = df[['MOD', 'OPISKELUVKO', 'VKONRO', 'OPISKELUPVÄ', 'VIIKONPV', 'PVM', 'STATUS', 'KPL']]
df['PVM'] = df['PVM'].dt.strftime('%d.%m.%Y')  # Muotoillaan päivämäärät

# add empty rows into dataframe
df_with_empty_rows = insert_empty_rows(df)

# OLD:
# save dataframe to an excel file
# excel_filename = 'studycalendar_2023_2024.xlsx'
# df_with_empty_rows.to_excel(excel_filename, index=False)

# NEW:
# save dataframe to excel file using xlxwriter
with pd.ExcelWriter('studycalendar_2023_2024.xlsx', engine='xlsxwriter') as writer:
    df_with_empty_rows.to_excel(writer, sheet_name='Sheet1', index=False)

    # get xlsxwriter workbook and worksheet objects
    workbook  = writer.book
    worksheet = writer.sheets['Sheet1']

    # format cells
    cell_format = ''
    #cell_format = workbook.add_format({'bg_color': '#FFC7CE'})

    for row_num in range(len(df)):
        worksheet.set_row(row_num + 1, None, cell_format)  # row_num+1 as title row is the first