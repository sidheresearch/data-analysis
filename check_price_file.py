import pandas as pd

df = pd.read_excel('ExcelFile/AVG PRICE 2.xlsx', header=None)
print('Showing all rows to find HSN Code:')
for i in range(min(15, len(df))):
    print(f'Row {i}: {list(df.iloc[i])}')
