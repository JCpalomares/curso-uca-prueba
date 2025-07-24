import pandas as pd
#import openpyxl

df = pd.read_excel('datos/top10s.xlsx', sheet_name='TOP10S')

print (df)

#excel_dataframe = openpyxl.load_workbook('datos/top10s.xlsx')
#dataframe = excel_dataframe.active

# print(dataframe)

#data = []

#for row in range(1, dataframe.max_row):
    #print(row)
#    _row = dataframe.cell(row, 1).value
#    print(_row)
    #for col in dataframe.iter_cols(0, dataframe.max_column):
        # print(col[row].value)
        #_row.append(col[row].value)
    #data.append(_row)
