import pandas as pd
import json
import ast

excel=pd.read_excel('all_sheet_names.xlsx')
report_name=excel['workbook']
len_of_reports=len(report_name)


for x in range(2):
    # Parse the JSON string
    data = report_name.loc[x]
    dict_from_ast = ast.literal_eval(data)

    if dict_from_ast['name']==report_name_input
    

    # print(data)



