from openpyxl import Workbook
from openpyxl.drawing.image import Image

def ins_img(image):
    wb = Workbook()
    ws = wb.active
    img = Image(image)
    ws.add_image(img, 'A5')
    wb.save('output.xlsx')

import pandas as pd
import json
import ast
def read_input_excel():
    original_excel=pd.read_excel('tableau_input.xlsx')
    report_name=original_excel['Report_name'][0]
    sheet_name=original_excel['sheet_name'][0]
    filter_name=original_excel['filter_name'][0]
    field_name=original_excel['field_name'][0]
    return report_name,sheet_name,filter_name,field_name

def export():
    excel=pd.read_excel('all_sheet_names.xlsx')
    report_name=excel['workbook']
    sheet_name_all=excel['name']
    og_id=excel['id']
    
    len_of_reports=len(report_name)
    res1,res2,res3,res4=read_input_excel()

    for x in range(len_of_reports):
        report_name_from_excel=res1
        data_in_dict = report_name.loc[x]
        og_idx = og_id.loc[x]
        # print(type(og_id))
        view_id_ip=res2
        view_name=sheet_name_all.loc[x]
        dict_from_ast = ast.literal_eval(data_in_dict)

        if dict_from_ast['name']==report_name_from_excel and view_id_ip==view_name:
            
            return og_idx
# print(export())




