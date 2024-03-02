from tableau_api_lib import TableauServerConnection as conn
from tableau_api_lib.utils import querying,flatten_dict_column
from tableau_api_lib.utils import extract_pages
from tableau_api_lib.utils.querying import get_views_dataframe, get_view_data_dataframe
import pandas as pd
import excel_ops 

# def read_excel():
#     pd.read_excel('')

# def get_view_id():
#     view_id=pd.read_excel('')
#     return view_id
def tableau():
    config={
        'tableau_prod':{
            'server':'https://analyticsreportsauth.boltinc.com/',
            'api_version':'3.19',
            'username':'shivanandk@boltinc.com',
            'password':'Shivu@123',
            'personal_access_token':'Qa_test',
            'personal_access_token_secret':'EJrMRTEVTUeNQvfze3or5g==:B0siqlGpQ4udGnBpsQHfFKeEAqjrjDK5',
            'site_name':'boltInternalReports',
            'site_url':'boltInternalReports'
            # 'site_id':'b81033dd-6d79-49e0-809a-bed4070b39ca',
        }

    }
    connection=conn(config)
    response=connection.sign_in()
    # response2=connection.server()
    # print(connection.default_headers)

    # all_sites = extract_pages(connection.query_sites)
    # site_list = [(site['name'], site['id']) for site in all_sites]
    # print(site_list)


    #exported excel on feb 2024
    # file_name=r'C:\Users\shivanandk\Desktop\Practise_SK\all_sheet_names.xlsx'
    # views_df.to_excel(file_name, sheet_name='sheet_name', index=False)
    # print('exported...')
    
    def blank_to_char(text:str)->str:
        text=text.replace('','%20')
        return text

    def coded_params(param:dict)->dict:
        encoded_dict={}
        for key in param.dict.keys():
            encoded_dict[key]=blank_to_char(str(param[key]))
        return encoded_dict
    
    report_name,sheet_name,filter_name,field_name=excel_ops.read_input_excel()
    # field_name='Partner Name'
    # filter_name='NGIC'
    field_name2='State'
    filter_name2='AR'
    v_id_from_input=excel_ops.export()
    params={"filter_!":f"vf_{field_name}={filter_name}&vf_{field_name}={filter_name2}"}
    df3=querying.get_view_data_dataframe(connection ,view_id=v_id_from_input,parameter_dict=params)
    print(df3)

    img=connection.query_view_image(view_id=v_id_from_input,parameter_dict=params)

    with open('tableau_img2.png','wb') as file:
        file.write(img.content)
    excel_ops.ins_img('tableau_img2.png')
    print('completed..')

tableau()



