 # coding: utf-8
import time
from download_sup_sheet import download_sup_table
from download_sup_sheet import sheet_download
from download_sup_sheet import download_find
from download_sup_sheet import download_export
from automatic_download_everything import full_file_download_sequence
from download_F24data import process_data
from create_F24table import create_table
from high_quality_match import high_quality_match
from alumnus_match import alumnus_match 
from download_clue import process_clue
from bitable_updater import update_all_table, append_records
import os
from  handle_batch11and12 import filter_csv_files 
from  handle_batch11and12 import  merge_csv_files

app_id = "cli_a4e5badd833bd00d"
app_token = "n16oQih9feqSrPRvx4y10eoImrUOLDqd"
superior_sheet_token = "YSuZs9MbahpihTtcb2hcFi5nncf"
target_sheet_token = "Fw3WsIDzlhut9UtrSKFcs1uinNh"

download_sup_table('FT18bZpzLaCiUJs4lOlceZIOn3g','tblRHMq6gt75pKWp','./关联上级.xlsx')
#获取当前文件夹所在路径
time.sleep(.5)
current_folder = os.path.dirname(os.path.abspath(__file__))
for sub_id in ['WU9HZB','z8iap0','2vWoH4']:
    #985+意向学校，QS100（蓝色待定），顶尖大厂sub_id
    ticket = sheet_download('Tq68s4womhTLLJtWGumcyNObnOd','csv','sheet',sub_id)
    print('ticket:',ticket)
    time.sleep(.5)
    file_token = download_find(ticket,'Tq68s4womhTLLJtWGumcyNObnOd')
    print("file_token:",file_token)
    download_export(file_token,current_folder)


base_url = 'https://apply.miracleplus.com'
group_info_url = f"{base_url}/admin/contact_statistics/groups"
login_url = f"{base_url}/users/sign_in"
contacts_download = f"{base_url}/admin/contact_statistics/export"
exportation_url = f"{base_url}/evaluation/applications/"
exportation_contact_url = f"{exportation_url}"
exportation_application_url = f"{exportation_url}/exportation"
history_export_url = f"{base_url}/export_history"

#下载线索
process_clue()

batch_id = 12

csv_name = ["每日数据/F24人脉匹配名单.csv","每日数据/F24自定义导出.csv"]
full_file_download_sequence(
    batch_id, 
    base_url, 
    group_info_url,  
    login_url,
    contacts_download,
    exportation_url, 
    exportation_contact_url,
    exportation_application_url,
    history_export_url,
    csv_name
)
time.sleep(1)




# 调用handle_batch11and12第一个函数进行筛选
filter_csv_files("./人脉数据batch_id11", '2024-01-13', '2024-03-06')

# 调用handle_batch11and12第二个函数进行合并
merge_csv_files("./人脉数据", "./人脉数据20240113-20240306", "./人脉数据总")




process_data(start_date ='2024-01-13') #在这里面把两期的数据合并了

high_quality_match()

create_table()

alumnus_match()


app_token = 'NrqIbSAtoa08cdsHBOgckOwbngb'
relation_dict = {
    'tblQ0kxP02JETOXG':'./输出结果/提交明细.xlsx',
    'tblgJ4y1gh71LWJB':'./输出结果/渠道统计.xlsx',
    'tblGTNrUM55XP0pF':'./输出结果/小组统计.xlsx',
    'tblLFxcCMRHblTuu':'./输出结果/每日个人绩效.xlsx',  
    }
update_all_table(app_token,relation_dict)

app_token = 'BpuwbnQksaPGGvsZEOscXAL8nlc'
relation_dict = {
    'tblFboWZZaq0ZEvn':'./输出结果/重复人脉记录表.xlsx',  
    }
append_records(app_token,relation_dict)

