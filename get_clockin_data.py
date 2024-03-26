import lark_oapi as lark
from lark_oapi.api.contact.v3 import *
from lark_oapi.api.attendance.v1 import *
import pandas as pd
import json
from datetime import datetime, timedelta

# SDK 使用说明: https://github.com/larksuite/oapi-sdk-python#readme
# 复制该 Demo 后, 需要将 "APP_ID", "APP_SECRET" 替换为自己应用的 APP_ID, APP_SECRET.
def get_users_of_department(department_id):
    # 创建client
    client = lark.Client.builder() \
        .app_id("cli_a46fd0e390b8900d") \
        .app_secret("ueiQHthqI8znv0nnvvSzIbcty53Htudp") \
        .log_level(lark.LogLevel.DEBUG) \
        .build()

    # 构造请求对象  
    #user_id_type 可选：open_id / union_id / user_id
    #department_id_type 可选：department_id / open_department_id
    request: FindByDepartmentUserRequest = FindByDepartmentUserRequest.builder() \
        .user_id_type("user_id") \
        .department_id_type("open_department_id") \
        .department_id(department_id) \
        .page_size(40) \
        .build()

    # 发起请求
    response: FindByDepartmentUserResponse = client.contact.v3.user.find_by_department(request)

    # 处理失败返回
    if not response.success():
        lark.logger.error(
            f"client.contact.v3.user.find_by_department failed, code: {response.code}, msg: {response.msg}, log_id: {response.get_log_id()}")
        return

    # 处理业务结果
    user_ids = []
    return_data = json.loads(lark.JSON.marshal(response.data.items))
    for item in return_data:
        user_ids.append(item['user_id'])

    return user_ids

# SDK 使用说明: https://github.com/larksuite/oapi-sdk-python#readme
# 复制该 Demo 后, 需要将 "APP_ID", "APP_SECRET" 替换为自己应用的 APP_ID, APP_SECRET.
# 计算每组打卡人数，并输出打卡名单
def count_department_workers_today(users_of_department,startday_int,yesterday_int):
    # 创建client
    client = lark.Client.builder() \
        .app_id("cli_a46fd0e390b8900d") \
        .app_secret("ueiQHthqI8znv0nnvvSzIbcty53Htudp") \
        .log_level(lark.LogLevel.DEBUG) \
        .build()

    # 构造请求对象
    request: QueryUserTaskRequest = QueryUserTaskRequest.builder() \
        .employee_type("employee_id") \
        .request_body(QueryUserTaskRequestBody.builder()\
            .user_ids(users_of_department)
            .check_date_from(startday_int)
            .check_date_to(yesterday_int)
            .need_overtime_result(True)
            .build()) \
        .build()

    # 发起请求
    response: QueryUserTaskResponse = client.attendance.v1.user_task.query(request)

    # 处理失败返回
    if not response.success():
        lark.logger.error(
            f"client.attendance.v1.user_task.query failed, code: {response.code}, msg: {response.msg}, log_id: {response.get_log_id()}")
        return

    # 处理业务结果
    return_data = json.loads(lark.JSON.marshal(response.data.user_task_results))
    count = 0
    worker_num = 0 #总打卡人数
    onsite_worker = 0 #线下打卡人数
    outsite_worker = 0 #线上打卡人数
    all_name_list = [] #所有打卡用户名
    onsite_name_list = [] #线下办公打卡用户
    outsite_name_list = [] #线上办公打卡用户
    print(return_data)
    for item in return_data:
        flag = False #设置判断今天是否打卡flag
        flag1 = False #设置判断打卡地点是否为办公区--智源大厦  如果打卡地点为智源大厦则为TRUE 否则则为FALSE
        records = item['records']
        for day_record in records:
#            print(day_record)
            if day_record['check_in_record_id'] != '':
                count += 0.5
                flag = True
            if day_record['check_out_record_id'] != '':
                count += 0.5
                flag = True
            if(day_record.get('check_in_record') is not None or day_record.get('check_out_record') is not None): #上午或者下午有打卡明细
#              print(day_record.get('check_in_record')['location_name'])
              if(day_record.get('check_in_record') is not None and day_record['check_in_record']['location_name'] !=''):
                  flag1 = True
              if(day_record.get('check_out_record') is not None and day_record['check_out_record']['location_name'] !=''):
                  flag1 = True
#              if(day_record['check_in_record']['location_name'] =='' and day_record['check_out_record']['location_name'] ==''):
#                  outsite_worker +=1 
#                  flag1 = False
        if flag:
            all_name_list.append(item['employee_name'])
        if flag1:
            onsite_name_list.append(item['employee_name'])
        if (flag and (flag1 == False)):
            outsite_name_list.append(item['employee_name'])
        onsite_worker = len(onsite_name_list)
        outsite_worker = len(outsite_name_list)
        worker_num = len(all_name_list)
    return {'count': count, '总打卡人数': worker_num, 'onsite_worker': onsite_worker, 'outsite_worker': outsite_worker, '总打卡名单':all_name_list, '线上打卡名单': outsite_name_list, '线下打卡名单': onsite_name_list}
    return(return_data)
def unique_names(row):
    if not row:
        return []
    else:
        names = row.split(',')
        unique = set(names)
        return unique

def get_clockin_data(start_date,end_date):
    # 将昨天的日期转换为int格式
    endday_int = int(end_date.strftime('%Y%m%d'))
    startday_int = int(start_date.strftime('%Y%m%d'))

    department_ids = {
        '周京华':'od-49018cf4912e5c7ccc8786f52bc8ef35',#周京华 CCC
        '潘艳青':'od-bd0eb22fe85aaf57bd213412fc622ad1', #潘艳青CE3C
        '郑子宜':'od-0376416c1735b2a2652df33df5345025',
        '魏郅杰':'od-5ad0cffc0383713d905da0be3be8d7c0',
        '王鑫':'od-8a10fcb7ccf90104bea612b3f2154e0b',#Watson组
        '江哲昊':'od-badcfb95ff9a10d697ff28421d2afe98', #Jocelyn组
        '张正端':'od-961ab5447020fa4397f9cc4ade105e07',
        '王泰然':'od-951f8e196e6d4684d3f52e6e3c99e44b',
        '张茜':'od-bbdf3852df2becc113bb57d359da9347',
        #'HR':'od-9ad4908fb909bfa475c60d537b0a831c',
        #'Fellow':'od-20fde6c1257efdfec6fec3009c6719d6',
        #'Campus scout':'od-79030b75debf77e0f10717e337a3d59f',
        #'H-Fellow':'od-d3da718f7b10fde14de48ba3e0c38c26',
        #'海外global组':'od-61871c61acec8b7c4f8bdf4614d01be5',
        #'曾博文活动组':'od-ccf80791005951d31c005957e0671fe0',
        '郭梦瑶':'od-a4979f65bcbcd879565bcfa5dca9468a',#郭梦瑶Asher组
        '谭炜川':'od-46b7c22964b808875a55791aef0e1530',
        '刘奕龙':'od-f0567b4720f90fd493e033919ccf5746',        
        '陈奕璇':'od-0746e04a02b3e9e554a63e3041b38a43',
        '王勉':'od-0a187632a367adfa57e0cb2808264348',        
        '王三川':'od-4c1f6e2ec5cc27b0280359f865dc2173',        
        #'Finance':'od-15bb464b248d8e630084495b5e611fb2',
        #'运营与后勤':'od-432418538b2a26b6805dbffb67a04de5',
        #'彭书航':'od-d807610d4c81d645eccafe49936e828a',
        #'行研部门':'od-690b7ccfe4c6b2197ec488f026a6660c',
        #'产品部门':'od-db44b2721eeb2e50d442a3ec384c3bc6',
        #'校友社区运营':'od-6d89d4cbe6605f1d6c5924b37f37d8a2',
        #'Wenwen组':'od-7507098be96eb31d624e991bc5b6dd76'
        #'陈聪组':'od-d3717702c1ba68f8490a2b43a4c55975'
        }

    worker_count_result = {}

    for k, v in department_ids.items():
        print(k)
        worker_count_result[k] = count_department_workers_today(get_users_of_department(v), startday_int,endday_int)
    worker_count_result = pd.DataFrame.from_dict(worker_count_result, orient='index').reset_index().rename(columns={'index': '小组长','count': '打卡人次',
                                                                                                                    'onsite_worker': '线下办公人数',
                                                                                                                    'outsite_worker': '线上办公人数'})
    worker_count_result['总名单'] = worker_count_result['总打卡名单'].apply(lambda x: ','.join(x))

    #worker_count_result['线上打卡名单'] = worker_count_result['线上打卡名单'].apply(lambda x: ','.join(x))
    #worker_count_result['线下打卡名单'] = worker_count_result['线下打卡名单'].apply(lambda x: ','.join(x))
    #工作人数是打卡名单去重后统计的该时间段内打过卡的人数
    worker_count_result['总名单'] = worker_count_result['总名单'].apply(unique_names)
    worker_count_result['工作人数'] = worker_count_result['总名单'].apply(len)
    return worker_count_result