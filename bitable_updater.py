import json
import pandas as pd
import lark_oapi as lark
from lark_oapi.api.bitable.v1 import *
import math
import pandas as pd
import time

def delete_bitable_records(app_token,table_id):
    while True:
        time.sleep(0.5)
        # 创建client
        client = lark.Client.builder() \
            .app_id("cli_a46fd0e390b8900d") \
            .app_secret("ueiQHthqI8znv0nnvvSzIbcty53Htudp") \
            .log_level(lark.LogLevel.DEBUG) \
            .build()

        # 构造请求对象
        request: ListAppTableRecordRequest = ListAppTableRecordRequest.builder() \
            .app_token(app_token) \
            .table_id(table_id) \
            .page_size(500) \
            .build()

        # 发起请求
        response: ListAppTableRecordResponse = client.bitable.v1.app_table_record.list(request)

        # 处理失败返回
        if not response.success():
            lark.logger.error(
                f"client.bitable.v1.app_table_record.list failed, code: {response.code}, msg: {response.msg}, log_id: {response.get_log_id()}")
            return

        if (lark.JSON.marshal(response.data.items)) == None:
            return '删除完成'
        return_data = json.loads(lark.JSON.marshal(response.data.items))
        id_list = [item['record_id'] for item in return_data]

        # 构造请求对象
        request: BatchDeleteAppTableRecordRequest = BatchDeleteAppTableRecordRequest.builder() \
            .app_token(app_token) \
            .table_id(table_id) \
            .request_body(BatchDeleteAppTableRecordRequestBody.builder()
                .records(id_list)
                .build()) \
            .build()

        # 发起请求
        response: BatchDeleteAppTableRecordResponse = client.bitable.v1.app_table_record.batch_delete(request)

        # 处理失败返回
        if not response.success():
            lark.logger.error(
                f"client.bitable.v1.app_table_record.batch_delete failed, code: {response.code}, msg: {response.msg}, log_id: {response.get_log_id()}")
            return 

def upload_bitable_records(app_token,table_id,upload_file_path,chunk_size = 500):
    df = pd.read_excel(upload_file_path)

    # 检查每一列的列名，如果包含 '时间' 或 '日期'，则转换为时间戳
    for col in df.columns:
        if '时间' in col or '日期' in col:
            df[col] = pd.to_datetime(df[col]).astype('int64') // 10**6

    # 将 DataFrame 转换为字典的列表
    data_list = df.to_dict('records')

    # 清除每个字典中值为 NaN 或 None 的键
    cleaned_dict_list = []
    for record in data_list:
        # 删除值为 None, NaN 或 inf 的键值对
        cleaned_dict = {k: v for k, v in record.items() if v is not None and not (isinstance(v, float) and (math.isnan(v) or math.isinf(v)))}
        cleaned_dict_list.append(cleaned_dict)
    
    def get_chunks(lst, chunk_size):
        # 将列表划分为指定大小的子列表
        for i in range(0, len(lst), chunk_size):
            yield lst[i:i + chunk_size]

    client = lark.Client.builder() \
        .app_id("cli_a46fd0e390b8900d") \
        .app_secret("ueiQHthqI8znv0nnvvSzIbcty53Htudp") \
        .log_level(lark.LogLevel.DEBUG) \
        .build()

    # 从列表中获取每个包含500个记录的块
    for data_chunk in get_chunks(cleaned_dict_list,chunk_size):
        upload_data = [AppTableRecord.builder().fields(item).build() for item in data_chunk]

        # 构造请求对象
        request: BatchCreateAppTableRecordRequest = BatchCreateAppTableRecordRequest.builder() \
            .app_token(app_token) \
            .table_id(table_id) \
            .request_body(BatchCreateAppTableRecordRequestBody.builder()
                .records(upload_data)
                .build()) \
            .build()

        # 发起请求
        response: BatchCreateAppTableRecordResponse = client.bitable.v1.app_table_record.batch_create(request)

        # 处理失败返回
        if not response.success():
            lark.logger.error(
                f"client.bitable.v1.app_table_record.batch_create failed, code: {response.code}, msg: {response.msg}, log_id: {response.get_log_id()}")
            return

        # 处理业务结果
        lark.logger.info(lark.JSON.marshal(response.data, indent=4))

        time.sleep(0.5)

def update_all_table(app_token,relation_dict):
    for table_id,file_path in relation_dict.items():
        delete_bitable_records(app_token,table_id)
        upload_bitable_records(app_token,table_id,file_path)

def append_records(app_token,relation_dict):
    for table_id,file_path in relation_dict.items():
        upload_bitable_records(app_token,table_id,file_path)