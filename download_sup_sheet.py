import lark_oapi as lark
from lark_oapi.api.bitable.v1 import *
import json
import pandas as pd
import os

def download_sup_table(app_token,table_id,file_path):
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

    # 处理业务结果
    return_data = (json.loads(lark.JSON.marshal(response.data, indent=4)))['items']  #indent是生成JSON字符串时的缩进量

    return_data_list = [item['fields'] for item in return_data]
    df = pd.DataFrame(return_data_list)

    for column in df.columns:
        # 检查列名是否包含'时间'或'日期'
        if '时间' in column or '日期' in column:
            # 转换该列为datetime格式
            df[column] = pd.to_datetime(df[column], unit='ms').dt.date

    df.to_excel(file_path)


from lark_oapi.api.drive.v1 import *
# SDK 使用说明: https://github.com/larksuite/oapi-sdk-python#readme
# token是想要导出文档 token，file_extension导出文件扩展名,type1导出文档类型,sub_id导出子表ID，仅当将电子表格导出为 csv 时使用
# sheet_download得到一个ticket，该函数为创建导出任务
def sheet_download(token,file_extension,type1,sub_id):
    # 创建client
    # 使用 user_access_token 需开启 token 配置, 并在 request_option 中配置 token
    client = lark.Client.builder() \
        .app_id("cli_a46fd0e390b8900d") \
        .app_secret("ueiQHthqI8znv0nnvvSzIbcty53Htudp") \
        .log_level(lark.LogLevel.DEBUG) \
        .build()

    # 构造请求对象
    request: CreateExportTaskRequest = CreateExportTaskRequest.builder() \
        .request_body(ExportTask.builder()
            .file_extension(file_extension)
            .token(token)
            .type(type1)
            .sub_id(sub_id)
            .build()) \
        .build()

    # 发起请求
    option = lark.RequestOption.builder().user_access_token("u-evKNYADYx3xXgXPqleHxlg0504blkhh9h8005l2w06hr").build()
    response: CreateExportTaskResponse = client.drive.v1.export_task.create(request, option)

    # 处理失败返回
    if not response.success():
        lark.logger.error(
            f"client.drive.v1.export_task.create failed, code: {response.code}, msg: {response.msg}, log_id: {response.get_log_id()}")
        return

    # 处理业务结果
    ticket = json.loads(lark.JSON.marshal(response.data, indent=4))['ticket']
    return(ticket)


# SDK 使用说明: https://github.com/larksuite/oapi-sdk-python#readme
# 复制该 Demo 后, 需要将 "YOUR_APP_ID", "YOUR_APP_SECRET" 替换为自己应用的 APP_ID, APP_SECRET.
#download_find查看导出任务结果
def download_find(ticket,token):
    # 创建client
    client = lark.Client.builder() \
        .app_id("cli_a46fd0e390b8900d") \
        .app_secret("ueiQHthqI8znv0nnvvSzIbcty53Htudp") \
        .log_level(lark.LogLevel.DEBUG) \
        .build()

    # 构造请求对象
    request: GetExportTaskRequest = GetExportTaskRequest.builder() \
        .ticket(ticket) \
        .token(token) \
        .build()

    # 发起请求
    response: GetExportTaskResponse = client.drive.v1.export_task.get(request)

    # 处理失败返回
    if not response.success():
        lark.logger.error(
            f"client.drive.v1.export_task.get failed, code: {response.code}, msg: {response.msg}, log_id: {response.get_log_id()}")
        return

    # 处理业务结果
    file_token=json.loads(lark.JSON.marshal(response.data, indent=4))["result"]["file_token"]
    return(file_token)

# SDK 使用说明: https://github.com/larksuite/oapi-sdk-python#readme
# 复制该 Demo 后, 需要将 "YOUR_APP_ID", "YOUR_APP_SECRET" 替换为自己应用的 APP_ID, APP_SECRET.
def download_export(file_token,current_folder):
    # 创建client
    client = lark.Client.builder() \
        .app_id("cli_a46fd0e390b8900d") \
        .app_secret("ueiQHthqI8znv0nnvvSzIbcty53Htudp") \
        .log_level(lark.LogLevel.DEBUG) \
        .build()

    # 构造请求对象
    request: DownloadExportTaskRequest = DownloadExportTaskRequest.builder() \
        .file_token(file_token) \
        .build()

    # 发起请求
    response: DownloadExportTaskResponse = client.drive.v1.export_task.download(request)

    # 处理失败返回
    if not response.success():
        lark.logger.error(
            f"client.drive.v1.export_task.download failed, code: {response.code}, msg: {response.msg}, log_id: {response.get_log_id()}")
        return

    file_path = os.path.join(current_folder, response.file_name) # 处理业务结果
    # 处理业务结果
    with open(file_path, "wb") as f:
        f.write(response.file.read())  
        f.close()
#current_folder = os.path.dirname(os.path.abspath(__file__))
#file_token=   "QSR9bSYrUor8L4xhwHscgUHlnwb"
#download_export(file_token, current_folder)