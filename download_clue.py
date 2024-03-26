import requests
import json
import os
import io
import sys
from pathlib import Path
import time
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
import pandas as pd
import numpy as np

def get_login_cookies(base_url, login_url, session):
    login_page = session.get(login_url)
    soup = BeautifulSoup(login_page.text, 'html.parser')
    token = soup.select_one('meta[name=csrf-token]').get('content')
    payload = {
      'login_ways': 'email',
      'user[source]': '',
      'user[login]': 'miracleplusbrain@gmail.com',
      'user[password]': 'MPB87654321',
      'commit': 'Login',
      'authenticity_token': token
    }

    headers = { 
      'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36', 
      'Content-Type': 'application/x-www-form-urlencoded',
      'Origin': base_url,
      'Referer': login_url,
    }
    response = session.post(login_url, headers=headers, data=payload)
    # to save cookies:  # get some cookies
    cookies = requests.utils.dict_from_cookiejar(session.cookies)  # turn cookiejar into dict
    return cookies

def download_all_contacts(contacts_url, cookies, batch_id):

    contacts = []

    # The contacts URL you want to download from
    url = contacts_url
    # The parameters for the GET request
    params = {
        'batch_id': batch_id
    }
    cookies = cookies
    # The headers for the request
    headers = {
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
        'Accept-Encoding': 'gzip, deflate, br',
        'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
        'Referer': 'https://apply.miracleplus.com/admin/contacts?filters=0&page=1',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.0.0 Safari/537.36',
        'Sec-Fetch-Dest': 'document',
        'Sec-Fetch-Mode': 'navigate',
        'Sec-Fetch-Site': 'same-origin',
        'Sec-Fetch-User': '?1',
        'Upgrade-Insecure-Requests': '1',
    }
    # Send the GET request
    response = requests.get(url, headers=headers, params=params, cookies=cookies)
    # Check if the request was successful
    if response.status_code == 200:
        # If successful, write the content to a file (replace 'file_path' with the path where you want to save the file)
        data_dic = json.loads(response.content)
        total = data_dic['total']
        for page in range(1,(total//50)+2):
            page = download_page_contacts(contacts_url,cookies,batch_id,page)
            for line in page:
                contacts.append(line)
        return contacts
    else:
        print(f"encountered error when visiting Page1 !")
        return

def download_page_contacts(contacts_url, cookies, batch_id, page):
    # The contacts URL you want to download from
    url = contacts_url + '&page=' + str(page)
    # The parameters for the GET request
    params = {
        'batch_id': batch_id
    }
    cookies = cookies
    # The headers for the request
    headers = {
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
        'Accept-Encoding': 'gzip, deflate, br',
        'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
        'Referer': 'https://apply.miracleplus.com/admin/contacts?filters=0&page=1',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/103.0.0.0 Safari/537.36',
        'Sec-Fetch-Dest': 'document',
        'Sec-Fetch-Mode': 'navigate',
        'Sec-Fetch-Site': 'same-origin',
        'Sec-Fetch-User': '?1',
        'Upgrade-Insecure-Requests': '1',
    }
    # Send the GET request
    response = requests.get(url, headers=headers, params=params, cookies=cookies)
    # Check if the request was successful
    if response.status_code == 200:
        # If successful, write the content to a file (replace 'file_path' with the path where you want to save the file)
        data_dic = json.loads(response.content)
        return data_dic['data']
        '''print(line)
        with open('./test.csv', 'wb') as file:
            file.write(response.content)
        print('downloaded group')'''
    else:
        print(f"encountered error when downloading id !")
        return

def download():
    print('start downloading contacts...')
    # 获取今天的日期
    today = datetime.now()
    # 计算明天的日期
    enddate = today + timedelta(days=1)
    # 计算昨天的日期
    startdate= today - timedelta(days=7)
    # 将日期格式化为字符串
    date_format = "%Y-%m-%d"
    tomorrow_str = enddate.strftime(date_format)
    yesterday_str = startdate.strftime(date_format)

    batch_id = 11
    full_download_sessions = requests.Session() # start session
    base_url = 'https://apply.miracleplus.com'
    login_url = f"{base_url}/users/sign_in"
    contacts_url = f"{base_url}/admin/contacts?format=json&filters=0"

    yesterday_contacts_url = 'https://apply.miracleplus.com/admin/contacts?format=json&filters=1&created_at[]='+yesterday_str+'&created_at[]='+tomorrow_str
    cookies = get_login_cookies(base_url, login_url, full_download_sessions)
    list_of_dicts = download_all_contacts(yesterday_contacts_url, cookies, batch_id)

    return list_of_dicts

def process_clue(start_date='2024-01-13'):
    list_of_dicts = download()
    print('processing contacts...')
    # 将列表转换为DataFrame
    newclue = pd.DataFrame(list_of_dicts)
    newclue= newclue[['id','name','phone', 'email', 'category', 'created_at', 'clue_user_name','source', 'source_other']]
    # 将该列转换为datetime类型
    newclue['created_at'] = pd.to_datetime(newclue['created_at'])

    # 格式化日期时间列为"YYYY/MM/D"格式
    newclue['Formatted_Date'] = newclue['created_at'].dt.strftime('%Y/%m/%d')

    # 删除原始的日期时间列
    newclue['created_at'] = newclue['Formatted_Date']
    newclue = newclue[newclue['created_at'] >= start_date]
    newclue = newclue.drop('Formatted_Date',axis=1)
    # 新列名列表，顺序和原始列名的顺序一致
    new_column_names = ['线索id', '线索姓名', '电话号码','电子邮箱','画像','提交时间','姓名','来源渠道','来源渠道备注']
    # 使用rename方法修改列名
    newclue.columns = new_column_names
    # 新增列并设置默认值为None
    newclue['来源渠道备注'] = None
    newclue['关联上级'] = None
    newclue['所属渠道'] = None
    # 将所有None替换为NaN
    newclue.replace({None: np.nan}, inplace=True)
    oldclue = pd.read_excel("./下载数据/线索明细.xlsx")
    # 使用concat纵向拼接
    clues = pd.concat([newclue, oldclue], axis=0)
    # 去除重复行
    clues = clues.drop_duplicates(subset='线索id')
    # 将电话号码列的数据类型转换为字符串
    #clues['电话号码'] = clues['电话号码'].astype(str)
    clues.to_excel("./下载数据/线索明细.xlsx",index=False)