import pandas as pd
from datetime import date, timedelta, datetime
import pandas as pd
import os
import xlsxwriter
def process_data(start_date ='2024-01-13'):
    end_date = datetime.today().strftime('%Y-%m-%d')

    # 定义筛选和统计函数
    def process_csv(start_date, end_date, file_path):
        df = pd.read_csv(file_path)
        df['日期'] = pd.to_datetime(df['\\uFEFF日期'])
        filtered_df = df[(df['日期'] >= start_date) & (df['日期'] <= end_date)]
        result = filtered_df.groupby(['姓名', '日期']).agg({'人脉增量': 'sum', '自增总量': 'sum', '开表增量': 'sum', '提交增量': 'sum'}).reset_index()
        return result

    # 读取并处理文件夹中的所有 CSV 文件
    def process_all_csvs(folder_path, start_date, end_date):
        print('合并人脉数据...')
        all_dfs = []
        for filename in os.listdir(folder_path):
            if filename.endswith('.csv'):
                file_path = os.path.join(folder_path, filename)
                df = process_csv(start_date, end_date, file_path)
                all_dfs.append(df)
        return pd.concat(all_dfs)

    #将S24，F24两期数据合并
    df1 = pd.read_csv('./每日数据/F24人脉匹配名单.csv')
    df2 = pd.read_csv('./每日数据/S24人脉匹配名单.csv')
    merged_df = pd.concat([df1, df2])# 将合并后的 DataFrame 写入新的 CSV 文件
    merged_df.to_csv('./每日数据/人脉匹配名单.csv', index=False)
    df1 = pd.read_csv('./每日数据/F24自定义导出.csv')
    df2 = pd.read_csv('./每日数据/S24自定义导出.csv')
    merged_df = pd.concat([df1, df2])# 将合并后的 DataFrame 写入新的 CSV 文件
    merged_df.to_csv('./每日数据/自定义导出.csv', index=False)
    
    contacts_match = pd.read_csv('./每日数据/人脉匹配名单.csv')
    apply_detail = pd.read_csv('./每日数据/自定义导出.csv')
    submit_detail = apply_detail.merge(contacts_match,how='left',on='链接')
    submit_detail['提交时间'] = pd.to_datetime(submit_detail['提交时间'])
    submit_detail = submit_detail[(submit_detail['提交时间'] >= start_date) & (submit_detail['提交时间'] <= end_date)]
    submit_detail['提交时间'] = submit_detail['提交时间'].dt.strftime('%Y-%m-%d')
    # 将字符串以"###"分隔，并取出第一部分
    submit_detail['教育经历 - 请务必填写本科学校、专业、毕业年份（如有）'] = submit_detail['教育经历 - 请务必填写本科学校、专业、毕业年份（如有）'].str.split('\n\r###\n\r').str[0]
    submit_detail['你是个技术型创始人吗？'] = submit_detail['你是个技术型创始人吗？'].str.split('\n\r###\n\r').str[0]
    submit_detail['工作经历 - 工作过的公司，职位/头衔和日期'] = submit_detail['工作经历 - 工作过的公司，职位/头衔和日期'].str.split('\n\r###\n\r').str[0]
    # 将"出生日期"列转换为日期格式
    submit_detail['出生日期 (第一位创始人)'] = pd.to_datetime(submit_detail['出生日期 (第一位创始人)'])
    # 计算到今天的年龄
    today = datetime.now()
    submit_detail['年龄 - 第一创始人'] = (today - submit_detail['出生日期 (第一位创始人)'])/pd.Timedelta(days=365.25)
    submit_detail['年龄 - 第一创始人'] = submit_detail['年龄 - 第一创始人'].astype(int)

    submit_detail = submit_detail.rename(columns={
        '核心创始人目前所处的职业阶段是？':'创始人画像',
        '请对产品进行一句话概括（15字以内，如“极其方便的云端协作文档”、“开源自动驾驶技术”）':'产品介绍',
        '教育经历 - 请务必填写本科学校、专业、毕业年份（如有）':'教育经历 - 第一创始人',
        '你是个技术型创始人吗？':'是否技术型 - 第一创始人',
        '工作经历 - 工作过的公司，职位/头衔和日期':'工作经历 - 第一创始人',
        '匹配到的申请表中人员':'申请人&创始人',
        '匹配到的联系方式':'联系方式',
    })
    # 将字符串以";"分隔为列表
    submit_detail['申请人对应的奇绩负责人_list'] = submit_detail['申请人对应的奇绩负责人'].fillna('空').str.split(';')
    # 处理列表，去除"奇绩主编"元素，新增"项目归属"列为列表中的最后一个元素
    submit_detail['申请人对应的奇绩负责人_list'] = submit_detail['申请人对应的奇绩负责人_list'].apply(lambda x: [item for item in x if item != '重复数据'])
    submit_detail['姓名'] = submit_detail['申请人对应的奇绩负责人_list'].apply(lambda x: x[-1] if x else '空')

    application_with2_intern = submit_detail[submit_detail['申请人对应的奇绩负责人_list'].apply(lambda x: len(x) >= 2)][['产品名称','链接','开表时间','提交时间','申请人&创始人','联系方式','推荐人','申请人对应的奇绩负责人']]
    application_with2_intern.to_excel('./输出结果/重复人脉记录表.xlsx',index=False)
    # 输出处理后的结果
    submit_detail = submit_detail[['产品名称','申请人&创始人','链接','开表时间','提交时间','推荐人','产品介绍','创始人画像','姓名', '教育经历 - 第一创始人','工作经历 - 第一创始人','是否技术型 - 第一创始人','年龄 - 第一创始人']]
    submit_detail.to_excel('./下载数据/提交明细.xlsx',engine='xlsxwriter',index=False)
    print('生成提交明细...')

    # 文件夹路径
    folder_path = './人脉数据总'
    # 处理并合并数据
    performance_df = process_all_csvs(folder_path, start_date, end_date)

    sup_sheet = pd.read_excel('./关联上级.xlsx')
    sup_sheet = sup_sheet[['姓名','关联上级','所属渠道']]

    performance_df = performance_df.merge(sup_sheet,how='left',on='姓名')

    clue_df = pd.read_excel('./下载数据/线索明细.xlsx')
    clue_df['提交时间'] = pd.to_datetime(clue_df['提交时间'])

    # 按照提交时间和姓名进行分组，并统计每组的数量
    count_clue = clue_df.groupby([clue_df['提交时间'], '姓名']).size().reset_index(name='线索增量')

    performance_df = performance_df[['姓名','日期','开表增量','提交增量','关联上级','所属渠道']]
    performance_df['日期'] = pd.to_datetime(performance_df['日期'])
    # 将两个DataFrame合并
    merged_df = pd.merge(performance_df, count_clue, left_on=['姓名', '日期'], right_on=['姓名', '提交时间'], how='left')
    # 修改列名
    # merged_df = merged_df.rename(columns={'日期': '考试日期', '提交时间': '提交日期', '数学成绩': '数学', '英语成绩': '英语'})

    # 调换列的顺序
    merged_df = merged_df[['姓名', '日期', '提交增量', '开表增量', '线索增量', '关联上级', '所属渠道']]
    merged_df['线索增量'] = merged_df['线索增量'].fillna(0)
    merged_df['线索增量'] = merged_df['线索增量'].astype(int)
    merged_df['日期'] = performance_df['日期'].dt.strftime('%Y-%m-%d')
    # 保存为 Excel 文件
    print('输出每日个人绩效...')
    merged_df.to_excel('./下载数据/每日个人绩效.xlsx', index=False)
