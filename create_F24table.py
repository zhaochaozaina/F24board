import pandas as pd
from datetime import datetime, timedelta
import numpy as np
from get_clockin_data import get_clockin_data 

def create_table():
    #提交明细
    sup_sheet = pd.read_excel('./关联上级.xlsx')
    application_details = pd.read_excel('./下载数据/提交明细.xlsx')

    application_details = application_details.merge(sup_sheet[['姓名','关联上级','所属渠道']],how='left',on='姓名')
    application_details.to_excel('./输出结果/提交明细.xlsx',index=False)

    everyday_performance = pd.read_excel('./下载数据/每日个人绩效.xlsx')

    # 按提交时间、姓名进行分组，并计算每组的积分总和
    application_details['提交时间'] = pd.to_datetime(application_details['提交时间'])
    count_score_df = application_details.groupby(['提交时间', '姓名'])['积分'].sum().reset_index()

    everyday_performance['日期'] = pd.to_datetime(everyday_performance['日期'])
    everyday_performance = pd.merge(everyday_performance, count_score_df, left_on=['姓名', '日期'], right_on=['姓名', '提交时间'], how='left')
    everyday_performance['积分'] = everyday_performance['积分'].fillna(0)
    everyday_performance['积分'] = everyday_performance['积分'].astype(int)
    everyday_performance['日期'] = everyday_performance['日期'].dt.strftime('%Y-%m-%d')
    everyday_performance = everyday_performance[['姓名','日期','积分','提交增量','开表增量','线索增量','关联上级','所属渠道']]
    everyday_performance.to_excel('./输出结果/每日个人绩效.xlsx',index=False)

    everyday_performance['日期'] = pd.to_datetime(everyday_performance['日期'])
    # 自startdate起统计
    all_m2 = everyday_performance.groupby(['关联上级','所属渠道']).agg({'积分': 'sum', '提交增量': 'sum', '开表增量': 'sum', '线索增量': 'sum'}).reset_index()
    all_m2 = all_m2.rename(columns={'关联上级': '小组长', '积分': '累计提交积分', '提交增量': '累计提交', '开表增量': '累计开表', '线索增量': '累计新增线索'})
    all_m3 = everyday_performance.groupby(['所属渠道']).agg({'积分': 'sum', '提交增量': 'sum', '开表增量': 'sum', '线索增量': 'sum'}).reset_index()
    all_m3 = all_m3.rename(columns={'积分': '累计提交积分', '提交增量': '累计提交', '开表增量': '累计开表', '线索增量': '累计新增线索'})
    # 昨日数据
    # 获取昨天和今天的日期
    today = pd.to_datetime('today').normalize()
    yesterday = today - pd.Timedelta(days=1)
    yesterday_m1 = everyday_performance[(everyday_performance['日期'] >= yesterday) & (everyday_performance['日期'] < today)]
    yesterday_m2 = yesterday_m1.groupby(['关联上级','所属渠道']).agg({'积分': 'sum', '提交增量': 'sum', '开表增量': 'sum', '线索增量': 'sum'}).reset_index()
    yesterday_m2 = yesterday_m2.rename(columns={'关联上级': '小组长', '积分': '昨日提交积分', '提交增量': '昨日提交数', '开表增量': '昨日开表数', '线索增量': '昨日新增线索'})
    yesterday_m3 = yesterday_m1.groupby(['所属渠道']).agg({'积分': 'sum', '提交增量': 'sum', '开表增量': 'sum', '线索增量': 'sum'}).reset_index()
    yesterday_m3 = yesterday_m3.rename(columns={'积分': '昨日提交积分', '提交增量': '昨日提交数', '开表增量': '昨日开表数', '线索增量': '昨日新增线索'})
    '''# 前三日数据
    threedays_m1 = everyday_performance[everyday_performance['日期'] == (datetime.today() - timedelta(days=3)).strftime('%Y-%m-%d')]
    threedays_m2 = threedays_m1.groupby(['关联上级','所属渠道']).agg({'积分': 'sum', '提交增量': 'sum', '开表增量': 'sum', '线索增量': 'sum'}).reset_index()
    threedays_m2 = threedays_m2.rename(columns={'关联上级': '小组长', '积分': '前3日提交积分', '提交增量': '前3日提交数', '开表增量': '前3日开表数', '线索增量': '前3日新增线索'})
    threedays_m3 = threedays_m1.groupby(['所属渠道']).agg({'积分': 'sum', '提交增量': 'sum', '开表增量': 'sum', '线索增量': 'sum'}).reset_index()
    threedays_m3 = threedays_m3.rename(columns={'积分': '前3日提交积分', '提交增量': '前3日提交数', '开表增量': '前3日开表数', '线索增量': '前3日新增线索'})'''
    # 计算上周一和本周一的日期
    last_monday = today - pd.DateOffset(days=today.weekday() + 7)
    this_monday = last_monday + pd.DateOffset(days=7)
    # 筛选上周一到周天的数据
    lastweek_m1 = everyday_performance[(everyday_performance['日期'] >= last_monday) & (everyday_performance['日期'] < this_monday)]
    lastweek_m2 = lastweek_m1.groupby(['关联上级','所属渠道']).agg({'积分': 'sum', '提交增量': 'sum', '开表增量': 'sum', '线索增量': 'sum'}).reset_index()
    lastweek_m2 = lastweek_m2.rename(columns={'关联上级': '小组长', '积分': '上周提交积分', '提交增量': '上周提交数', '开表增量': '上周开表数', '线索增量': '上周新增线索'})
    lastweek_m3 = lastweek_m1.groupby(['所属渠道']).agg({'积分': 'sum', '提交增量': 'sum', '开表增量': 'sum', '线索增量': 'sum'}).reset_index()
    lastweek_m3 = lastweek_m3.rename(columns={'积分': '上周提交积分', '提交增量': '上周提交数', '开表增量': '上周开表数', '线索增量': '上周新增线索'})    
    # 筛选本周的数据
    thisweek_m1 = everyday_performance[everyday_performance['日期'] >= this_monday]
    thisweek_m2 = thisweek_m1.groupby(['关联上级','所属渠道']).agg({'积分': 'sum', '提交增量': 'sum', '开表增量': 'sum', '线索增量': 'sum'}).reset_index()
    thisweek_m2 = thisweek_m2.rename(columns={'关联上级': '小组长', '积分': '本周提交积分', '提交增量': '本周提交数', '开表增量': '本周开表数', '线索增量': '本周新增线索'})
    thisweek_m3 = thisweek_m1.groupby(['所属渠道']).agg({'积分': 'sum', '提交增量': 'sum', '开表增量': 'sum', '线索增量': 'sum'}).reset_index()
    thisweek_m3 = thisweek_m3.rename(columns={'关联上级': '小组长', '积分': '本周提交积分', '提交增量': '本周提交数', '开表增量': '本周开表数', '线索增量': '本周新增线索'})

    # 合并四个DataFrame，按照"小组长"列进行左连接
    merged_df = pd.merge(all_m2, yesterday_m2, on=['小组长','所属渠道'], how='left')
    merged_df = pd.merge(merged_df, lastweek_m2, on=['小组长','所属渠道'], how='left')
    merged_df = pd.merge(merged_df, thisweek_m2, on=['小组长','所属渠道'], how='left')
    m2_sheet = merged_df
    #得到新人名单
    newman = sup_sheet[sup_sheet['是否新人']==1]['姓名'].tolist()
    #计算昨日的工作人数，新人工作人数，有效人力工作人数
    yesterday_clockin_data = get_clockin_data(yesterday,yesterday)
    yesterday_clockin_data = yesterday_clockin_data.rename(columns={'工作人数': '昨日工作人数','总名单':'昨日总名单'})
    yesterday_clockin_data = yesterday_clockin_data[['小组长','昨日工作人数','昨日总名单']]
    yesterday_clockin_data['昨日新人工作人数'] = yesterday_clockin_data['昨日总名单'].apply(lambda x : 1 if any(name in newman for name in x) else 0)
    yesterday_clockin_data['昨日有效人力工作人数'] = yesterday_clockin_data['昨日工作人数']-yesterday_clockin_data['昨日新人工作人数']
    yesterday_clockin_data.to_excel('./昨日打卡名单.xlsx')
    m2_sheet = m2_sheet.merge(yesterday_clockin_data.drop('昨日总名单',axis=1),how='left')

    lastweek_clockin_data = get_clockin_data(last_monday,last_monday + pd.DateOffset(days=6))
    lastweek_clockin_data = lastweek_clockin_data.rename(columns={'工作人数': '上周工作人数','总名单':'上周总名单'})
    lastweek_clockin_data = lastweek_clockin_data[['小组长','上周工作人数','上周总名单']]
    lastweek_clockin_data['上周新人工作人数'] = lastweek_clockin_data['上周总名单'].apply(lambda x : 1 if any(name in newman for name in x) else 0)
    lastweek_clockin_data['上周有效人力工作人数'] = lastweek_clockin_data['上周工作人数']-lastweek_clockin_data['上周新人工作人数']
    m2_sheet = m2_sheet.merge(lastweek_clockin_data.drop('上周总名单',axis=1),how='left')

    thisweek_clockin_data = get_clockin_data(this_monday,today)
    thisweek_clockin_data = thisweek_clockin_data.rename(columns={'工作人数': '本周工作人数','总名单':'本周总名单'})
    thisweek_clockin_data = thisweek_clockin_data[['小组长','本周工作人数','本周总名单']]
    thisweek_clockin_data['本周新人工作人数'] = thisweek_clockin_data['本周总名单'].apply(lambda x : 1 if any(name in newman for name in x) else 0)
    thisweek_clockin_data['本周有效人力工作人数'] = thisweek_clockin_data['本周工作人数']-thisweek_clockin_data['本周新人工作人数']
    m2_sheet = m2_sheet.merge(thisweek_clockin_data.drop('本周总名单',axis=1),how='left')

    m2_sheet.to_excel('./输出结果/小组统计.xlsx',index=False)

    # 合并四个DataFrame，按照"小组长"列进行左连接
    merged_df = pd.merge(all_m3, yesterday_m3, on='所属渠道', how='left')
    merged_df = pd.merge(merged_df, lastweek_m3, on='所属渠道', how='left')
    merged_df = pd.merge(merged_df, thisweek_m3, on='所属渠道', how='left')
    m3_sheet = merged_df
    m3_sheet.to_excel('./输出结果/渠道统计.xlsx',index=False)

