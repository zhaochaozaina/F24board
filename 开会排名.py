# -*- coding: utf-8 -*-
"""
Created on Mon Feb 26 10:18:36 2024

@author: LBW
"""
import pandas as pd
kaihui = pd.read_excel('./F24看板（大脑数据迭代用）_【总表】过程管理记录表1.0_过程管理记录表填写明细.xlsx')
kaihui = kaihui[kaihui['绩效归属日期']=='2024/03/15']
kaihui = kaihui.drop_duplicates(subset=['姓名'])
kaihui = kaihui[['绩效归属日期','姓名','组长','渠道部门','个人开会数量']]
kaihui = kaihui.sort_values(by='个人开会数量',ascending=False)
kaihui.to_excel('开会排名.xlsx')

zhuanhua = pd.read_excel('./F24看板（大脑数据迭代用）_个人统计_All.xlsx')
zhuanhua = zhuanhua[zhuanhua['日期（计算用）']=='2024/03/15']
zhuanhua = zhuanhua[['姓名','日期（计算用）','开表增量','关联上级','所属渠道','昨日开会数','开会开表转化率']]
zhuanhua = zhuanhua.sort_values(by='开会开表转化率',ascending=False)
zhuanhua.to_excel('转化率排名.xlsx')
