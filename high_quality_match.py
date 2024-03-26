import pandas as pd

def high_quality_match():
    print('计算积分...')
    application_details = pd.read_excel("./下载数据/提交明细.xlsx")

    #famous_schools = pd.read_excel("./国内985和qs100.xlsx")["优质学校"].tolist()
    schools985 = pd.read_csv("./高优先级项目（精准项目）定义名单 副本-985+意向学校.csv",sep=',')
    #删掉有备注的学校
    schools985 = schools985[pd.isna(schools985['备注'])]
    schoolsqs100 = pd.read_csv("./高优先级项目（精准项目）定义名单 副本-QS100（蓝色待定）.csv", sep=',')
    schoolsqs100 = schoolsqs100[pd.isna(schoolsqs100['备注'])]
    #将名校命大合并
    famous_schools = pd.concat([schools985['学校'], schoolsqs100[['英文名', '中文名', '简称']]], ignore_index=True)
    famous_schools = famous_schools.stack().tolist() #保存为list
    # y_list_as_strings = [str(item) for item in my_list]
    famous_schools = [str(item) for item in famous_schools]
    # 创建是否为名校的新列并设置初始值为 "否"
    application_details["是否为名校"] = "否"
    # 遍历每行数据，检查教育经历是否包含名校
    for index, row in application_details.iterrows():
        education = row["教育经历 - 第一创始人"]
        education = str(education)
        for school in famous_schools:
            if school in education:
                application_details.at[index, "是否为名校"] = "是"
                break

    company_list = ['滴滴', '旷视', '商汤', '快手', '爱奇艺', 'iqiyi', 'adobe', 'youtube', '谷歌', 'google','Google','uber','ByteDance','Tiktok','Amazon','amazon', 'aws', 'apple', 'Apple','苹果', 'meta', 'alibaba', '阿里','Alipay', '腾讯', 'tencent', '百度', 'baidu', '京东', '蚂蚁金融', '网易', '美团', '字节', 'bytedance', '360', '新浪', '上海寻梦', '搜狐', '五八同城', '苏宁', '小米', '携程', '用友', '猎豹移动', '车之家', '唯品会', '浪潮', '同程旅游', '斗鱼', '咪咕', '鹏博士', '迅雷', '米哈游', '完美世界', '波克城市', '科大讯飞', '房多多', '美图', '美柚', '汇量科技', '创梦天地', '二三四五', '游族', '好未来', '金蝶软件', '贝壳找房', '途牛科技', '东方财富', '乐游网络', '蓝鲸人', '大众书网', '淘友天下', '多点', '蚂蚁金服', '三六零', '拼多多', '58集团', '58同城', '智联招聘', '4399', '东软', '盛趣游戏', '哔哩哔哩', '拉卡拉', '吉比特', '小红书', '学霸君', 'bilibili', 'b站', '甲骨文', 'oracle', '华为', 'huawei','Huawei', 'sap', 'linkedin', '领英', '三星', 'samsung', '微软', 'microsoft', 'msra', 'ibm', '汽车之家', '摩拜', '哈啰', 'at&t','小鹏', '小鹏汽车', '理想汽车', '蔚来汽车', '英特尔', 'intel', '英伟达', 'nvidia', '爱立信', '脸书', 'facebook', '大众点评', '支付宝', 'alipay', '大疆', '雅虎', 'yahoo', 'cisco', 'salesforce', 'netflix','deepmind',]

    # 创建一个空的列 "是否为大厂"
    application_details["是否为大厂"] = ""

    # 遍历数据行
    for i, row in application_details.iterrows():
        experience = str(row["工作经历 - 第一创始人"])       
        is_company = False
        if pd.isnull(experience):
            application_details.at[i, "是否为大厂"] = "否"
        # 检查工作经历是否包含大厂，排除实习和 "intern"
        for company in company_list:
            if (company in experience) and ("实习" not in experience) and ("intern" not in experience.lower()):
                is_company = True
                break
        # 将结果写入 "是否为大厂" 列
        if is_company:
            application_details.at[i, "是否为大厂"] = "是"
        else:
            application_details.at[i, "是否为大厂"] = "否"

    application_details['积分'] = 0
    application_details.loc[((application_details["年龄 - 第一创始人"] <= 35) & ((application_details["是否为大厂"] == "是") | (application_details["是否为名校"] == "是")) & (application_details["是否技术型 - 第一创始人"] == "是")),'积分'] = 10
    application_details.loc[((application_details["年龄 - 第一创始人"] <= 35) & ((application_details["是否为大厂"] == "是") | (application_details["是否为名校"] == "是")) & (application_details["是否技术型 - 第一创始人"] == "否")),'积分'] = 5
    application_details.loc[((application_details["年龄 - 第一创始人"] <= 35) & ((application_details["是否为大厂"] == "否") & (application_details["是否为名校"] == "否")) & (application_details["是否技术型 - 第一创始人"] == "是")),'积分'] = 6
    application_details.loc[((application_details["年龄 - 第一创始人"] <= 35) & ((application_details["是否为大厂"] == "否") & (application_details["是否为名校"] == "否")) & (application_details["是否技术型 - 第一创始人"] == "否")),'积分'] = 2

    application_details.loc[((application_details["年龄 - 第一创始人"] > 35) & (application_details["年龄 - 第一创始人"] <= 40)) & ((application_details["是否为大厂"] == "是") | (application_details["是否为名校"] == "是")) & (application_details["是否技术型 - 第一创始人"] == "是"),'积分'] = 6
    application_details.loc[((application_details["年龄 - 第一创始人"] > 35) & (application_details["年龄 - 第一创始人"] <= 40)) & ((application_details["是否为大厂"] == "是") | (application_details["是否为名校"] == "是")) & (application_details["是否技术型 - 第一创始人"] == "否"),'积分'] = 3
    application_details.loc[((application_details["年龄 - 第一创始人"] > 35) & (application_details["年龄 - 第一创始人"] <= 40)) & ((application_details["是否为大厂"] == "否") & (application_details["是否为名校"] == "否")) & (application_details["是否技术型 - 第一创始人"] == "是"),'积分'] = 3

    application_details.loc[((application_details["年龄 - 第一创始人"] > 40) & ((application_details["是否为大厂"] == "是") | (application_details["是否为名校"] == "是")) & (application_details["是否技术型 - 第一创始人"] == "是")),'积分'] = 3
    application_details.loc[((application_details["年龄 - 第一创始人"] > 40) & ((application_details["是否为大厂"] == "是") | (application_details["是否为名校"] == "是")) & (application_details["是否技术型 - 第一创始人"] == "否")),'积分'] = 1
    application_details.loc[((application_details["年龄 - 第一创始人"] > 40) & ((application_details["是否为大厂"] == "否") & (application_details["是否为名校"] == "否")) & (application_details["是否技术型 - 第一创始人"] == "是")),'积分'] = 1

    application_details.to_excel('./下载数据/提交明细.xlsx',index=False)


