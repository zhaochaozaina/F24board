import pandas as pd
import re

def alumnus_match():
    alumnus = pd.read_excel("./S24校友名单.xlsx")
    submitted = pd.read_excel("./输出结果/提交明细.xlsx")
    alumnus_list = alumnus["name"].to_list()

    # 假设df1和df2是您的两个数据帧
    # 创建一个空的df3，用于存储符合条件的行
    alumnus_matched = pd.DataFrame(columns=submitted.columns)
    first_alumnus_matched = pd.DataFrame(columns=submitted.columns)
    submitted_notna = submitted.dropna(subset=["推荐人"])
    # 遍历df2的每一行
    for index, row in submitted_notna.iterrows():
        recommendation = row["推荐人"]
        
        for name in alumnus_list:
        # 检查df1中的每个姓名是否出现在推荐人列中
            if (name in recommendation):
                # 如果找到匹配，将该行添加到df3
                alumnus_matched = pd.concat([alumnus,row],ignore_index=True)
                #alumnus_matched = alumnus_matched.append(row, ignore_index=True)
                # 正则表达式，匹配多种分隔符（逗号、顿号、分号）
            pattern = r'[、,;，；\s]'
            # 使用正则表达式分割字符串
            # 如果没有找到分隔符，则列表只会有一个元素，即原字符串
            split_result = re.split(pattern, recommendation)
            # 获取第一部分或原字符串
            first_part_or_original = split_result[0]
            if (name in first_part_or_original):
                # 如果找到匹配，将该行添加到df3
                first_alumnus_matched = pd.concat([first_alumnus_matched,row],ignore_index=True)
                #first_alumnus_matched = first_alumnus_matched.append(row, ignore_index=True)
                # 正则表达式，匹配多种分隔符（逗号、顿号、分号）

    alumnus_group = submitted[submitted['关联上级'] == 'Huangqi']
    alumnus_matched = pd.concat([alumnus_matched, alumnus_group], axis=0)
    alumnus_matched = alumnus_matched.drop_duplicates()
    alumnus_matched.to_excel("./输出结果/校友裂变提交项目.xlsx",index=False)