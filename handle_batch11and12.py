import os
import pandas as pd

##筛选batch11中数据
def filter_csv_files(folder_path, start_date, end_date):
    # 获取文件夹中所有的 CSV 文件
    csv_files = [file for file in os.listdir(folder_path) if file.endswith('.csv')]

    # 循环处理每个 CSV 文件
    for file in csv_files:
        # 读取 CSV 文件
        file_path = os.path.join(folder_path, file)
        df = pd.read_csv(file_path, encoding='utf-8-sig')

        # 使用iloc[:, 0]定位第一列，并根据其值进行筛选
        # 注意：这里假设第一列是日期，如果不是日期，请根据实际情况调整代码
        df['\ufeff日期'] = pd.to_datetime(df.iloc[:, 0])  # 将第一列转换为日期格式
        df_filtered = df[(df['\ufeff日期'] >= start_date) & (df['\ufeff日期'] <= end_date)]
        # 删除临时添加的列
        df_filtered.drop(columns=['\ufeff日期'], inplace=True)

        # 保存筛选后的数据到新的 CSV 文件
        output_file_path = os.path.join("./人脉数据20240113-20240306", f"{file}")
        df_filtered.to_csv(output_file_path, index=False)

        print(f"筛选后的数据已保存到文件：{output_file_path}")

###将batch11中筛选出的数据与下载下来的batch12的数据进行合并
def merge_csv_files(source_folder, target_folder, result_folder):
    # 创建结果文件夹
    if not os.path.exists(result_folder):
        os.makedirs(result_folder)

    # 遍历“人脉数据”文件夹中的每个 CSV 文件
    for file in os.listdir(source_folder):
        if file.endswith('.csv'):
            # 读取源 CSV 文件的数据
            source_file_path = os.path.join(source_folder, file)
            source_df = pd.read_csv(source_file_path)

            # 在“人脉数据20240113-20240306”文件夹中查找对应名称的 CSV 文件
            target_file_path = os.path.join(target_folder, file)
            if os.path.exists(target_file_path):
                # 如果目标文件存在，则将源文件的数据追加到目标文件中
                target_df = pd.read_csv(target_file_path)
                target_df = pd.concat([target_df, source_df], ignore_index=True)
                # 将合并后的数据保存到结果文件夹中
                result_file_path = os.path.join(result_folder, file)
                target_df.to_csv(result_file_path, index=False)
                print(f"合并后的数据已保存到文件：{result_file_path}")
            else:
                # 如果目标文件不存在，则直接将源文件的数据保存到结果文件夹中
                result_file_path = os.path.join(result_folder, file)
                source_df.to_csv(result_file_path, index=False)
                print(f"合并后的数据已保存到文件：{result_file_path}")

    print("数据合并完成并保存到\"人脉数据总\"文件夹中。")

