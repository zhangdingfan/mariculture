# # ⚠️2024.11.12(0) 合并多个文件的数据
# import pandas as pd
# import glob
# import os
#
# # 通用的文件合并函数
# def merge_excel_files(folder_path, file_patterns, output_folder, output_filename):
#     # 获取所有符合条件的文件路径
#     file_paths = []
#     for pattern in file_patterns:
#         file_paths.extend(glob.glob(pattern))
#
#     # 打印读取的文件路径，确认正确性
#     print("读取的文件路径:")
#     for file_path in file_paths:
#         print(file_path)
#
#     # 读取所有 Excel 文件到一个列表，并打印每个 DataFrame 的列名
#     data_frames = []
#     for file in file_paths:
#         df = pd.read_excel(file)
#         data_frames.append(df)
#         print(f"\n文件: {file}")
#         print("列名:")
#         print(df.columns.tolist())
#
#     # 合并所有 DataFrame（横向拼接，即将每个文件的内容列拼接到一起）
#     combined_df = pd.concat(data_frames, axis=1)
#
#     # 保存到新的 Excel 文件
#     output_path = os.path.join(output_folder, output_filename)
#     combined_df.to_excel(output_path, index=False)
#
#     print(f"\n合并后的数据已保存到: {output_path}")
#
# # 设置路径和文件模式
# folder_path = r'C:\A. climate change and mariculture\Data_Process\new'
# output_folder = r'D:\A 科研\A. climate change and mariculture\Model\Data\Assessment'
#
# # 合并 tos_ssp126_* 和 tos_historical_5year_判断m.xlsx
# merge_excel_files(
#     folder_path,
#     [f'{folder_path}/tos_historical_5year_判断m.xlsx',f'{folder_path}/tos_ssp126_*_判断m.xlsx' ],
#     output_folder,
#     'tos_ssp126_eezpoint.xlsx'
# )
#
# # 合并 tos_ssp585_* 和 tos_historical_5year_判断m.xlsx
# merge_excel_files(
#     folder_path,
#     [f'{folder_path}/tos_historical_5year_判断m.xlsx',f'{folder_path}/tos_ssp585_*_判断m.xlsx' ],
#     output_folder,
#     'tos_ssp585_eezpoint.xlsx'
# )
#
# # 合并 so_ssp126_* 和 so_historical_5year_判断m.xlsx
# merge_excel_files(
#     folder_path,
#     [f'{folder_path}/so_historical_5year_判断m.xlsx', f'{folder_path}/so_ssp126_*_判断m.xlsx'],
#     output_folder,
#     'so_ssp126_eezpoint.xlsx'
# )
#
# # 合并 so_ssp585_* 和 so_historical_5year_判断m.xlsx
# merge_excel_files(
#     folder_path,
#     [f'{folder_path}/so_historical_5year_判断m.xlsx', f'{folder_path}/so_ssp585_*_判断m.xlsx'],
#     output_folder,
#     'so_ssp585_eezpoint.xlsx'
# )

# # ⚠️2024.11.12(0)
# import pandas as pd
# import glob
# import os
#
# # 设置路径到你的 Excel 文件夹，使用原始字符串避免转义问题
# folder_path = r'C:\A. climate change and mariculture\Data_Process\new\climate_chl'
#
# # 获取所有符合条件的文件路径
# file_patterns = [
#     f'{folder_path}/chl_historical_5year.xlsx',
#     f'{folder_path}/chl_ssp126_*.xlsx'
# ]
#
# file_paths = []
# for pattern in file_patterns:
#     file_paths.extend(glob.glob(pattern))
#
# # 打印读取的文件路径，确认正确性
# print("读取的文件路径:")
# for file_path in file_paths:
#     print(file_path)
#
# # 读取所有 Excel 文件到一个列表，并打印每个 DataFrame 的列名
# data_frames = []
# for file in file_paths:
#     df = pd.read_excel(file)
#     data_frames.append(df)
#     print(f"\n文件: {file}")
#     print("列名:")
#     print(df.columns.tolist())
#
# # 合并所有 DataFrame（横向拼接，即将每个文件的内容列拼接到一起）
# combined_df = pd.concat(data_frames, axis=1)
#
# # 保存到新的 Excel 文件
# output_path = os.path.join(folder_path, 'chl_ssp126_eezpoint.xlsx')
# combined_df.to_excel(output_path, index=False)
#
# print(f"\n合并后的数据已保存到: {output_path}")
#
# # ⚠️2024.11.12(0)
# import pandas as pd
# import glob
# import os
#
# # 设置路径到你的 Excel 文件夹，使用原始字符串避免转义问题
# folder_path = r'C:\A. climate change and mariculture\Data_Process\new\climate_chl'
#
# # 获取所有符合条件的文件路径
# file_patterns = [
#     f'{folder_path}/chl_historical_5year.xlsx',
#     f'{folder_path}/chl_ssp585_*.xlsx'
# ]
#
# file_paths = []
# for pattern in file_patterns:
#     file_paths.extend(glob.glob(pattern))
#
# # 打印读取的文件路径，确认正确性
# print("读取的文件路径:")
# for file_path in file_paths:
#     print(file_path)
#
# # 读取所有 Excel 文件到一个列表，并打印每个 DataFrame 的列名
# data_frames = []
# for file in file_paths:
#     df = pd.read_excel(file)
#     data_frames.append(df)
#     print(f"\n文件: {file}")
#     print("列名:")
#     print(df.columns.tolist())
#
# # 合并所有 DataFrame（横向拼接，即将每个文件的内容列拼接到一起）
# combined_df = pd.concat(data_frames, axis=1)
#
# # 保存到新的 Excel 文件
# output_path = os.path.join(folder_path, 'chl_ssp585_eezpoint.xlsx')
# combined_df.to_excel(output_path, index=False)
#
# print(f"\n合并后的数据已保存到: {output_path}")


# import pandas as pd
# import glob
# import os
#
# # 设置路径到你的 Excel 文件夹，使用原始字符串避免转义问题
# folder_path = r'C:\A. climate change and mariculture\Data_Process\new\climate_o2'
#
# # 获取所有符合条件的文件路径
# file_patterns = [
#     f'{folder_path}/o2_historical_5year_判断m.xlsx',
#     f'{folder_path}/o2_ssp126_*_判断m.xlsx'
# ]
#
# file_paths = []
# for pattern in file_patterns:
#     file_paths.extend(glob.glob(pattern))
#
# # 打印读取的文件路径，确认正确性
# print("读取的文件路径:")
# for file_path in file_paths:
#     print(file_path)
#
# # 读取所有 Excel 文件到一个列表，并打印每个 DataFrame 的列名
# data_frames = []
# for file in file_paths:
#     df = pd.read_excel(file)
#     data_frames.append(df)
#     print(f"\n文件: {file}")
#     print("列名:")
#     print(df.columns.tolist())
#
# # 合并所有 DataFrame（横向拼接，即将每个文件的内容列拼接到一起）
# combined_df = pd.concat(data_frames, axis=1)
#
# # 保存到新的 Excel 文件
# output_path = os.path.join(folder_path, 'o2_ssp126结果合并.xlsx')
# combined_df.to_excel(output_path, index=False)
#
# print(f"\n合并后的数据已保存到: {output_path}")
#
#
# import pandas as pd
# import glob
# import os
#
# # 设置路径到你的 Excel 文件夹，使用原始字符串避免转义问题
# folder_path = r'C:\A. climate change and mariculture\Data_Process\new\climate_o2'
#
# # 获取所有符合条件的文件路径
# file_patterns = [
#     f'{folder_path}/o2_historical_5year_判断m.xlsx',
#     f'{folder_path}/o2_ssp585_*_判断m.xlsx'
#
# ]
#
# file_paths = []
# for pattern in file_patterns:
#     file_paths.extend(glob.glob(pattern))
#
# # 打印读取的文件路径，确认正确性
# print("读取的文件路径:")
# for file_path in file_paths:
#     print(file_path)
#
# # 读取所有 Excel 文件到一个列表，并打印每个 DataFrame 的列名
# data_frames = []
# for file in file_paths:
#     df = pd.read_excel(file)
#     data_frames.append(df)
#     print(f"\n文件: {file}")
#     print("列名:")
#     print(df.columns.tolist())
#
# # 合并所有 DataFrame（横向拼接，即将每个文件的内容列拼接到一起）
# combined_df = pd.concat(data_frames, axis=1)
#
# # 保存到新的 Excel 文件
# output_path = os.path.join(folder_path, 'o2_ssp585结果合并.xlsx')
# combined_df.to_excel(output_path, index=False)
#
# print(f"\n合并后的数据已保存到: {output_path}")




# # ⚠️2024.11.08(1) 合并多个文件的数据
# import pandas as pd
# import os
#
# # 文件路径
# mariculture_map_file = r'D:\A 科研\A. climate change and mariculture\Model\Data\Mariculture map.xlsx'
# assessment_folder = r'D:\A 科研\A. climate change and mariculture\Model\Data\Assessment'
# output_file = r'D:\A 科研\A. climate change and mariculture\Model\Data\Assessment\assessment_maripoint.xlsx'
#
# # 读取Mariculture map.xlsx文件中的Type1_species sheet
# type1_species_df = pd.read_excel(mariculture_map_file, sheet_name='Type1_species', engine='openpyxl')
#
# # 遍历Assessment文件夹中的每个_eezpoint.xlsx文件
# for filename in os.listdir(assessment_folder):
#     if filename.endswith('_eezpoint.xlsx'):
#         file_path = os.path.join(assessment_folder, filename)
#         try:
#             assessment_df = pd.read_excel(file_path, engine='openpyxl')
#
#             # 筛选出列名中包含'suitable'的列
#             suitable_columns = [col for col in assessment_df.columns if 'suitable' in col]
#
#             # 根据文件名前缀判断匹配方式
#             if filename.startswith('chl'):
#                 # chl开头的文件：i, j 匹配 tos_i, tos_j
#                 merged_df = pd.merge(
#                     type1_species_df,
#                     assessment_df,
#                     left_on=['tos_i', 'tos_j'],
#                     right_on=['i', 'j'],
#                     how='left',
#                     suffixes=('', f'_{filename.split(".")[0]}')
#                 )
#             else:
#                 # 其他文件：i 匹配 tos_i, j 匹配 tos_j, species_no 匹配 species_no
#                 merged_df = pd.merge(
#                     type1_species_df,
#                     assessment_df,
#                     left_on=['tos_i', 'tos_j', 'species_no'],
#                     right_on=['i', 'j', 'species_no'],
#                     how='left',
#                     suffixes=('', f'_{filename.split(".")[0]}')
#                 )
#
#             # 更新type1_species_df，将匹配后的数据存储在一起
#             type1_species_df = merged_df
#         except Exception as e:
#             print(f"文件 {filename} 读取或处理时发生错误: {e}")
#
# # 保留前10列和列名中包含'suitable'的列
# columns_to_keep = list(type1_species_df.columns[:10]) + [col for col in type1_species_df.columns if 'suitable' in col]
# filtered_df = type1_species_df[columns_to_keep].copy()  # 确保是副本而非视图
#
# # 对列名中包含'suitable'的列重命名
# suitable_columns = [col for col in filtered_df.columns if 'suitable' in col]
# new_column_names = {}
#
# for i, col in enumerate(suitable_columns):
#     if i < 5:
#         new_name = f'chl_ssp126_t{i}'
#     elif i < 10:
#         new_name = f'chl_ssp585_t{i % 5}'
#     elif i < 15:
#         new_name = f'o2_ssp126_t{i % 5}'
#     elif i < 20:
#         new_name = f'o2_ssp585_t{i % 5}'
#     elif i < 25:
#         new_name = f'so_ssp126_t{i % 5}'
#     elif i < 30:
#         new_name = f'so_ssp585_t{i % 5}'
#     elif i < 35:
#         new_name = f'tos_ssp126_t{i % 5}'
#     else:
#         new_name = f'tos_ssp585_t{i % 5}'
#
#     new_column_names[col] = new_name
#
# filtered_df = filtered_df.rename(columns=new_column_names)
#
# # 调整列顺序
# columns_sorted = (
#         list(filtered_df.columns[:10]) +  # 保留前10列
#         [col for col in filtered_df.columns if 'ssp126_t0' in col] +
#         [col for col in filtered_df.columns if 'ssp126_t1' in col] +
#         [col for col in filtered_df.columns if 'ssp126_t2' in col] +
#         [col for col in filtered_df.columns if 'ssp126_t3' in col] +
#         [col for col in filtered_df.columns if 'ssp126_t4' in col] +
#         [col for col in filtered_df.columns if 'ssp585_t0' in col] +
#         [col for col in filtered_df.columns if 'ssp585_t1' in col] +
#         [col for col in filtered_df.columns if 'ssp585_t2' in col] +
#         [col for col in filtered_df.columns if 'ssp585_t3' in col] +
#         [col for col in filtered_df.columns if 'ssp585_t4' in col] +
#         [col for col in filtered_df.columns if col not in new_column_names]  # 其他未包含的列
# )
#
# # 保存到输出文件
# filtered_df.to_excel(output_file, index=False)
# print(f"匹配后的数据已保存到 {output_file} 中。")


# # ⚠️2024.11.08(2) 调整格式的过程文件
# import pandas as pd
#
# # 文件路径
# mariculture_map_file = r'D:\A 科研\A. climate change and mariculture\Model\Data\Mariculture map.xlsx'
# output_file = r'D:\A 科研\A. climate change and mariculture\Model\Data\Mariculture_map_transposed（过程文件可删除1）.xlsx'
#
# # 读取Mariculture map.xlsx文件中的Type1 sheet
# df = pd.read_excel(mariculture_map_file, sheet_name='Type1', engine='openpyxl')
#
# # 要竖着列的品种列
# species_columns = [
#     'Prod_Sea Bass', 'Prod_Chinese Tongue Sole', 'Prod_Large Yellow Croaker', 'Prod_Chinese Pomfret', 'Prod_Yellowtail Amberjack',
#     'Prod_Snapper', 'Prod_Red Drum', 'Prod_River Pufferfish', 'Prod_Grouper', 'Prod_Flounder', 'Prod_Oval Pomfret'
# ]
#
# # 保留的其他列
# base_columns = [
#     'UID_Fishnet_EEZ', 'UID_Fishnet_mari', 'Inversion Area（hm²）', 'Province', 'tos_i', 'tos_j', 'The 2010 yield of finfish matched to grids (t)'
# ]
#
# # 创建一个新的DataFrame用于存储结果
# expanded_rows = []
#
# # 遍历原始数据并对每一行进行扩展
# for _, row in df.iterrows():
#     for species in species_columns:
#         new_row = row[base_columns].to_dict()  # 获取基本信息的字典
#         new_row['Species'] = species  # 添加品种列
#         new_row['Yield'] = row[species]  # 添加相应品种的产量
#         expanded_rows.append(new_row)
#
# # 将扩展后的行转换为DataFrame
# expanded_df = pd.DataFrame(expanded_rows)
#
# # 保存到新的Excel文件
# expanded_df.to_excel(output_file, index=False)
#
# print(f"扩展后的数据已保存到 {output_file} 中。")
#
# import pandas as pd
#
# # 文件路径
# mariculture_map_file = r'D:\A 科研\A. climate change and mariculture\Model\Data\Mariculture map.xlsx'
# output_file = r'D:\A 科研\A. climate change and mariculture\Model\Data\Mariculture_map_transposed（过程文件可删除2）.xlsx'
#
# # 读取Mariculture map.xlsx文件中的Type1 sheet
# df = pd.read_excel(mariculture_map_file, sheet_name='Type1', engine='openpyxl')
#
# # 要竖着列的品种列
# species_columns = [
#     'Area_Sea Bass', 'Area_Chinese Tongue Sole', 'Area_Large Yellow Croaker', 'Area_Chinese Pomfret', 'Area_Yellowtail Amberjack',
#     'Area_Snapper', 'Area_Red Drum', 'Area_River Pufferfish', 'Area_Grouper', 'Area_Flounder', 'Area_Oval Pomfret'
# ]
#
# # 保留的其他列
# base_columns = [
#     'UID_Fishnet_EEZ', 'UID_Fishnet_mari', 'Inversion Area（hm²）', 'Province', 'tos_i', 'tos_j', 'The 2010 yield of finfish matched to grids (t)'
# ]
#
# # 创建一个新的DataFrame用于存储结果
# expanded_rows = []
#
# # 遍历原始数据并对每一行进行扩展
# for _, row in df.iterrows():
#     for species in species_columns:
#         new_row = row[base_columns].to_dict()  # 获取基本信息的字典
#         new_row['Species'] = species  # 添加品种列
#         new_row['Yield'] = row[species]  # 添加相应品种的产量
#         expanded_rows.append(new_row)
#
# # 将扩展后的行转换为DataFrame
# expanded_df = pd.DataFrame(expanded_rows)
#
# # 保存到新的Excel文件
# expanded_df.to_excel(output_file, index=False)
#
# print(f"扩展后的数据已保存到 {output_file} 中。")


# # ⚠️2024.11.08(3) 调整格式的过程文件
# import pandas as pd
#
# # 文件路径
# input_file = r'D:\A 科研\A. climate change and mariculture\Model\Data\Assessment\assessment_maripoint.xlsx'
#
# # 读取assessment_maripoint.xlsx文件
# assessment_df = pd.read_excel(input_file, engine='openpyxl')
#
# # 创建新的ExcelWriter对象用于保存多个Sheet
# with pd.ExcelWriter(input_file, engine='openpyxl', mode='a') as writer:
#     # 创建chl_ssp126 sheet并复制1-15列
#     chl_ssp126_df = assessment_df.iloc[:, :15]
#     chl_ssp126_df.to_excel(writer, sheet_name='chl_ssp126', index=False)
#
#     # 创建chl_ssp585 sheet并复制1-10, 16-20列
#     chl_ssp585_df = pd.concat([assessment_df.iloc[:, :10], assessment_df.iloc[:, 15:20]], axis=1)
#     chl_ssp585_df.to_excel(writer, sheet_name='chl_ssp585', index=False)
#
#     # 创建o2_ssp126 sheet并复制1-10, 21-25列
#     o2_ssp126_df = pd.concat([assessment_df.iloc[:, :10], assessment_df.iloc[:, 20:25]], axis=1)
#     o2_ssp126_df.to_excel(writer, sheet_name='o2_ssp126', index=False)
#
#     # 创建o2_ssp585 sheet并复制1-10, 26-30列
#     o2_ssp585_df = pd.concat([assessment_df.iloc[:, :10], assessment_df.iloc[:, 25:30]], axis=1)
#     o2_ssp585_df.to_excel(writer, sheet_name='o2_ssp585', index=False)
#
#     # 创建so_ssp126 sheet并复制1-10, 31-35列
#     so_ssp126_df = pd.concat([assessment_df.iloc[:, :10], assessment_df.iloc[:, 30:35]], axis=1)
#     so_ssp126_df.to_excel(writer, sheet_name='so_ssp126', index=False)
#
#     # 创建so_ssp585 sheet并复制1-10, 36-40列
#     so_ssp585_df = pd.concat([assessment_df.iloc[:, :10], assessment_df.iloc[:, 35:40]], axis=1)
#     so_ssp585_df.to_excel(writer, sheet_name='so_ssp585', index=False)
#
#     # 创建tos_ssp126 sheet并复制1-10, 41-45列
#     tos_ssp126_df = pd.concat([assessment_df.iloc[:, :10], assessment_df.iloc[:, 40:45]], axis=1)
#     tos_ssp126_df.to_excel(writer, sheet_name='tos_ssp126', index=False)
#
#     # 创建tos_ssp585 sheet并复制1-10, 46-50列
#     tos_ssp585_df = pd.concat([assessment_df.iloc[:, :10], assessment_df.iloc[:, 45:50]], axis=1)
#     tos_ssp585_df.to_excel(writer, sheet_name='tos_ssp585', index=False)
#
# print("新的sheet已添加到文件中。")


# # ⚠️2024.11.08(4) 调整格式的过程文件
# import pandas as pd
# import openpyxl
# from openpyxl import load_workbook
#
# # 文件路径
# input_file = r'D:\A 科研\A. climate change and mariculture\Model\Data\Assessment\assessment_maripoint.xlsx'
# output_file = r'D:\A 科研\A. climate change and mariculture\Model\Data\Assessment\assessment_maripoint_updated.xlsx'
#
# # 读取 Excel 文件中所有带 "_" 的 sheet
# xls = pd.ExcelFile(input_file, engine='openpyxl')
#
# # 创建一个新的字典以保存修改后的数据
# sheet_dict = {}
#
# # 遍历所有带 "_" 的 sheet
# for sheet_name in xls.sheet_names:
#     if "_" in sheet_name:
#         # 读取当前 sheet 为 DataFrame
#         df = pd.read_excel(xls, sheet_name=sheet_name, engine='openpyxl')
#
#         # 合并第 11 列和第 12, 13, 14, 15 列的内容，中间用 "-" 链接，并放在第 16-19 列
#         df['judge_t1'] = df.iloc[:, 10].fillna('').astype(str) + '-' + df.iloc[:, 11].fillna('').astype(str)
#         df['judge_t2'] = df.iloc[:, 10].fillna('').astype(str) + '-' + df.iloc[:, 12].fillna('').astype(str)
#         df['judge_t3'] = df.iloc[:, 10].fillna('').astype(str) + '-' + df.iloc[:, 13].fillna('').astype(str)
#         df['judge_t4'] = df.iloc[:, 10].fillna('').astype(str) + '-' + df.iloc[:, 14].fillna('').astype(str)
#
#         # 将修改后的 DataFrame 保存在字典中
#         sheet_dict[sheet_name] = df
#
# # 将所有修改后的数据保存到一个新的 Excel 文件中
# with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
#     for sheet_name, df in sheet_dict.items():
#         df.to_excel(writer, sheet_name=sheet_name, index=False)
#
# print("所有带 '_' 的 sheet 已更新并保存到新的文件中。")


# # ⚠️2024.11.08(5) 调整格式的过程文件
# import pandas as pd
# import os
#
# # 定义文件路径和输出文件名
# input_file = r'D:\A 科研\A. climate change and mariculture\Model\Data\Assessment\assessment_maripoint_updated.xlsx'
# output_file = r'D:\A 科研\A. climate change and mariculture\Model\Data\Assessment\assessment_maripoint_judge.xlsx'
#
# # 读取Excel文件的所有sheet名称
# xls = pd.ExcelFile(input_file)
# sheet_names = xls.sheet_names
#
# # 读取第一个sheet中的前1-10列
# first_sheet_data = pd.read_excel(xls, sheet_name=sheet_names[0], usecols=range(10))
#
# # 创建一个字典来存储提取结果
# result_sheets = {}
#
# # 遍历所有sheet名称，提取符合条件的列
# for sheet_name in sheet_names:
#     if 'ssp126' in sheet_name:
#         data = pd.read_excel(xls, sheet_name=sheet_name)
#         result_sheets.setdefault('ssp126_t1', []).append(data.iloc[:, 15])  # 第15列
#         result_sheets.setdefault('ssp126_t2', []).append(data.iloc[:, 16])  # 第16列
#         result_sheets.setdefault('ssp126_t3', []).append(data.iloc[:, 17])  # 第17列
#         result_sheets.setdefault('ssp126_t4', []).append(data.iloc[:, 18])  # 第18列
#     elif 'ssp585' in sheet_name:
#         data = pd.read_excel(xls, sheet_name=sheet_name)
#         result_sheets.setdefault('ssp585_t1', []).append(data.iloc[:, 15])  # 第15列
#         result_sheets.setdefault('ssp585_t2', []).append(data.iloc[:, 16])  # 第16列
#         result_sheets.setdefault('ssp585_t3', []).append(data.iloc[:, 17])  # 第17列
#         result_sheets.setdefault('ssp585_t4', []).append(data.iloc[:, 18])  # 第18列
#
# # 将提取的数据与第一个sheet的前1-10列合并并写入新的Excel文件
# with pd.ExcelWriter(output_file) as writer:
#     for key, data_list in result_sheets.items():
#         combined_data = pd.concat([first_sheet_data] + data_list, axis=1)
#         combined_data.to_excel(writer, sheet_name=key, index=False)
#
# print(f"提取完成，文件已保存到: {output_file}")


# # ⚠️2024.11.12(7) 插入前面遗漏的单位网格分配面积
# import pandas as pd
#
# # 文件路径
# mariculture_map_file = r'D:\A 科研\A. climate change and mariculture\Model\Data\Mariculture map.xlsx'
# assessment_file = r'D:\A 科研\A. climate change and mariculture\Model\Data\Assessment\assessment_maripoint_judge.xlsx'
#
# # 读取Mariculture map.xlsx中Type1_species的sheet
# mariculture_map_df = pd.read_excel(mariculture_map_file, sheet_name='Type1_species', engine='openpyxl')
#
# # 读取assessment_maripoint_judge.xlsx
# assessment_df = pd.read_excel(assessment_file, sheet_name=None, engine='openpyxl')
#
# # 访问The 2010 area of species matched to grids (hm2)
# map_columns = ['species_no', 'UID_Fishnet_mari', 'The 2010 area of species matched to grids (hm2)']
# mariculture_map_df = mariculture_map_df[map_columns]
#
# # 清理数据，只保留species_no, UID_Fishnet_mari, The 2010 area of species matched to grids (hm2)
# mariculture_map_df.dropna(subset=['species_no', 'UID_Fishnet_mari'], inplace=True)
# mariculture_map_df['species_no'] = mariculture_map_df['species_no'].astype(int)
# mariculture_map_df['UID_Fishnet_mari'] = mariculture_map_df['UID_Fishnet_mari'].astype(int)
#
# # 为每个sheet进行匹配并更新数据
# with pd.ExcelWriter(assessment_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
#     for sheet_name, sheet_df in assessment_df.items():
#         # 确保species_no和UID_Fishnet_mari都存在
#         if 'species_no' in sheet_df.columns and 'UID_Fishnet_mari' in sheet_df.columns:
#             # 进行匹配，使用species_no和UID_Fishnet_mari两个条件
#             sheet_df = sheet_df.merge(mariculture_map_df, on=['species_no', 'UID_Fishnet_mari'], how='left')
#
#         # 保存更新后的sheet
#         sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
#
# print("数据匹配并更新完成。")

# # ⚠️2024.11.11(6) 判断出现、损失、不变
# import pandas as pd
# import os
#
# # 文件路径
# assessment_file = r'D:\A 科研\A. climate change and mariculture\Model\Data\Assessment\assessment_maripoint_judge.xlsx'
# condition_file = r'D:\A 科研\A. climate change and mariculture\Model\Data\Judgment condition.xlsx'
#
# # 读取Judgment condition.xlsx文件，获取tos, so, o2, chl和Species influence五列
# condition_df = pd.read_excel(condition_file, sheet_name='Sheet1', usecols=['tos', 'so', 'o2', 'chl', 'Species influence'], engine='openpyxl')
#
# # 将所有匹配列转换为小写，避免大小写差异
# condition_df[['tos', 'so', 'o2', 'chl']] = condition_df[['tos', 'so', 'o2', 'chl']].apply(lambda x: x.str.lower())
#
# # 读取assessment_maripoint_judge.xlsx中的所有sheet
# sheets = pd.read_excel(assessment_file, sheet_name=None, engine='openpyxl')
#
# # 创建一个ExcelWriter对象，追加模式
# with pd.ExcelWriter(assessment_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
#     for sheet_name, sheet_df in sheets.items():
#         # 修改第11, 12, 13, 14列的列名为 chl, o2, so, tos
#         if len(sheet_df.columns) >= 14:
#             sheet_df.columns.values[10] = 'chl'
#             sheet_df.columns.values[11] = 'o2'
#             sheet_df.columns.values[12] = 'so'
#             sheet_df.columns.values[13] = 'tos'
#
#         # 确保四列存在并转换为小写
#         for col in ['chl', 'o2', 'so', 'tos']:
#             if col in sheet_df.columns:
#                 sheet_df[col] = sheet_df[col].astype(str).str.lower()
#
#         # 使用merge进行匹配
#         merged_df = pd.merge(sheet_df, condition_df, on=['chl', 'o2', 'so', 'tos'], how='left')
#
#         # 如果匹配成功，将Judgment condition.xlsx中'Species influence'列的内容添加到sheet_df最后一列
#         if 'Species influence_y' in merged_df.columns:
#             sheet_df['Species influence'] = merged_df['Species influence_y']
#         else:
#             # 如果没有匹配项，则将'Species influence'列填充为NaN
#             sheet_df['Species influence'] = pd.NA
#
#         # 保存修改后的sheet
#         sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
#
# print("匹配后的数据已保存到原文件中。")

# # ⚠️2024.11.12(8) 计算新增损失面积和产量
# import pandas as pd
#
# # 文件路径
# assessment_file = r'D:\A 科研\A. climate change and mariculture\Model\Data\Assessment\assessment_maripoint_judge.xlsx'
# output_file = r'D:\A 科研\A. climate change and mariculture\Model\Data\Assessment\assessment_maripoint_calculate.xlsx'
#
# # 读取assessment_maripoint_judge.xlsx
# assessment_df = pd.read_excel(assessment_file, sheet_name=None, engine='openpyxl')
#
# # 访问The 2010 area of species matched to grids (hm2)
# map_columns = ['species_no', 'UID_Fishnet_mari', 'The 2010 area of species matched to grids (hm2)',
#                'The 2010 yield of species matched to grids (t)']
#
# # 为每个sheet进行匹配并更新数据
# with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
#     for sheet_name, sheet_df in assessment_df.items():
#         # 确保species_no和UID_Fishnet_mari都存在
#         if 'species_no' in sheet_df.columns and 'UID_Fishnet_mari' in sheet_df.columns:
#             # 进行匹配，使用species_no和UID_Fishnet_mari两个条件
#             merged_df = sheet_df.copy()
#             for col in map_columns[2:]:
#                 if col in sheet_df.columns:
#                     merged_df[col] = sheet_df[col]
#
#             # 添加新的几列
#             merged_df['final area (hm2)'] = merged_df.apply(
#                 lambda row: row['The 2010 area of species matched to grids (hm2)'] if row[
#                                                                                           'Species influence'] == 'No change' else (
#                     0 if row['Species influence'] == 'Lost' else 2 * row[
#                         'The 2010 area of species matched to grids (hm2)']), axis=1
#             )
#             merged_df['final prod (t)'] = merged_df.apply(
#                 lambda row: row['The 2010 yield of species matched to grids (t)'] if row[
#                                                                                          'Species influence'] == 'No change' else (
#                     0 if row['Species influence'] == 'Lost' else 2 * row[
#                         'The 2010 yield of species matched to grids (t)']), axis=1
#             )
#             merged_df['loss area (hm2)'] = merged_df.apply(
#                 lambda row: 0 if row['Species influence'] == 'No change' else (
#                     row['The 2010 area of species matched to grids (hm2)'] if row[
#                                                                                   'Species influence'] == 'Lost' else 0),
#                 axis=1
#             )
#             merged_df['loss prod (t)'] = merged_df.apply(
#                 lambda row: 0 if row['Species influence'] == 'No change' else (
#                     row['The 2010 yield of species matched to grids (t)'] if row['Species influence'] == 'Lost' else 0),
#                 axis=1
#             )
#             merged_df['new area (hm2)'] = merged_df.apply(
#                 lambda row: 0 if row['Species influence'] == 'No change' else (
#                     0 if row['Species influence'] == 'Lost' else row[
#                         'The 2010 area of species matched to grids (hm2)']), axis=1
#             )
#             merged_df['new area (t)'] = merged_df.apply(
#                 lambda row: 0 if row['Species influence'] == 'No change' else (
#                     0 if row['Species influence'] == 'Lost' else row['The 2010 yield of species matched to grids (t)']),
#                 axis=1
#             )
#
#             # 保存更新后的sheet
#             merged_df.to_excel(writer, sheet_name=sheet_name, index=False)
#
# print("数据匹配并更新完成，已保存到 assessment_maripoint_calculate.xlsx。")

# # ⚠️2024.11.12(9) 结果分析
# import pandas as pd
#
# # 读取Excel文件
# data_path = r"D:\A 科研\A. climate change and mariculture\Model\Data\Assessment\assessment_maripoint_calculate.xlsx"
# xls = pd.ExcelFile(data_path)
#
# # 输出文件路径
# output_path = r"D:\A 科研\A. climate change and mariculture\Model\Data\Assessment\assessment_summary.xlsx"
#
# # 创建一个空的 DataFrame，用于存储所有sheet的结果
# sheet_results = {}
#
# # 遍历Excel文件中的每个sheet
# for sheet_name in xls.sheet_names:
#     # 读取当前sheet的数据
#     df = pd.read_excel(xls, sheet_name=sheet_name)
#
#     # 保留指定列
#     columns_to_keep = [
#         'UID_Fishnet_EEZ',
#         'UID_Fishnet_mari',
#         'Inversion Area（hm²）',
#         'Province',
#         'tos_i',
#         'tos_j',
#         'The 2010 yield of finfish matched to grids (t)'
#     ]
#     kept_df = df[columns_to_keep]
#
#     # 按照UID_Fishnet_mari分类汇总所需列
#     columns_to_sum = [
#         'The 2010 area of species matched to grids (hm2)',
#         'The 2010 yield of species matched to grids (t)',
#         'final area (hm2)',
#         'final prod (t)',
#         'loss area (hm2)',
#         'loss prod (t)',
#         'new area (hm2)',
#         'new area (t)'
#     ]
#
#     grouped_df = df.groupby('UID_Fishnet_mari')[columns_to_sum].sum().reset_index()
#
#     # 合并保留的列和汇总的列
#     merged_df = pd.merge(kept_df.drop_duplicates(subset=['UID_Fishnet_mari']), grouped_df, on='UID_Fishnet_mari', how='left')
#
#     # 计算新列并加入到数据框中
#     merged_df['final area (%)'] = (merged_df['final area (hm2)'] / merged_df['The 2010 area of species matched to grids (hm2)']) * 100
#     merged_df['final prod (%)'] = (merged_df['final prod (t)'] / merged_df['The 2010 yield of species matched to grids (t)']) * 100
#     merged_df['loss area (%)'] = (merged_df['loss area (hm2)'] / merged_df['The 2010 area of species matched to grids (hm2)']) * 100
#     merged_df['loss prod (%)'] = (merged_df['loss prod (t)'] / merged_df['The 2010 yield of species matched to grids (t)']) * 100
#     merged_df['new area (%)'] = (merged_df['new area (hm2)'] / merged_df['The 2010 area of species matched to grids (hm2)']) * 100
#     merged_df['new prod (%)'] = (merged_df['new area (t)'] / merged_df['The 2010 yield of species matched to grids (t)']) * 100
#
#     # 新建列 'loss area_total (hm2)'，当 'loss area (hm2)' 不等于 0 时，值为 'The 2010 area of species matched to grids (hm2)'
#     merged_df['loss area_total (hm2)'] = merged_df.apply(lambda row: row['The 2010 area of species matched to grids (hm2)'] if row['loss area (hm2)'] != 0 else 0, axis=1)
#
#     merged_df['final area_total (hm2)'] = merged_df.apply(
#         lambda row: row['The 2010 area of species matched to grids (hm2)'] - row['loss area_total (hm2)']
#         if row['loss area (hm2)'] != 0 else row['The 2010 area of species matched to grids (hm2)'],
#         axis=1
#     )
#
#     # 将当前sheet的结果存储到字典中
#     sheet_results[sheet_name] = merged_df
#
# # 将每个sheet的结果写入输出的Excel文件中
# with pd.ExcelWriter(output_path) as writer:
#     for sheet_name, result_df in sheet_results.items():
#         result_df.to_excel(writer, sheet_name=sheet_name, index=False)
#
# print(f"汇总和计算完成，结果已保存到: {output_path}")

# # ⚠️2024.11.12(10) 画图数据提取
# import pandas as pd
#
# # 定义输入和输出文件路径
# input_path = r'D:\A 科研\A. climate change and mariculture\Model\Data\Assessment\assessment_summary.xlsx'
# output_path = r'D:\A 科研\A. climate change and mariculture\Model\Data\Assessment\figure1_1.xlsx'
#
# # 读取Excel文件
# excel_file = pd.ExcelFile(input_path)
#
# # 创建一个空的DataFrame用于存储分类汇总的结果
# summary_df = pd.DataFrame()
#
# # 遍历每一个sheet，对数据进行处理
# for sheet_name in excel_file.sheet_names:
#     # 读取当前sheet
#     df = excel_file.parse(sheet_name)
#
#     # 确保包含需要的列
#     if 'Province' in df.columns and 'final prod (t)' in df.columns and 'final area_total (hm2)' in df.columns:
#         # 按照Province分类汇总final prod (t)和final area_total (hm2)
#         grouped = df.groupby('Province').agg({'The 2010 area of species matched to grids (hm2)': 'sum', 'The 2010 yield of species matched to grids (t)': 'sum', 'final area_total (hm2)': 'sum', 'final prod (t)': 'sum'}).reset_index()
#
#         # 添加一个列来标识来源sheet
#         grouped['Sheet Name'] = sheet_name
#
#         # 将结果添加到summary_df中
#         summary_df = pd.concat([summary_df, grouped], ignore_index=True)
#
# # 将结果保存到新的Excel文件中
# summary_df.to_excel(output_path, index=False)
#
# print(f"分类汇总结果已保存到 {output_path}")
