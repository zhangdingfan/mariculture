# # 已完成
# # 不需每块代码单独运算
# # ---------------------------------------section1.1气候数据格式数据处理50min---------------------------------------------------
# # ⚠️目的：处理so_ssp126的数据，把txt文件转成excel文件夹中
# import os
# import re
# import pandas as pd
#
# # 定义文件路径
# input_folder = r'C:\A. climate change and mariculture\ArcGIS\data\conclim\so'
# output_folder = r'C:\A. climate change and mariculture\Data_Process\climate_so\126'
#
# # 创建输出文件夹（如果不存在）
# if not os.path.exists(output_folder):
#     os.makedirs(output_folder)
#
# # 正则表达式模式匹配 'so_all_' 开头的文件名
# pattern = re.compile(r'so_ssp126_all_.*\.txt')
#
# # 遍历输入文件夹中的所有TXT文件
# for filename in os.listdir(input_folder):
#     if pattern.match(filename):
#         # 构建完整的文件路径
#         txt_file_path = os.path.join(input_folder, filename)
#
#         # 读取TXT文件
#         df = pd.read_csv(txt_file_path, delimiter='\t')
#
#         # 生成输出Excel文件名
#         base_name = os.path.splitext(filename)[0]
#         excel_file_name = f"{base_name}.xlsx"
#         excel_file_path = os.path.join(output_folder, excel_file_name)
#
#         # 保存为Excel文件
#         df.to_excel(excel_file_path, index=False)
#
# print("文件转换完成！")
#
# # ----------------------------------------section1.2原始数据格式数据处理2.5h--------------------------------------------------
# # ⚠️目的：对所有excel后缀的文件进行分列，结果保存在原文件中
# import pandas as pd
# import os
#
# folder_path = r'C:\A. climate change and mariculture\Data_Process\climate_so\126'
#
# # 获取文件夹中所有Excel文件的文件名
# excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx') or f.endswith('.xls')]
#
# for file in excel_files:
#     file_path = os.path.join(folder_path, file)
#     df = pd.read_excel(file_path, header=None)  # 读取Excel文件内容到DataFrame，不指定列名
#
#     # 对整列数据进行分列，假设逗号是分隔符
#     df = df[0].str.split(',', expand=True)
#
#     # 可以选择将处理后的DataFrame保存回Excel文件
#     df.to_excel(file_path, index=False, header=False)
#
#     print(f"Processed file: {file}")
#
#     print("\n")
#
# # ----------------------------------------section1.3原始数据格式数据处理3h--------------------------------------------------
# # ⚠️目的：对所有excel后缀的文件挑选EEZ内的点位
# import pandas as pd
# import os
# from concurrent.futures import ThreadPoolExecutor
#
# # 原始Excel文件夹路径
# folder_path = r'C:\A. climate change and mariculture\Data_Process\climate_so\126'
#
# # 获取文件夹中所有Excel文件的文件名
# excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
#
# # 要提取的i和j值对
# conditions = [
# ……
# ]
#
# # 定义处理单个文件的函数
# def process_file(file):
#     file_path = os.path.join(folder_path, file)
#     df = pd.read_excel(file_path)  # 读取Excel文件内容到DataFrame
#
#     # 根据条件提取符合条件的行
#     extracted_data = df[df.apply(lambda row: (row['i'], row['j']) in conditions, axis=1)]
#
#     # 生成新文件名
#     base_filename, extension = os.path.splitext(file)
#     parts = base_filename.split('_')
#     parts[1] = 'eez'  # 将 'all' 替换为 'eez'
#     new_filename = '_'.join(parts) + extension
#
#     new_file_path = os.path.join(folder_path, new_filename)
#
#     # 保存提取的数据到新的Excel文件
#     extracted_data.to_excel(new_file_path, index=False, engine='xlsxwriter')
#
#     print(f"Extracted data saved to: {new_filename}")
#
# # 使用 ThreadPoolExecutor 进行并行处理
# with ThreadPoolExecutor(max_workers=16) as executor:
#     executor.map(process_file, excel_files)
#
# # ----------------------------------------section1.4原始数据格式数据处理3s--------------------------------------------------
# # ⚠️目的：将so_ssp126_eez_{year}.{month}.xlsx 按年份合并数。挑选出一年中最高温和最低温月份盐度值。
# import os
# import pandas as pd
#
# # 文件夹路径
# folder_path = r'C:\A. climate change and mariculture\Data_Process\climate_so\126'
#
# # 获取所有符合命名规则的文件
# file_pattern = "so_eez_all_{year}.{month}.xlsx"
# files = [f for f in os.listdir(folder_path) if f.startswith("so_eez_all") and f.endswith(".xlsx")]
#
# # 按年份合并文件
# yearly_files = {}
# for file in files:
#     parts = file.split('_')
#     year_month = parts[-1].split('.')
#     year = year_month[0]
#     month = year_month[1].split('.')[0]
#
#     if year not in yearly_files:
#         yearly_files[year] = []
#     yearly_files[year].append((month, file))
#
# for year, month_files in yearly_files.items():
#     combined_data = {}
#
#     # 将每个月的数据合并到一个DataFrame中
#     for month, file in month_files:
#         file_path = os.path.join(folder_path, file)
#         df = pd.read_excel(file_path)
#         df['Month'] = month
#         combined_data[month] = df
#
#     # 按'i'和'j'分组，计算每个月份'so'的最大值和最小值
#     result = pd.concat(combined_data.values())
#     result_grouped = result.groupby(['i', 'j']).agg(
#         so_min=('so', 'min'),
#         so_max=('so', 'max')
#     ).reset_index()
#
#     # 保存结果到新的Excel文件
#     output_path = os.path.join(folder_path, f'so_ssp126_eez_m_{year}.xlsx')
#     result_grouped.to_excel(output_path, index=False, sheet_name='so_min_max')
#
# # -------------------------------------------section2.1补充so空白值3s----------------------------------------------------
# # ⚠️目的：补充无so数据的点位。
# # 使用最近邻插值方法，找到每个空白值的最近非空值进行填补。
# # 使用 cKDTree 创建一个 KD 树，用于快速查找最近邻。
# # 遍历所有空值点，使用 tree.query 方法找到最近的非空值点，并用该点的值填补空白值。
# import pandas as pd
# from scipy.spatial import cKDTree
# import glob
# import os
#
# # 定义文件路径
# folder_path = r'C:\A. climate change and mariculture\Data_Process\climate_so\126'
# points_files = glob.glob(os.path.join(folder_path, 'so_ssp126_eez_m_*.xlsx'))
#
#
# # 定义函数来填补空白值
# def fill_missing_values(df):
#     # 获取需要填补的列
#     cols_to_fill = ['so_min', 'so_max']
#
#     for col in cols_to_fill:
#         # 获取所有非空值的索引
#         non_null_points = df[['i', 'j', col]].dropna()
#
#         if non_null_points.empty:
#             continue
#
#         # 创建KDTree
#         tree = cKDTree(non_null_points[['i', 'j']])
#
#         # 获取所有空值的索引
#         null_points = df[df[col].isnull()]
#
#         for idx, null_point in null_points.iterrows():
#             dist, pos = tree.query([null_point['i'], null_point['j']], k=1)
#             nearest_value = non_null_points.iloc[pos][col]
#             df.at[idx, col] = nearest_value
#
#     return df
#
#
# # 对于每个年份的点位数据文件
# for file_path in points_files:
#     # 获取年份
#     year = os.path.basename(file_path).split('_')[-1].split('.')[0]
#
#     # 读取点位数据
#     points_df = pd.read_excel(file_path)
#
#     # 填补空白值
#     filled_df = fill_missing_values(points_df)
#
#     # 定义输出文件路径
#     output_file_path = os.path.join(folder_path, f'so_ssp126_eez_m_filled_{year}.xlsx')
#
#     # 保存填补后的结果到Excel文件
#     with pd.ExcelWriter(output_file_path) as writer:
#         filled_df.to_excel(writer, index=False)
#
#     print(f"填补后的结果已保存到 {output_file_path} 文件中")
#

# # 已完成
# # 不需每块代码单独运算
# # ---------------------------------------section1.1气候数据格式数据处理50min---------------------------------------------------
# # ⚠️目的：处理so_ssp585的数据，把txt文件转成excel文件夹中
# import os
# import re
# import pandas as pd
#
# # 定义文件路径
# input_folder = r'C:\A. climate change and mariculture\ArcGIS\data\conclim\so'
# output_folder = r'C:\A. climate change and mariculture\Data_Process\climate_so\585'
#
# # 创建输出文件夹（如果不存在）
# if not os.path.exists(output_folder):
#     os.makedirs(output_folder)
#
# # 正则表达式模式匹配 'so_all_' 开头的文件名
# pattern = re.compile(r'so_ssp585_all_.*\.txt')
#
# # 遍历输入文件夹中的所有TXT文件
# for filename in os.listdir(input_folder):
#     if pattern.match(filename):
#         # 构建完整的文件路径
#         txt_file_path = os.path.join(input_folder, filename)
#
#         # 读取TXT文件
#         df = pd.read_csv(txt_file_path, delimiter='\t')
#
#         # 生成输出Excel文件名
#         base_name = os.path.splitext(filename)[0]
#         excel_file_name = f"{base_name}.xlsx"
#         excel_file_path = os.path.join(output_folder, excel_file_name)
#
#         # 保存为Excel文件
#         df.to_excel(excel_file_path, index=False)
#
# print("文件转换完成！")
#
# # ----------------------------------------section1.2原始数据格式数据处理2.5h--------------------------------------------------
# # ⚠️目的：对所有excel后缀的文件进行分列，结果保存在原文件中
# import pandas as pd
# import os
#
# folder_path = r'C:\A. climate change and mariculture\Data_Process\climate_so\585'
#
# # 获取文件夹中所有Excel文件的文件名
# excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx') or f.endswith('.xls')]
#
# for file in excel_files:
#     file_path = os.path.join(folder_path, file)
#     df = pd.read_excel(file_path, header=None)  # 读取Excel文件内容到DataFrame，不指定列名
#
#     # 对整列数据进行分列，假设逗号是分隔符
#     df = df[0].str.split(',', expand=True)
#
#     # 可以选择将处理后的DataFrame保存回Excel文件
#     df.to_excel(file_path, index=False, header=False)
#
#     print(f"Processed file: {file}")
#
#     print("\n")
#
# # ----------------------------------------section1.3原始数据格式数据处理3h--------------------------------------------------
# # ⚠️目的：对所有excel后缀的文件挑选EEZ内的点位
# import pandas as pd
# import os
# from concurrent.futures import ThreadPoolExecutor
#
# # 原始Excel文件夹路径
# folder_path = r'C:\A. climate change and mariculture\Data_Process\climate_so\585'
#
# # 获取文件夹中所有Excel文件的文件名
# excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
#
# # 要提取的i和j值对
# conditions = [
# (54, 149), (54, 150), (54, 151), (55, 149), (55, 150), (55, 151), (55, 152), (56, 148), (56, 149), (56, 150),
# (56, 151), (56, 152), (56, 153), (56, 154), (57, 143), (57, 144), (57, 145), (57, 146), (57, 148), (57, 149),
# (57, 150), (57, 151), (57, 152), (57, 153), (57, 154), (57, 155), (57, 156), (58, 143), (58, 144), (58, 145),
# (58, 146), (58, 147), (58, 148), (58, 149), (58, 150), (58, 151), (58, 152), (58, 153), (58, 154), (58, 155),
# (58, 156), (58, 157), (59, 143), (59, 144), (59, 145), (59, 146), (59, 147), (59, 148), (59, 149), (59, 150),
# (59, 151), (59, 152), (59, 153), (59, 154), (59, 155), (59, 156), (59, 157), (59, 185), (59, 186), (60, 143),
# (60, 144), (60, 145), (60, 146), (60, 147), (60, 148), (60, 149), (60, 153), (60, 154), (60, 155), (60, 156),
# (60, 157), (60, 158), (60, 185), (60, 186), (61, 144), (61, 145), (61, 146), (61, 147), (61, 148), (61, 149),
# (61, 153), (61, 154), (61, 155), (61, 156), (61, 157), (61, 158), (61, 184), (61, 185), (61, 186), (62, 144),
# (62, 145), (62, 146), (62, 147), (62, 148), (62, 149), (62, 150), (62, 151), (62, 152), (62, 153), (62, 154),
# (62, 155), (62, 156), (62, 157), (62, 158), (62, 159), (62, 184), (62, 185), (62, 186), (63, 144), (63, 145),
# (63, 146), (63, 147), (63, 148), (63, 149), (63, 150), (63, 151), (63, 152), (63, 153), (63, 154), (63, 155),
# (63, 156), (63, 157), (63, 158), (63, 159), (63, 182), (63, 183), (63, 184), (63, 185), (63, 186), (64, 144),
# (64, 145), (64, 146), (64, 147), (64, 148), (64, 149), (64, 150), (64, 151), (64, 152), (64, 153), (64, 154),
# (64, 155), (64, 156), (64, 157), (64, 158), (64, 159), (64, 181), (64, 182), (64, 183), (64, 184), (65, 144),
# (65, 145), (65, 146), (65, 147), (65, 148), (65, 149), (65, 150), (65, 151), (65, 152), (65, 153), (65, 154),
# (65, 155), (65, 156), (65, 157), (65, 158), (65, 159), (65, 160), (65, 178), (65, 179), (65, 180), (65, 181),
# (65, 182), (65, 183), (65, 184), (66, 144), (66, 145), (66, 146), (66, 147), (66, 148), (66, 149), (66, 150),
# (66, 151), (66, 152), (66, 153), (66, 154), (66, 155), (66, 156), (66, 157), (66, 158), (66, 159), (66, 160),
# (66, 172), (66, 173), (66, 174), (66, 175), (66, 176), (66, 177), (66, 178), (66, 179), (66, 180), (66, 181),
# (66, 182), (66, 183), (67, 143), (67, 144), (67, 145), (67, 146), (67, 147), (67, 148), (67, 149), (67, 150),
# (67, 151), (67, 152), (67, 153), (67, 154), (67, 155), (67, 156), (67, 157), (67, 158), (67, 159), (67, 160),
# (67, 161), (67, 162), (67, 163), (67, 164), (67, 165), (67, 166), (67, 167), (67, 168), (67, 169), (67, 170),
# (67, 171), (67, 172), (67, 173), (67, 174), (67, 175), (67, 176), (67, 177), (67, 178), (67, 179), (67, 180),
# (67, 181), (68, 140), (68, 141), (68, 142), (68, 143), (68, 144), (68, 145), (68, 146), (68, 147), (68, 148),
# (68, 149), (68, 150), (68, 151), (68, 152), (68, 153), (68, 154), (68, 155), (68, 156), (68, 157), (68, 158),
# (68, 159), (68, 160), (68, 161), (68, 162), (68, 163), (68, 164), (68, 165), (68, 166), (68, 167), (68, 168),
# (68, 169), (68, 170), (68, 171), (68, 172), (68, 173), (68, 174), (68, 175), (68, 176), (68, 177), (68, 178),
# (68, 179), (68, 180), (69, 141), (69, 142), (69, 143), (69, 144), (69, 145), (69, 146), (69, 147), (69, 148),
# (69, 149), (69, 150), (69, 151), (69, 152), (69, 153), (69, 154), (69, 155), (69, 156), (69, 157), (69, 158),
# (69, 159), (69, 160), (69, 161), (69, 162), (69, 163), (69, 164), (69, 165), (69, 166), (69, 167), (69, 168),
# (69, 169), (69, 170), (69, 171), (69, 172), (69, 173), (69, 174), (69, 175), (69, 176), (69, 177), (69, 178),
# (69, 179), (70, 139), (70, 140), (70, 141), (70, 142), (70, 143), (70, 144), (70, 145), (70, 146), (70, 147),
# (70, 148), (70, 149), (70, 150), (70, 151), (70, 152), (70, 153), (70, 154), (70, 155), (70, 156), (70, 157),
# (70, 158), (70, 159), (70, 160), (70, 161), (70, 162), (70, 163), (70, 164), (70, 165), (70, 166),
# (70, 167), (70, 168), (70, 169), (70, 170), (70, 171), (70, 172), (70, 173), (70, 174), (70, 175), (70, 176),
# (70, 177), (71, 141), (71, 142), (71, 143), (71, 144), (71, 145), (71, 146), (71, 147), (71, 148), (71, 149),
# (71, 150), (71, 151), (71, 152), (71, 153), (71, 154), (71, 155), (71, 156), (71, 157), (71, 158), (71, 159),
# (71, 160), (71, 161), (71, 162), (71, 163), (71, 164), (71, 165), (71, 166), (71, 167), (71, 168), (71, 169),
# (71, 170), (71, 171), (71, 172), (71, 173),  (71, 174), (72, 141), (72, 142), (72, 143), (72, 144), (72, 145),
# (72, 146), (72, 147), (72, 148), (72, 149), (72, 150), (72, 151), (72, 152), (72, 153), (72, 154), (72, 155),
# (72, 156), (72, 157), (72, 158), (72, 159), (72, 160), (72, 161), (72, 162), (72, 163), (72, 164), (72, 165),
# (72, 166), (72, 167), (72, 168), (72, 169), (72, 170), (72, 171), (72, 172), (72, 173), (73, 141), (73, 142),
# (73, 143),(73, 144), (73, 145), (73, 146), (73, 147), (73, 148), (73, 149), (73, 150), (73, 151), (73, 152),
# (73, 153),(73, 154), (73, 155), (73, 156), (73, 157), (73, 158), (73, 159), (73, 160), (73, 161), (73, 162),
# (73, 163),(73, 164), (73, 165), (73, 166), (73, 167), (73, 168), (73, 169), (73, 170), (73, 171), (73, 172),
# (74, 142), (74, 143), (74, 144), (74, 145), (74, 146), (74, 147),
# (74, 148), (74, 149), (74, 150), (74, 151), (74, 152), (74, 153), (74, 154), (74, 155), (74, 156), (74, 157),
# (74, 158), (74, 159), (74, 160), (74, 161), (74, 162), (74, 163), (74, 164), (74, 165), (74, 166), (74, 167),
# (74, 168), (74, 169), (74, 170), (74, 171), (75, 142), (75, 143), (75, 144), (75, 145), (75, 146), (75, 147),
# (75, 148), (75, 149), (75, 150),
# (75, 151), (75, 152), (75, 153), (75, 154), (75, 155), (75, 156), (75, 157), (75, 158), (75, 159), (75, 160),
# (75, 161), (75, 162), (75, 163), (75, 164), (75, 165), (75, 166), (75, 167), (75, 168), (76, 141), (76, 142),
# (76, 143), (76, 144), (76, 145), (76, 146), (76, 147), (76, 148), (76, 149), (76, 150), (76, 151), (76, 152),
# (76, 153), (76, 154), (76, 155), (76, 156), (76, 157), (77, 139), (77, 140), (77, 141), (77, 142), (77, 143),
# (77, 144), (77, 145), (77, 146), (77, 147), (77, 148), (77, 149), (77, 150), (77, 151), (77, 152), (77, 153),
# (77, 154), (77, 155), (77, 156), (77, 157), (78, 140), (78, 141), (78, 142), (78, 143), (78, 144), (78, 145),
# (78, 146), (78, 147), (78, 148), (78, 149), (78, 150), (78, 151), (78, 152), (78, 153), (78, 154), (78, 155), (78, 156),
# (79, 137), (79, 138), (79, 139), (79, 140), (79, 141), (79, 142), (79, 143), (79, 144), (79, 145), (79, 146),
# (79, 147), (79, 148), (79, 149), (79, 150), (79, 151), (79, 152), (79, 153), (79, 154), (79, 155), (80, 136),
# (80, 137), (80, 138), (80, 139), (80, 140), (80, 141), (80, 142), (80, 143), (80, 144), (80, 145), (80, 146),
# (80, 147), (80, 148), (80, 149), (80, 150), (80, 151), (80, 152), (80, 153), (80, 154), (81, 136), (81, 137),
# (81, 138), (81, 139), (81, 140), (81, 141), (81, 142), (81, 143), (81, 144), (81, 145), (81, 146), (81, 147),
# (81, 148), (81, 149), (81, 150), (81, 151), (81, 152), (82, 137), (82, 138), (82, 139), (82, 140), (82, 141),
# (82, 142), (82, 143), (82, 144), (82, 145), (82, 146), (82, 147), (82, 148), (82, 149), (82, 150), (82, 151),
# (83, 134), (83, 135), (83, 136), (83, 137), (83, 138), (83, 139), (83, 140), (83, 141), (83, 142), (83, 143),
# (83, 144), (83, 145), (83, 146), (83, 147), (83, 148), (83, 149), (83, 150), (84, 119), (84, 132), (84, 133),
# (84, 134), (84, 135), (84, 136), (84, 137), (84, 138), (84, 139), (84, 140), (84, 141), (84, 142), (84, 143),
# (84, 144), (84, 145), (84, 146), (84, 147), (84, 148), (84, 149), (84, 150), (85, 118), (85, 119), (85, 132),
# (85, 133), (85, 134), (85, 135), (85, 136), (85, 137), (85, 138), (85, 139), (85, 140), (85, 141), (85, 142),
# (85, 144), (85, 145), (85, 146), (85, 147), (85, 148), (85, 149), (86, 116), (86, 117), (86, 118), (86, 119),
# (86, 124), (86, 125), (86, 129), (86, 132), (86, 133), (86, 134), (86, 135), (86, 136), (86, 137), (86, 138),
# (86, 139), (86, 140), (86, 141), (86, 144), (86, 145), (86, 146), (86, 147), (86, 148), (86, 149), (87, 100),
# (87, 101), (87, 102), (87, 103), (87, 104), (87, 111), (87, 112), (87, 116), (87, 118), (87, 119), (87, 124),
# (87, 125), (87, 129), (87, 130), (87, 131), (87, 132), (87, 133), (87, 134), (87, 135), (87, 136), (87, 137),
# (87, 138), (87, 139), (87, 140), (87, 143), (87, 144), (87, 145), (87, 146), (87, 147), (87, 148), (88, 99),
# (88, 100), (88, 101), (88, 102), (88, 103), (88, 104), (88, 111), (88, 112), (88, 113), (88, 114), (88, 118),
# (88, 119), (88, 120), (88, 124), (88, 125), (88, 128), (88, 129), (88, 130), (88, 131), (88, 132), (88, 133),
# (88, 134), (88, 135), (88, 136), (88, 137), (88, 138), (88, 139), (88, 140), (88, 141), (88, 142), (88, 143),
# (88, 144), (88, 145), (88, 146), (88, 147), (89, 100), (89, 101), (89, 102), (89, 103), (89, 104), (89, 105),
# (89, 106), (89, 110), (89, 111), (89, 112), (89, 113), (89, 114), (89, 115), (89, 116), (89, 119), (89, 120),
# (89, 121), (89, 123), (89, 124), (89, 125), (89, 126), (89, 127), (89, 128), (89, 129), (89, 130), (89, 131),
# (89, 132), (89, 133), (89, 134), (89, 135), (89, 136), (89, 137), (89, 138), (89, 139), (89, 140), (89, 141),
# (89, 142), (89, 143), (89, 144), (89, 145), (89, 146), (89, 147), (90, 100), (90, 101), (90, 102), (90, 103),
# (90, 104), (90, 105), (90, 106), (90, 107), (90, 109), (90, 110), (90, 111), (90, 112),
# (90, 113), (90, 114), (90, 115), (90, 116), (90, 117), (90, 118), (90, 119), (90, 120), (90, 121), (90, 122),
# (90, 123), (90, 124), (90, 125), (90, 126), (90, 127), (90, 128), (90, 129), (90, 130), (90, 131), (90, 132),
# (90, 133), (90, 134), (90, 135), (90, 136), (90, 137), (90, 138), (90, 139), (90, 140), (90, 141), (90, 142),
# (90, 143), (90, 144), (90, 145), (90, 146), (90, 147), (91, 100), (91, 101), (91, 102), (91, 103), (91, 104),
# (91, 105), (91, 106), (91, 109), (91, 110), (91, 111), (91, 112), (91, 113), (91, 114), (91, 115), (91, 116),
# (91, 117), (91, 118), (91, 119), (91, 120), (91, 121), (91, 122), (91, 123), (91, 124), (91, 125), (91, 126),
# (91, 127), (91, 128), (91, 129), (91, 130), (91, 131), (91, 132), (91, 133), (91, 134), (91, 135), (91, 136),
# (91, 137), (91, 138), (91, 142), (91, 143), (91, 144), (91, 145), (91, 146), (92, 98), (92, 99), (92, 100),
# (92, 101), (92, 102), (92, 103), (92, 104), (92, 105), (92, 106), (92, 108), (92, 109), (92, 110), (92, 111),
# (92, 112), (92, 113), (92, 114), (92, 115), (92, 116), (92, 117), (92, 118), (92, 119), (92, 120), (92, 121),
# (92, 122), (92, 123), (92, 124), (92, 125), (92, 126), (92, 127), (92, 128), (92, 129), (92, 130), (92, 131),
# (92, 132), (92, 133), (92, 134), (92, 135), (92, 136), (92, 137), (92, 138), (92, 139), (92, 145), (92, 146),
# (93, 99), (93, 100), (93, 101), (93, 102), (93, 103), (93, 104), (93, 105), (93, 106), (93, 107), (93, 108),
# (93, 109), (93, 110), (93, 111), (93, 112), (93, 113), (93, 114), (93, 115), (93, 116), (93, 117), (93, 118),
# (93, 119), (93, 120), (93, 121), (93, 122), (93, 123), (93, 124), (93, 125), (93, 126), (93, 127), (93, 128),
# (93, 129), (93, 130), (93, 131), (93, 132), (93, 133), (93, 134), (93, 135), (93, 136), (93, 137), (93, 138),
# (93, 139), (94, 99), (94, 100), (94, 101), (94, 102), (94, 103), (94, 104), (94, 105), (94, 106), (94, 107),
# (94, 108), (94, 109), (94, 110), (94, 111), (94, 112), (94, 113), (94, 114), (94, 115), (94, 116), (94, 117),
# (94, 118), (94, 119), (94, 120), (94, 121), (94, 122), (94, 123), (94, 124), (94, 125), (94, 126), (94, 127),
# (94, 128), (94, 129), (94, 130), (94, 131), (94, 132), (94, 133), (94, 134), (94, 135), (94, 136), (94, 137),
# (94, 138), (95, 98), (95, 99), (95, 100), (95, 101), (95, 102), (95, 103), (95, 104), (95, 105), (95, 106),
# (95, 107), (95, 108), (95, 109), (95, 110), (95, 111), (95, 112), (95, 113), (95, 114), (95, 115), (95, 116),
# (95, 117), (95, 118), (95, 119), (95, 120), (95, 121), (95, 122), (95, 123), (95, 124), (95, 125), (95, 126),
# (95, 127), (95, 128), (95, 129), (95, 130), (95, 131), (95, 132), (95, 133), (95, 134), (95, 135), (95, 136),
# (95, 137), (95, 138), (96, 97), (96, 98), (96, 99), (96, 100), (96, 101), (96, 102), (96, 103), (96, 104),
# (96, 105), (96, 106), (96, 107), (96, 108), (96, 109), (96, 110), (96, 111), (96, 112), (96, 113), (96, 114),
# (96, 116), (96, 117), (96, 118), (96, 119), (96, 120), (96, 121), (96, 122), (96, 123), (96, 124), (96, 125),
# (96, 126), (96, 127), (96, 128), (96, 129), (96, 130), (96, 131), (96, 132), (96, 133), (96, 134), (96, 135),
# (96, 136), (96, 137), (96, 138), (97, 97), (97, 98), (97, 99), (97, 100), (97, 101), (97, 102), (97, 103),
# (97, 104), (97, 105), (97, 106), (97, 107), (97, 108), (97, 109), (97, 110), (97, 111), (97, 112), (97, 113),
#     (97, 119), (97, 120), (97, 121), (97, 122), (97, 123), (97, 124), (97, 125), (97, 126), (97, 127), (97, 128),
#     (97, 129), (97, 130), (97, 131), (97, 132), (97, 133), (97, 134), (97, 135), (97, 136), (97, 137),
#     (98, 97), (98, 98), (98, 99), (98, 100), (98, 101), (98, 102), (98, 103), (98, 104), (98, 105), (98, 106), (98, 107),
#     (98, 108), (98, 109), (98, 110), (98, 111), (98, 112), (98, 120), (98, 121), (98, 122), (98, 123), (98, 124),
#     (98, 125), (98, 126), (98, 127), (98, 128), (98, 129), (98, 133), (98, 134), (98, 135),
#     (99, 98), (99, 99), (99, 100), (99, 101), (99, 102), (99, 103), (99, 104), (99, 105), (99, 106), (99, 107),
#     (99, 108), (99, 109), (99, 110), (99, 111), (99, 112), (99, 123), (99, 124), (99, 125), (99, 126), (99, 127),
#     (99, 128), (99, 129), (100, 102), (100, 103), (100, 104), (100, 105), (100, 109), (100, 124), (100, 125),
#     (100, 126), (100, 127), (100, 128), (100, 129), (101, 101), (101, 102), (101, 103), (101, 104), (101, 105),
#     (102, 102)
# ]
#
# # 定义处理单个文件的函数
# def process_file(file):
#     file_path = os.path.join(folder_path, file)
#     df = pd.read_excel(file_path)  # 读取Excel文件内容到DataFrame
#
#     # 根据条件提取符合条件的行
#     extracted_data = df[df.apply(lambda row: (row['i'], row['j']) in conditions, axis=1)]
#
#     # 生成新文件名
#     base_filename, extension = os.path.splitext(file)
#     parts = base_filename.split('_')
#     parts[1] = 'eez'  # 将 'all' 替换为 'eez'
#     new_filename = '_'.join(parts) + extension
#
#     new_file_path = os.path.join(folder_path, new_filename)
#
#     # 保存提取的数据到新的Excel文件
#     extracted_data.to_excel(new_file_path, index=False, engine='xlsxwriter')
#
#     print(f"Extracted data saved to: {new_filename}")
#
# # 使用 ThreadPoolExecutor 进行并行处理
# with ThreadPoolExecutor(max_workers=16) as executor:
#     executor.map(process_file, excel_files)
#
# # ----------------------------------------section1.4原始数据格式数据处理3s--------------------------------------------------
# # ⚠️目的：将so_ssp585_eez_{year}.{month}.xlsx 按年份合并数。挑选出一年中最高温和最低温月份盐度值。
# import os
# import pandas as pd
#
# # 文件夹路径
# folder_path = r'C:\A. climate change and mariculture\Data_Process\climate_so\585'
#
# # 获取所有符合命名规则的文件
# file_pattern = "so_eez_all_{year}.{month}.xlsx"
# files = [f for f in os.listdir(folder_path) if f.startswith("so_eez_all") and f.endswith(".xlsx")]
#
# # 按年份合并文件
# yearly_files = {}
# for file in files:
#     parts = file.split('_')
#     year_month = parts[-1].split('.')
#     year = year_month[0]
#     month = year_month[1].split('.')[0]
#
#     if year not in yearly_files:
#         yearly_files[year] = []
#     yearly_files[year].append((month, file))
#
# for year, month_files in yearly_files.items():
#     combined_data = {}
#
#     # 将每个月的数据合并到一个DataFrame中
#     for month, file in month_files:
#         file_path = os.path.join(folder_path, file)
#         df = pd.read_excel(file_path)
#         df['Month'] = month
#         combined_data[month] = df
#
#     # 按'i'和'j'分组，计算每个月份'so'的最大值和最小值
#     result = pd.concat(combined_data.values())
#     result_grouped = result.groupby(['i', 'j']).agg(
#         so_min=('so', 'min'),
#         so_max=('so', 'max')
#     ).reset_index()
#
#     # 保存结果到新的Excel文件
#     output_path = os.path.join(folder_path, f'so_ssp585_eez_m_{year}.xlsx')
#     result_grouped.to_excel(output_path, index=False, sheet_name='so_min_max')
#
# # -------------------------------------------section2.1补充so空白值3s----------------------------------------------------
# # ⚠️目的：补充无so数据的点位。
# # 使用最近邻插值方法，找到每个空白值的最近非空值进行填补。
# # 使用 cKDTree 创建一个 KD 树，用于快速查找最近邻。
# # 遍历所有空值点，使用 tree.query 方法找到最近的非空值点，并用该点的值填补空白值。
# import pandas as pd
# from scipy.spatial import cKDTree
# import glob
# import os
#
# # 定义文件路径
# folder_path = r'C:\A. climate change and mariculture\Data_Process\climate_so\585'
# points_files = glob.glob(os.path.join(folder_path, 'so_ssp585_eez_m_*.xlsx'))
#
#
# # 定义函数来填补空白值
# def fill_missing_values(df):
#     # 获取需要填补的列
#     cols_to_fill = ['so_min', 'so_max']
#
#     for col in cols_to_fill:
#         # 获取所有非空值的索引
#         non_null_points = df[['i', 'j', col]].dropna()
#
#         if non_null_points.empty:
#             continue
#
#         # 创建KDTree
#         tree = cKDTree(non_null_points[['i', 'j']])
#
#         # 获取所有空值的索引
#         null_points = df[df[col].isnull()]
#
#         for idx, null_point in null_points.iterrows():
#             dist, pos = tree.query([null_point['i'], null_point['j']], k=1)
#             nearest_value = non_null_points.iloc[pos][col]
#             df.at[idx, col] = nearest_value
#
#     return df
#
#
# # 对于每个年份的点位数据文件
# for file_path in points_files:
#     # 获取年份
#     year = os.path.basename(file_path).split('_')[-1].split('.')[0]
#
#     # 读取点位数据
#     points_df = pd.read_excel(file_path)
#
#     # 填补空白值
#     filled_df = fill_missing_values(points_df)
#
#     # 定义输出文件路径
#     output_file_path = os.path.join(folder_path, f'so_ssp585_eez_m_filled_{year}.xlsx')
#
#     # 保存填补后的结果到Excel文件
#     with pd.ExcelWriter(output_file_path) as writer:
#         filled_df.to_excel(writer, index=False)
#
#     print(f"填补后的结果已保存到 {output_file_path} 文件中")
#
# # -------------------------------------------section3.1物种so判断10s------------------------------------------------------
# # ⚠️目的：判断哪个点位不再适合养殖物种的Q10Q90耐受盐度。
# import pandas as pd
# import glob
# import os
#
# # 定义文件路径
# folder_path = r'C:\A. climate change and mariculture\Data_Process\climate_so\585'
# species_file_path = r'C:\A. climate change and mariculture\Data_Process\species\species_so_q.xlsx'
#
# # 读取物种数据
# species_df = pd.read_excel(species_file_path)
#
# # 获取所有年份的点位数据文件路径
# points_files = glob.glob(os.path.join(folder_path, 'so_ssp585_eez_m_filled_*.xlsx'))
#
# # 对于每个年份的点位数据文件
# for file_path in points_files:
#     # 获取年份
#     year = os.path.basename(file_path).split('_')[-1].split('.')[0]
#
#     # 读取点位数据
#     points_df = pd.read_excel(file_path)
#
#     # 创建一个结果列表
#     results = []
#
#     # 对于每个点位
#     for _, point in points_df.iterrows():
#         i, j, point_so_min, point_so_max = point['i'], point['j'], point['so_min'], point['so_max']
#
#         # 对于每个物种
#         for _, species in species_df.iterrows():
#             species_no, species_name, species_so_min, species_so_max = species['No.'], species['species'], species[
#                 'so_q10'], species['so_q90']
#
#             # 判断物种是否能在该点位生存
#             suitable = species_so_min <= point_so_min and species_so_max >= point_so_max
#
#             result = {
#                 'i': i,
#                 'j': j,
#                 'species_no': species_no,
#                 'species': species_name,
#                 'suitable': suitable
#             }
#             results.append(result)
#
#     # 将结果转换为DataFrame
#     results_df = pd.DataFrame(results)
#
#     # 定义输出文件路径
#     output_file_path = os.path.join(folder_path, f'so_ssp585_eez_filled_species_q_{year}.xlsx')
#
#     # 保存结果到Excel文件
#     with pd.ExcelWriter(output_file_path) as writer:
#         results_df.to_excel(writer, index=False)
#
#     print(f"结果已保存到 {output_file_path} 文件中")
#
# # -------------------------------------------section3.2物种so判断10s------------------------------------------------------
# # ⚠️目的：判断哪个点位不再适合养殖物种的最大最小耐受盐度。
# # ⚠️注意：把_m去掉是因为后续可能会重复加载该前缀的文件导致报错
# import pandas as pd
# import glob
# import os
#
# # 定义文件路径
# folder_path = r'C:\A. climate change and mariculture\Data_Process\climate_so\585'
# species_file_path = r'C:\A. climate change and mariculture\Data_Process\species\species_so_m.xlsx'
#
# # 读取物种数据
# species_df = pd.read_excel(species_file_path)
#
# # 获取所有年份的点位数据文件路径
# points_files = glob.glob(os.path.join(folder_path, 'so_ssp585_eez_m_filled_*.xlsx'))
#
# # 对于每个年份的点位数据文件
# for file_path in points_files:
#     # 获取年份
#     year = os.path.basename(file_path).split('_')[-1].split('.')[0]
#
#     # 读取点位数据
#     points_df = pd.read_excel(file_path)
#
#     # 创建一个结果列表
#     results = []
#
#     # 对于每个点位
#     for _, point in points_df.iterrows():
#         i, j, point_so_min, point_so_max = point['i'], point['j'], point['so_min'], point['so_max']
#
#         # 对于每个物种
#         for _, species in species_df.iterrows():
#             species_no, species_name, species_so_min, species_so_max = species['No.'], species['species'], species[
#                 'so_min'], species['so_max']
#
#             # 判断物种是否能在该点位生存
#             suitable = species_so_min <= point_so_min and species_so_max >= point_so_max
#
#             result = {
#                 'i': i,
#                 'j': j,
#                 'species_no': species_no,
#                 'species': species_name,
#                 'suitable': suitable
#             }
#             results.append(result)
#
#     # 将结果转换为DataFrame
#     results_df = pd.DataFrame(results)
#
#     # 定义输出文件路径
#     output_file_path = os.path.join(folder_path, f'so_ssp585_eez_filled_species_m_{year}.xlsx')
#
#     # 保存结果到Excel文件
#     with pd.ExcelWriter(output_file_path) as writer:
#         results_df.to_excel(writer, index=False)
#
#     print(f"结果已保存到 {output_file_path} 文件中")


# # -------------------------------------------⚠️new5个年次判断：section4.1物种so判断最大最小平均值10s---------------------------
# # 接2.1后面，新的条件来判断---------------------------
# import pandas as pd
# import os
#
# def process_files(input_folder, output_folder, file_prefix, output_filename, years_range):
#     """
#     处理Excel文件，将它们合并，并计算每组的最小值和最大值。
#
#     参数:
#     - input_folder: 输入文件夹路径
#     - output_folder: 输出文件夹路径
#     - file_prefix: 文件名前缀
#     - output_filename: 输出文件名
#     - years_range: 年份范围
#     """
#     # 确保输出文件夹存在，如果不存在则创建
#     os.makedirs(output_folder, exist_ok=True)
#
#     # 初始化一个DataFrame用于存储合并后的结果
#     combined_df = pd.DataFrame()
#
#     # 遍历输入文件夹中的所有Excel文件
#     for filename in os.listdir(input_folder):
#         if filename.endswith('.xlsx') and filename.startswith(file_prefix):
#             file_path = os.path.join(input_folder, filename)
#             # 读取Excel文件中的数据
#             df = pd.read_excel(file_path)
#             # 合并到主DataFrame中
#             combined_df = pd.concat([combined_df, df])
#
#     # 按照i和j分组，计算每组的最小值和最大值
#     result_df = combined_df.groupby(['i', 'j']).agg({'so_min': 'min', 'so_max': 'max'}).reset_index()
#
#     # 保存结果到新的Excel文件中
#     output_file_path = os.path.join(output_folder, output_filename)
#     result_df.to_excel(output_file_path, index=False)
#
#     print(f"结果已保存到 {output_file_path}")
#
# # 文件夹路径
# input_folder = r'C:\A. climate change and mariculture\Data_Process\old\climate_so'
# output_folder = r'C:\A. climate change and mariculture\Data_Process\new'
#
# # 调用函数处理不同的文件
# process_files(os.path.join(input_folder, 'historical'), output_folder, 'so_historical_eez_m_filled_',
#               'so_historical_5year.xlsx', '5year')
# process_files(os.path.join(input_folder, '126'), output_folder, 'so_ssp126_eez_m_filled_202',
#               'so_ssp126_20-24_5year.xlsx', '20-24_5year')
# process_files(os.path.join(input_folder, '126'), output_folder, 'so_ssp126_eez_m_filled_203',
#               'so_ssp126_30-34_5year.xlsx', '30-34_5year')
# process_files(os.path.join(input_folder, '126'), output_folder, 'so_ssp126_eez_m_filled_204',
#               'so_ssp126_40-44_5year.xlsx', '40-44_5year')
# process_files(os.path.join(input_folder, '126'), output_folder, 'so_ssp126_eez_m_filled_205',
#               'so_ssp126_50-54_5year.xlsx', '50-54_5year')
# process_files(os.path.join(input_folder, '585'), output_folder, 'so_ssp585_eez_m_filled_202',
#               'so_ssp585_20-24_5year.xlsx', '20-24_5year')
# process_files(os.path.join(input_folder, '585'), output_folder, 'so_ssp585_eez_m_filled_203',
#               'so_ssp585_30-34_5year.xlsx', '30-34_5year')
# process_files(os.path.join(input_folder, '585'), output_folder, 'so_ssp585_eez_m_filled_204',
#               'so_ssp585_40-44_5year.xlsx', '40-44_5year')
# process_files(os.path.join(input_folder, '585'), output_folder, 'so_ssp585_eez_m_filled_205',
#               'so_ssp585_50-54_5year.xlsx', '50-54_5year')
#
