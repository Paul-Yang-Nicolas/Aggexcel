import pandas as pd
import glob
import os
import openpyxl
import config

# 获取文件夹下所有的 Excel 文件
file_list = glob.glob(config.location_list+'*.xlsx')  # 替换 'path_to_folder' 为你的文件夹路径
# 将合并后的数据保存为一个新的 Excel 文件
# 读取每个 Excel 文件的第二张分表并合并
all_data = pd.DataFrame()
for file in file_list:
    try:
        xls = pd.ExcelFile(file, engine='openpyxl')
        sheet_names = xls.sheet_names
        if len(sheet_names) > 1:  # 确保有第二张分表
            # 假设第二张分表的名称为第二个
            sheet_name = sheet_names[config.location_sheet]  # 假设第二张分表的名称为第二个
            df = pd.read_excel(file, sheet_name=sheet_name)
            file_name = os.path.basename(file)
            # file_name_prefix = file_name[-6]
            file_name_prefix = file_name
            df['备注'] = file_name_prefix
            all_data = pd.concat([all_data, df], ignore_index=True)
        print("成功使用pd.ExcelFile打开文件")
    except ValueError as ve:
        print(f"值错误: {ve}")
    except Exception as e:
        print(f"出现错误: {e}")
        # 将数据放回到原始文件的第二张分表
    # 保存修改后的数据到 Excel 文件
# 打开现有的工作簿
wb = openpyxl.load_workbook(config.location_agg)

# 选择要操作的工作表
sheet = wb.active

# 计算新数据要写入的行数
start_row = config.start_row

# 将数据逐行写入工作表，但不包括备注字段
for i, row in enumerate(all_data.itertuples(index=False), start=start_row):
    for j, value in enumerate(row, start=1):
        sheet.cell(row=i, column=j, value=value)
# 保存更改
wb.save(config.location_agg)