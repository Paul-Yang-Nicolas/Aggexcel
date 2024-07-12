import pandas as pd
import glob
import os
import openpyxl
import config
# 从combined_data_second_sheet.xlsx中提取数据
file_com = config.location_agg
combined_data = pd.read_excel(file_com)

# 获取文件夹下所有的 Excel 文件
file_list = glob.glob(config.location_list+'*.xlsx') # 替换 ‘path_to_folder’ 为你的文件夹路径

for file in file_list:
    file_name_prefix = os.path.basename(file)

    # 从combined_data中获取与文件名匹配的数据
    relevant_data = combined_data[combined_data['备注'].str.contains(file_name_prefix)]
    try:
        xls = pd.ExcelFile(file, engine='openpyxl')
        sheet_names = xls.sheet_names
        if len(sheet_names) > 1:  # 确保有第二张分表
            sheet_name = sheet_names[config.location_sheet]  # 假设第二张分表的名称为第二个

            # 打开现有的工作簿
            wb = openpyxl.load_workbook(config.location_list + file_name_prefix)

            # 选择要操作的工作表
            sheet = wb.active

            # 计算新数据要写入的行数
            start_row = config.start_row

            # 将数据逐行写入工作表，但不包括备注字段
            for i, row in enumerate(relevant_data.itertuples(index=False), start=start_row):
                for j, value in enumerate(row, start=1):
                    if j != relevant_data.columns.get_loc("备注") + 1:  # 检查是否为备注字段的列
                        sheet.cell(row=i, column=j, value=value)

                        # 保存更改
            wb.save(config.location_list + file_name_prefix)
        print("成功使用pd.ExcelFile打开文件")
    except ValueError as ve:
        print(f"值错误: {ve}")
    except Exception as e:
        print(f"出现错误: {e}")
        # 将数据放回到原始文件的第二张分表
