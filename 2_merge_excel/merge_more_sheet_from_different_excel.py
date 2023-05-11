"""
------------------------------------
# @FileName    :merge_more_sheet_from_different_excel.py
# @Time        :2023/5/9 19:19
# @Author      :jack
# @description :  将两个excel中的多个sheet合并生成一个新的excel
------------------------------------
"""
import time
import os
import sys
import yaml
import pandas as pd


if __name__ == '__main__':
    sys.stderr = open('error_log.txt', 'w')
    now_time = time.strftime('%Y-%m-%d',time.localtime(time.time()))
    with open('config.yaml', 'r', encoding="utf-8") as f:  # 用with读取文件更好
        configs = yaml.load(f, Loader=yaml.FullLoader)  # 按字典格式读取并返回
        # 显示读取后的内容
    first_excel_path = configs["first_excel_path"]
    second_excel_path = configs["second_excel_path"]
    new_merged_excel_path = configs["new_merged_excel_path"]

    # 读入表1和表2
    excel_1 = pd.read_excel(first_excel_path, sheet_name=None)
    excel_2 = pd.read_excel(second_excel_path, sheet_name=None)

    """判断sheet是否一致"""
    excel_1_sheets = excel_1.keys()
    excel_2_sheets = excel_2.keys()
    if excel_1_sheets!=excel_2_sheets:
        print("两个excel表中sheet名称不一致，请检查是否有误！")
        os.system('pause')
        exit()
    """合并多个sheet"""
    with pd.ExcelWriter(str(now_time)+new_merged_excel_path) as writer:
        for key in excel_1_sheets:
            new_excel_1 = excel_1[key].copy()
            new_excel_2 = excel_2[key].copy()
            new_excel = pd.concat([new_excel_1, new_excel_2])
            new_excel.drop_duplicates(inplace=True)
            new_excel["序号"] = range(1, len(new_excel) + 1)
            new_excel.to_excel(writer, sheet_name=key, index=False)
