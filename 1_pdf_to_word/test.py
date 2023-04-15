"""
------------------------------------
# @FileName    :test.py
# @Time        :2023/4/9 12:29
# @Author      :jack
# @description :
------------------------------------
"""
import yaml
if __name__ == '__main__':

    with open('config.yaml', 'r') as f:  # 用with读取文件更好
        configs = yaml.load(f, Loader=yaml.FullLoader)  # 按字典格式读取并返回
    # 显示读取后的内容
    pdf_path=configs["path"]["pdf_path"]
    print(pdf_path)