
""" 替换图片背景颜色"""
import os
import sys

from PIL import Image

import yaml

def inch_to_px(inch):
    if inch == 1:
        return (295, 413)
    elif inch == 2:
        return (413, 626)
    else:
        return None
def image_matting(old_image_path, new_image_path, color,inch):
    #使用这个u2net，会报错
    os.system('backgroundremover -i "'+str(old_image_path)+'" -m "u2net_human_seg" -o "cg_output.jpg"')

    # 加上背景颜色
    no_bg_image = Image.open("cg_output.jpg")
    x, y = no_bg_image.size
    new_image = Image.new('RGB', no_bg_image.size, color=color)
    new_image.paste(no_bg_image, (0, 0, x, y), no_bg_image)
    # # 转换照片尺寸
    # img_size = inch_to_px(inch)
    # if img_size == None:
    #     print("仅支持一寸和二寸，请重新输入")
    # new_image.resize(img_size)
    new_image.save(new_image_path)



if __name__ == '__main__':
   # sys.stderr = open('error_log.txt', 'w')

    """读取参数"""
    with open('config.yaml', 'r', encoding="utf-8") as f:  # 用with读取文件更好
        configs = yaml.load(f, Loader=yaml.FullLoader)  # 按字典格式读取并返回
    old_image_path = configs["old_image_path"]
    new_image_path = configs["new_image_path"]
    color = configs["color"]
    api = configs["api"]
    inch = configs["inch"]

    """转换"""
    image_matting(old_image_path, new_image_path, color,inch)
