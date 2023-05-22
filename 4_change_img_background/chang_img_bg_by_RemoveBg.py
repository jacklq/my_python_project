
""" 替换图片背景颜色"""
import os
import sys

from PIL import Image
from removebg import RemoveBg
import yaml

def inch_to_px(inch):
    if inch == 1:
        return (295, 413)
    elif inch == 2:
        return (413, 626)
    else:
        return None
def image_matting(old_image_path, new_image_path, api_key, color,inch):
    # API KEY获取官方网站：https://www.remove.bg/zh/api
    # LQobqo6z8Xum96oodrg6WTvx
    rmbg = RemoveBg(api_key, "error.log")
    rmbg.remove_background_from_img_file(old_image_path,"4k")
    # 打开原始图片
    image = Image.open('{}_no_bg.png'.format(old_image_path))
    # 新建背景图片
    background = Image.new('RGB', image.size,color)
    # 将原始图片粘贴到新背景上
    background.paste(image, mask=image.split()[3])
    #转换照片尺寸
    img_size=inch_to_px(inch)
    if img_size==None:
        print("仅支持一寸和二寸，请重新输入")
    background.resize(img_size)
    # 保存为新图片
    background.save(new_image_path)


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
    image_matting(old_image_path, new_image_path, api, color,inch)
