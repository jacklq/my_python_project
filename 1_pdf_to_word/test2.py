"""
------------------------------------
# @FileName    :test2.py
# @Time        :2023/4/9 17:40
# @Author      :jack
# @description :
------------------------------------
"""
from paddleocr import PaddleOCR, draw_ocr
from PIL import Image
import fitz
import os
if __name__ == '__main__':

    language = 'ch'
    img_path = 'C:/Users/jack8/Desktop/关于反馈全警实战练兵相关文件意见建议的函-信通处_0_new.png'
    result_img = 'C:/Users/jack8/Desktop/1result.jpg'
    ocr = PaddleOCR(use_angle_cls=True, use_gpu=False,
                    lang=language)  # need to run only once to download and load model into memory
    img_path = img_path
    result = ocr.ocr(img_path, cls=True)
    for line in result:
        # print(line[-1][0], line[-1][1])
        print(line[1][1][0])
        print(line)
        with open("a.txt", "a") as f:
            str=""
            for i in range(len(line)):
                str=str+line[i][1][0]+"\n"

            f.write(str)

