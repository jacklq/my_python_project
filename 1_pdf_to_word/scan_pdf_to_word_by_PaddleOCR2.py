"""
------------------------------------
# @FileName    :scan_pdf_to_word_by_PaddleOCR2.py
# @Time        :2023/4/9 10:11
# @Author      :jack
# @description :图片文字识别
------------------------------------
"""


import cv2
from math import *
import numpy as np
from PIL import Image
from paddleocr import PaddleOCR, draw_ocr


def img_match(img_address):
    # Paddleocr目前支持的多语言语种可以通过修改lang参数进行切换
    # 例如：`ch`, `en`, `fr`, `german`, `korean`, `japan`
    # 这里 use_angle_cls=False 为不使用自定义训练集
    ocr = PaddleOCR(use_angle_cls=False, lang="ch", use_gpu=False)
    # use_angle_cls=True使用训练模型，模型放在models目录下
    # ocr = PaddleOCR(use_angle_cls=True,lang="ch",
    #                 rec_model_dir='../models/ch_PP-OCRv3_rec_slim_infer/',
    #                 cls_model_dir='../models/ch_ppocr_mobile_v2.0_cls_slim_infer/',
    #                 det_model_dir='../models/ch_PP-OCRv3_det_slim_infer/')
    src_img = cv2.imdecode(np.fromfile(img_address, dtype=np.uint8), cv2.IMREAD_COLOR)
    #src_img = cv2.imread(img_address)
    h, w = src_img.shape[:2]
    big = int(sqrt(h * h + w * w))
    big_img = np.empty((big, big, src_img.ndim), np.uint8)
    yoff = round((big - h) / 2)
    xoff = round((big - w) / 2)
    big_img[yoff:yoff + h, xoff:xoff + w] = src_img
    # 文字识别
    matRotate = cv2.getRotationMatrix2D((big * 0.5, big * 0.5), 0, 1)
    dst = cv2.warpAffine(big_img, matRotate, (big, big))
    result = ocr.ocr(dst, cls=True)
    boxes = [line[0] for line in result]
    txts = [line[1][0] for line in result]
    scores = [line[1][1] for line in result]
    # simsun.ttc 是一款很常见、实用的电脑字体，这里作为识别的模板
    # 我们利用该模板进行文字识别
    im_show = draw_ocr(dst, boxes, txts, scores, font_path='C:/Users/jack8/Desktop/simsun.ttc')
    im_show = Image.fromarray(im_show)
    img = np.asarray(im_show)
    # 展示结果
    cv2.imshow('img', img)
    cv2.waitKey(0)
    # 图片识别结果保存在代码同目录下
    # im_show.save('result.jpg')
    # 关闭页面
    cv2.destroyAllWindows()
    pass


if __name__ == '__main__':
    print("———————————————————— start ————————————————————\n")
    # 图片路径自己设置，下面是我本地的路径，记得替换！！！
    img_match('C:/Users/jack8/Desktop/关于反馈全警实战练兵相关文件意见建议的函-信通处_0_new.png')
    print("———————————————————— end ————————————————————\n")