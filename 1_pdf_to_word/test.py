"""
------------------------------------
# @FileName    :test.py
# @Time        :2023/4/9 12:29
# @Author      :jack
# @description :
------------------------------------
"""
from paddleocr import PaddleOCR, draw_ocr
from PIL import Image
import fitz
import os

# Paddleocr目前支持的多语言语种可以通过修改lang参数进行切换
# 例如`ch`, `en`, `fr`, `german`, `korean`, `japan`
def ocrImg(language,img_path,result_img):
    ocr = PaddleOCR(use_angle_cls=True, use_gpu=False,lang=language)  # need to run only once to download and load model into memory
    img_path = img_path
    result = ocr.ocr(img_path, cls=True)
    for line in result:
        # print(line[-1][0], line[-1][1])
        print(line)
    print(result[0][1][1][0])
    # 显示结果
    image = Image.open(img_path).convert('RGB')
    boxes = [line[0] for line in result]
    txts = [line[1][0] for line in result]
    scores = [line[1][1] for line in result]
    im_show = draw_ocr(image, boxes, txts, scores, font_path='./fonts/simfang.ttf')
    im_show = Image.fromarray(im_show)
    im_show.save(result_img)
def pdf_to_jpg(name,language):
    ocr = PaddleOCR(use_angle_cls=True, use_gpu=False,lang=language)  # need to run only once to download and load model into memory
    pdfdoc=fitz.open(name)
    temp = 0
    for pg in range(pdfdoc.page_count):
        page = pdfdoc[pg]
        rotate = int(0)
        # 每个尺寸的缩放系数为2，这将为我们生成分辨率提高四倍的图像。
        zoom_x = 2.0
        zoom_y =2.0
        trans = fitz.Matrix(zoom_x, zoom_y).prerotate(rotate)
        pm = page.get_pixmap(matrix=trans, alpha=False)
        pm._writeIMG('temp.jpg',1)

        #ocr识别
        result =ocr.ocr('temp.jpg', cls=True)

        #提取文件名
        xx=os.path.splitext(name)
        filename=xx[0].split('\\')[-1]+'.txt'
        #存储结果
        with open(filename,mode='a') as f:
            for line in result:
                if line[1][1]>0.5:
                    print(line[1][0].encode('utf-8').decode('utf-8'))
                    f.write(line[1][0].encode('utf-8').decode('utf-8')+'\n')
        print(pg)
if __name__ == '__main__':
    # language = sys.argv[1]
    # img_path = sys.argv[2]
    # result_img = sys.argv[2]
    # ocrImg(language, img_path, result_img)
    language = 'ch'
    img_path = 'C:/Users/jack8/Desktop/关于反馈全警实战练兵相关文件意见建议的函-信通处_0_new.png'
    result_img = 'C:/Users/jack8/Desktop/1result.jpg'
    ocrImg(language,img_path,result_img)
    # pdf_to_jpg(r'F:/1docx.pdf','ch')
