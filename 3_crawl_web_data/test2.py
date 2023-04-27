"""
------------------------------------
# @FileName    :test2.py
# @Time        :2023/4/18 19:27
# @Author      :jack
# @description :
------------------------------------
"""
import docxcompose
import os
import win32com
import win32com.client
import bs4
import docx
import requests
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.opc.oxml import qn
from docx.shared import Pt
from docxcompose.composer import Composer
from htmldate import find_date
from bs4 import BeautifulSoup  # 导入bs4库
import aspose.words as aw  # aspose-words

if __name__ == '__main__':
    each_info_url="http://www.zgnx.gov.cn/gov/zwgk/tongzhigonggao/index.jhtml"
    response_each_info = requests.get(each_info_url)  # 用变量response_gwy保存访问网址后获得的信息
    response_each_info_content = response_each_info.content.decode('utf-8')  # 用'utf-8'的编码模式来记录网址内容，防止出现中文乱码
    each_info_soup = BeautifulSoup(response_each_info_content, features="lxml")
    text_contents = each_info_soup.find("div", id="con_01 list_bg pt5 pl20 pr20")
    print(text_contents)