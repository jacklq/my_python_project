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
    string = "2023-04-24 17:04:48 山东省公安厅2023年度面向社会招录公务员（人民警察）体能测评公告"
    date_time = string.split()[:2]
    date_time_str="".join(date_time)
    print(date_time_str)