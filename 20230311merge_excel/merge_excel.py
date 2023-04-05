import os
import xlrd
import xlsxwriter
import glob

biao_tou = "NULL"
wei_zhi = "C:\\Users\\Administrator\\Desktop"


# 获取要合并的所有exce表格
def get_exce():
    global wei_zhi
    # wei_zhi = input("请输入Excel文件所在的目录：")

    all_exce = glob.glob(wei_zhi + "\\*.xls")
    print("该目录下有" + str(len(all_exce)) + "个excel文件：")
    if (len(all_exce) == 0):
        return 0
    else:
        for i in range(len(all_exce)):
            print(all_exce[i])
        return all_exce


# 打开Exce文件
def open_exce(name):
    fh = xlrd.open_workbook(name)
    return fh


# 获取exce文件下的所有sheet
def get_sheet(fh):
    sheets = fh.sheets()
    return sheets


# 获取sheet下有多少行数据
def get_sheetrow_num(sheet):
    return sheet.nrows


# 获取sheet下的数据
def get_sheet_data(sheet, row):
    for i in range(row):
        if (i == 0):
            global biao_tou
            biao_tou = sheet.row_values(i)
            continue
        values = sheet.row_values(i)
        all_data_temp.append(values)

    return all_data_temp


if __name__ == '__main__':
    # 新建的exce文件名字
    new_exce_name = wei_zhi + "\\总数据.xlsx"
    # 新建一个exce表
    new_exce = xlsxwriter.Workbook(new_exce_name)

    all_exce = get_exce()
    # 得到要合并的所有exce表格数据
    if (all_exce == 0):
        print("该目录下无.xls文件！请检查您输入的目录是否有误！")
        os.system('pause')
        exit()

    all_data_sheet = []
    # 用于保存合并的所有行的数据
    fh = open_exce(all_exce[0])
    # 打开文件
    sheets = get_sheet(fh)


    # 下面开始文件数据的获取
    for sheet_num in range(len(sheets)):
        all_data_temp=[]
        for exce in all_exce:
            fh = open_exce(exce)
            # 打开文件
            sheets = get_sheet(fh)
            # 获取文件下的sheet数量
            row = get_sheetrow_num(sheets[sheet_num])
            all_data_temp = get_sheet_data(sheets[sheet_num], row)
        # 表头写入
        all_data_temp.insert(0, biao_tou)
        new_sheet = new_exce.add_worksheet(sheets[sheet_num].name)
        for i in range(len(all_data_temp)):
            for j in range(len(all_data_temp[i])):
                c = all_data_temp[i][j]
                new_sheet.write(i, j, c)

    new_exce.close()
    # 关闭该exce表

    print("文件合并成功,请查看“" + wei_zhi + "”目录下的总数据.xlsx文件！")

    os.system('pause')
    os.system('pause')
