# This is a sample Python script.

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.
import os

import pandas as pd
import time
import sys

directory = "txt"
parseFileDirectory = "excel"
# 是否需要展示最小价格一列
isNeedMinimumRow = "false"

# 解析文件，返回key=sellerId，value=[(sku， price， minmum_seller）, ...]的字典
def parseFile(filePath):
    # 对面输出的是一个dictionary ，key为第一列的sellerId，value为sellerId的数据所组成的list
    # list item为从第一列开始读取到的合并数据的元组（sku， price， minmum_seller）
    data = {}

    df = pd.read_excel(filePath)
    # 遍历每一行
    for row in df.itertuples():
        key = str(getattr(row, "sellerID"))
        list = data.get(key)
        if (list is None):
            list = []
        data[key] = parseRowToListItem(list, row)
    return data

# 解析单行数据的元组, 并返回成新的list
def parseRowToListItem(list, row):
    if isNeedMinimumRow == "true":
        listItem = (getattr(row, 'sku'), getattr(row, 'price'))
    else:
        listItem = (getattr(row, 'sku'), getattr(row, 'price'), getattr(row, '_4'))
    print(listItem)
    list.append(listItem)
    return list

# 写成文件
def write2Txt(data):
    datetime = time.strftime("%m.%d", time.localtime())
    print("current date is " + datetime)
    checkFilePath(directory)
    for key, value in data.items():
        txtName = directory + "/" + key + "调价" + datetime + ".txt"
        fw = open(txtName, 'w')
        if isNeedMinimumRow == "true":
            fw.write('sku\tprice\tminimum-seller-allowed-price\n')
        else:
            fw.write('sku\tprice\t\n')
        for line in value:
            for a in line:
                fw.write(str(a))
                fw.write('\t')
            fw.write('\n')

#检查文件目录是否存在，不存在则创建
def checkFilePath(_directory):
    os.makedirs(_directory, exist_ok=True)

# 解析excel文件和写入文件
def parseExcelFileAndMakeTxt(file):
    data = parseFile(file)
    print(data)
    write2Txt(data)

def app_path():
    """Returns the base application path."""
    if hasattr(sys, 'frozen'):
        # Handles PyInstaller
        return os.path.dirname(sys.executable)  #使用pyinstaller打包后的exe目录
    return os.path.dirname(__file__)                 #没打包前的py目录

# 遍历当前文件夹
def traverseCurrentDirectory(_directory):
    for file in os.listdir(_directory):
        if (file.endswith("xlsx") or file.endswith("xls")):
            parseExcelFileAndMakeTxt(file)

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    traverseCurrentDirectory(app_path())

