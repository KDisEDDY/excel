# This is a sample Python script.

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.
import os
import sys
import time
import zipfile

import pandas as pd

# 创建调价模板
CREATE_TYPE_PRICE_UPDATE = 1
# 创建跟帖模板
CREATE_TYPE_FOLLOW_UP = 2
createType = CREATE_TYPE_FOLLOW_UP


def getCurrentDateTime():
    return time.strftime("%m.%d", time.localtime())


if createType == CREATE_TYPE_PRICE_UPDATE:
    directorytemp = getCurrentDateTime() + "调价"
else:
    directorytemp = getCurrentDateTime() + "跟帖"

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
        if list is None:
            list = []

        if createType == CREATE_TYPE_PRICE_UPDATE:
            data[key] = parseRowToListItem(list, row)
        else:
            data[key] = parseRowToListItemForFollow(list, row)
    return data


# 解析单行数据的元组, 并返回成新的list
def parseRowToListItem(list, row):
    if isNeedMinimumRow == "true":
        listItem = (getattr(row, 'sku'), getattr(row, 'price'), getattr(row, '_4'))
    else:
        listItem = (getattr(row, 'sku'), getattr(row, 'price'))
    print(listItem)
    list.append(listItem)
    return list


def parseRowToListItemForFollow(list, row):
    listItem = (getattr(row, 'sku'), getattr(row, '_3'), getattr(row, '_4'), getattr(row, 'price'),
                getattr(row, '_6'), getattr(row, '_7'),
                getattr(row, 'batteries_required'), getattr(row, 'supplier_declared_dg_hz_regulation1'))
    print(listItem)
    list.append(listItem)
    return list


# 写成文件
def write2Txt(fileName, data):
    datetime = time.strftime("%m.%d", time.localtime())
    print("current date is " + datetime)
    finaldirectory = directorytemp + "/" + fileName
    checkFilePath(finaldirectory)
    for key, value in data.items():
        if createType == CREATE_TYPE_PRICE_UPDATE:
            txtName = finaldirectory + "/" + key + "调价" + datetime + ".txt"
        else:
            txtName = finaldirectory + "/" + key + "跟帖" + datetime + ".txt"
        fw = open(txtName, 'w')
        if createType == CREATE_TYPE_PRICE_UPDATE:
            if isNeedMinimumRow == "true":
                fw.write('sku\tprice\tminimum-seller-allowed-price\n')
            else:
                fw.write('sku\tprice\t\n')
        else:
            fw.write('sku\tproduct-id\tproduct-id-type\tprice\titem-condition\tfulfillment-center-id'
                     '\tbatteries_required\tsupplier_declared_dg_hz_regulation1\t\n')
        for line in value:
            for a in line:
                fw.write(str(a))
                fw.write('\t')
            fw.write('\n')


# 检查文件目录是否存在，不存在则创建
def checkFilePath(_directory):
    os.makedirs(_directory, exist_ok=True)


# 解析excel文件和写入文件
def parseExcelFileAndMakeTxt(file):
    data = parseFile(file)
    print(data)
    write2Txt(file.title(), data)


def app_path():
    """Returns the base application path."""
    if hasattr(sys, 'frozen'):
        # Handles PyInstaller
        return os.path.dirname(sys.executable)  # 使用pyinstaller打包后的exe目录
    return os.path.dirname(__file__)  # 没打包前的py目录


# 遍历当前文件夹
def traverseCurrentDirectory(_directory):
    for file in os.listdir(_directory):
        if (file.endswith("xlsx") or file.endswith("xls")):
            parseExcelFileAndMakeTxt(file)
    zipFile(directorytemp)


def zipFile(_directory):
    zip_compress(_directory, _directory + '.zip')


def zip_compress(to_zip, save_zip_name):  # save_zip_name是带目录的，也可以不带就是当前目录
    # 1.先判断输出save_zip_name的上级是否存在(判断绝对目录)，否则创建目录
    save_zip_dir = os.path.split(os.path.abspath(save_zip_name))[0]  # save_zip_name的上级目录
    print(save_zip_dir)
    if not os.path.exists(save_zip_dir):
        os.makedirs(save_zip_dir)
        print('创建新目录%s' % save_zip_dir)
    f = zipfile.ZipFile(os.path.abspath(save_zip_name), 'w', zipfile.ZIP_DEFLATED)
    # 2.判断要被压缩的to_zip是否目录还是文件，是目录就遍历操作，是文件直接压缩
    if not os.path.isdir(os.path.abspath(to_zip)):  # 如果不是目录,那就是文件
        if os.path.exists(os.path.abspath(to_zip)):  # 判断文件是否存在
            f.write(to_zip)
            f.close()
            print('%s压缩为%s' % (to_zip, save_zip_name))
        else:
            print('%s文件不存在' % os.path.abspath(to_zip))
    else:
        if os.path.exists(os.path.abspath(to_zip)):  # 判断目录是否存在，遍历目录
            zipList = []
            for dir, subdirs, files in os.walk(to_zip):  # 遍历目录，加入列表
                for fileItem in files:
                    zipList.append(os.path.join(dir, fileItem))
                    # print('a',zipList[-1])
                for dirItem in subdirs:
                    zipList.append(os.path.join(dir, dirItem))
                    # print('b',zipList[-1])
            # 读取列表压缩目录和文件
            for i in zipList:
                f.write(i, i.replace(to_zip, ''))  # replace是减少压缩文件的一层目录，即压缩文件不包括to_zip这个目录
                # print('%s压缩到%s'%(i,save_zip_name))
            f.close()
            print('%s压缩为%s' % (to_zip, save_zip_name))
        else:
            print('%s文件夹不存在' % os.path.abspath(to_zip))


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    traverseCurrentDirectory(app_path())
