import os
import pandas as pd
import time
import zipfile


def getCurrentDateTime():
    return time.strftime("%m.%d", time.localtime())


directory_temp = getCurrentDateTime() + "调价"


# 检查文件目录是否存在，不存在则创建
def checkFilePath(_directory):
    os.makedirs(_directory, exist_ok=True)


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


def zipFile(_directory):
    zip_compress(_directory, _directory + '.zip')


class BaseFollowTemplate:

    def traverseCurrentDirectory(self, _directory):  # 遍历当前文件夹
        for file in os.listdir(_directory):
            if file.endswith("xlsx") or file.endswith("xls"):
                self.parseExcelFileAndMakeTxt(file)
        zipFile(directory_temp)

    # 解析excel文件和写入文件
    def parseExcelFileAndMakeTxt(self, file):
        data = self.parseFile(file)
        print(data)
        self.write2Txt(file.title(), data)

    # 解析文件，返回key=sellerId，value=[(sku， price， minmum_seller）, ...]的字典
    def parseFile(self, filepath):
        # 对面输出的是一个dictionary ，key为第一列的sellerId，value为sellerId的数据所组成的list
        # list item为从第一列开始读取到的合并数据的元组（sku， price， minmum_seller）
        data = {}

        df = pd.read_excel(filepath)
        # 遍历每一行
        for row in df.itertuples():
            key = str(getattr(row, "sellerID"))
            list = data.get(key)
            if list is None:
                list = []
            data[key] = self.parseRowToListItem(list, row)
        return data

    def parseRowToListItem(self, list, row):
        listItem = (getattr(row, 'sku'), getattr(row, 'price'))
        print(listItem)
        list.append(listItem)
        return list

    # 写成文件
    def write2Txt(self, fileName, data):
        datetime = time.strftime("%m.%d", time.localtime())
        print("current date is " + datetime)
        finaldirectory = directory_temp + "/" + fileName
        checkFilePath(finaldirectory)
        for key, value in data.items():
            txtName = finaldirectory + "/" + key + "调价" + datetime + ".txt"
            fw = open(txtName, 'w')
            fw.write('sku\tprice\t\n')
            for line in value:
                for a in line:
                    fw.write(str(a))
                    fw.write('\t')
                fw.write('\n')


#  林组长站点跟帖模板
class MXFollowTemplateFunc(BaseFollowTemplate):
    isNeedMinimumRow = "false"

    def parseRowToListItem(self, list, row):
        if self.isNeedMinimumRow == "true":
            list_item = (getattr(row, 'sku'), getattr(row, 'price'), getattr(row, '_4'))
        else:
            list_item = (getattr(row, 'sku'), getattr(row, 'price'))
        print(list_item)
        list.append(list_item)
        return list

    def write2Txt(self, fileName, data):
        datetime = time.strftime("%m.%d", time.localtime())
        print("current date is " + datetime)
        finaldirectory = directory_temp + "/" + fileName
        checkFilePath(finaldirectory)
        for key, value in data.items():
            txtName = finaldirectory + "/" + key + "调价" + datetime + ".txt"
            fw = open(txtName, 'w+')
            if self.isNeedMinimumRow == "true":
                fw.write('sku\tprice\tminimum-seller-allowed-price\n')
            else:
                fw.write('sku\tprice\t\n')
            for line in value:
                for a in line:
                    fw.write(str(a))
                    fw.write('\t')
                fw.write('\n')


#  欧洲站点跟帖模板
class EUFollowTemplateFunc(BaseFollowTemplate):

    # 通过pandas 库 拿到的每一行数据是这个，从sku字段开始要一一对应
    # Pandas(Index=0, sellerID='Amazon(AAWEU_FBA)', sku='10028742@@UKFBAAAW1', _3='B0711LMYK9', _4='ASIN',
    # price='76,99', _6=11, quantity=nan, _8='a', _9='AMAZON_EU', batteries_required=False,
    # are_batteries_included=False, supplier_declared_dg_hz_regulation1='Not Applicable',
    # supplier_declared_dg_hz_regulation2='Not Applicable')
    def parseRowToListItem(self, list, row):
        list_item = (getattr(row, 'sku'), getattr(row, '_3'), getattr(row, '_4'),
                     getattr(row, 'price'), getattr(row, '_6'), getattr(row, 'quantity'),
                     getattr(row, '_8'), getattr(row, '_9'),
                     getattr(row, 'batteries_required'),
                     getattr(row, 'are_batteries_included'), getattr(row, 'supplier_declared_dg_hz_regulation1'),
                     getattr(row, 'supplier_declared_dg_hz_regulation2'))
        # print(list_item)
        list.append(list_item)
        return list

    def write2Txt(self, fileName, data):
        datetime = time.strftime("%m.%d", time.localtime())
        print("current date is " + datetime)
        final_directory = directory_temp + "/" + fileName
        checkFilePath(final_directory)
        for key, value in data.items():
            txt_name = final_directory + "/" + key + "调价" + datetime + ".txt"
            fw = open(txt_name, 'w')
            fw.write('sku\tproduct-id\tproduct-id-type\tproduct-id-type'
                     '\tprice\titem-condition\tquantity\tadd-delete\tfulfillment-center-id'
                     '\tbatteries_required\tare_batteries_included\tsupplier_declared_dg_hz_regulation1'
                     '\tsupplier_declared_dg_hz_regulation2\n')
            for line in value:
                for a in line:
                    fw.write(str(a))
                    fw.write('\t')
                fw.write('\n')
