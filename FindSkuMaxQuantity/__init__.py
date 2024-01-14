import os
import re
import time

import pandas as pd

pattern = re.compile(r'[0-9]+@@[a-zA-Z0-9]+[ \s]+[0-9]+')  # 根据格式匹配（10101381@@tsmxfba1   28）


def getCurrentDateTime():
    return time.strftime("%m.%d", time.localtime())


def getOutputExcelName():
    return getCurrentDateTime() + "分货-sku汇总" + ".xlsx"


def writeInExcel(split_arrays):
    if split_arrays:
        df = pd.DataFrame(split_arrays, columns=["sku", "quantity"])
        with pd.ExcelWriter(getOutputExcelName()) as writer:
            df.to_excel(writer, sheet_name="Sheet1", index= False)


def isMatch(content, ptn):
    if ptn.match(content):
        return True
    return False


class FindSkuMaxQuantity:
    def traverseCurrentDirectory(self, _directory):
        split_arrays = []
        for file in os.listdir(_directory):
            if file.endswith("txt"):
                self.readFile(file, split_arrays)

        writeInExcel(split_arrays)

    def readFile(self, file, split_arrays):
        fr = open(file, 'r')
        for line in fr:
            line = line.strip('\n')
            if isMatch(line, pattern):
                split_array = line.split('\t')
                split_arrays.append(split_array)
        fr.close()
