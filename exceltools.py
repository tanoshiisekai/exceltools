import os
from pyexcel_xlsx import get_data as get_data_xlsx


def getcolnum(colname):
    """
    列名转下标，0起
    :param colname:列名
    :return:下标
    """
    thesum = 0
    length = len(colname)
    loop = length - 1
    while loop >= 0:
        thesum = thesum + \
            (ord(colname[length - loop - 1]) - ord('A') + 1) * (26 ** loop)
        loop = loop - 1
    return thesum - 1


def colnumgenerator():
    sourcevalue = 1
    while True:
        valuestr = ""
        remainlist = []
        value = sourcevalue
        while value:
            remain = value % 26
            value = value // 26
            if remain == 0:
                remainlist.append(26)
                value = value - 1
            else:
                remainlist.append(remain)
        remainlist.reverse()
        for rem in remainlist:
            valuestr = valuestr + chr(ord('A') + rem - 1)
        sourcevalue = sourcevalue + 1
        yield valuestr


def getcolname(colnum):
    """
    下标转列名，0起
    :param colnum:列标
    :return:
    """
    count = 0
    for i in colnumgenerator():
        count = count + 1
        if count == colnum + 1:
            return i


def readdata(sourcefilename, sourcetablename, rowstart, rowend, colstart, colend, dirname="data"):
    """
    读取源数据
    :return: 数据对象
    """
    if sourcefilename.endswith(".xlsx"):
        f1 = get_data_xlsx("/".join(
            [os.getcwd(), dirname, sourcefilename]))[sourcetablename][rowstart - 1:rowend]
        f1 = [x[getcolnum(colstart): getcolnum(colend)+1] for x in f1]
        return f1
    else:
        return None
