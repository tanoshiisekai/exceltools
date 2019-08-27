import exceltools
from newworkbook import NewWorkbook


def parse(data):
    datatitle = data[0]
    data = data[1:]
    resultlist = []
    for row in data:
        classname = row[exceltools.getcolnum("B")]
        finishedksh = []
        unfinishedksh = []
        totalksh = 0
        finishedtotal = 0
        unfinishedtotal = 0
        for col in range(exceltools.getcolnum("D"), len(row) - 1, 2):
            tname = str(row[col])
            ksh = str(row[col+1])
            if len(ksh.strip()) != 0:
                totalksh = totalksh + int(ksh)
                if len(tname.strip()) == 0:
                    unfinishedksh.append((datatitle[col], "", ksh))
                    unfinishedtotal = unfinishedtotal + int(ksh)
                else:
                    finishedksh.append((datatitle[col], tname, ksh))
                    finishedtotal = finishedtotal + int(ksh)
        resultlist.append({"班级名称": classname,
                           "总课时": totalksh,
                           "未排课总课时": unfinishedtotal,
                           "未排课科目": unfinishedksh,
                           "已排课总课时": finishedtotal,
                           "已排课科目": finishedksh})
    return resultlist


def summary(data):
    teacherdict = {}
    for index, val in enumerate(data):
        ksdata = val["已排课科目"]
        print(ksdata)
        for ks in ksdata:
            if ks[1] in teacherdict.keys():
                teacherdict[ks[1]] = int(teacherdict[ks[1]]) + int(ks[2])
            else:
                teacherdict[ks[1]] = int(ks[2])
    with NewWorkbook("汇总结果.xlsx") as wb:
        sheet1 = wb.insertsheet("汇总详情")
        keylist = ["班级名称", "课程", "教师", "未排课课时", "已排课课时"]
        wb.insertrow(sheet1, "A1", keylist)
        linum = 2
        for index, val in enumerate(data):
            if val["总课时"] == 0:
                continue
            lstart = linum
            for pair in val["未排课科目"]:
                wb.insertcell(
                    sheet1, linum, exceltools.getcolnum("A"), val["班级名称"])
                wb.insertcell(
                    sheet1, linum, exceltools.getcolnum("B"), pair[0])
                wb.insertcell(
                    sheet1, linum, exceltools.getcolnum("C"), pair[1])
                wb.insertcell(
                    sheet1, linum, exceltools.getcolnum("D"), int(pair[2]))
                linum = linum + 1
            for pair in val["已排课科目"]:
                wb.insertcell(
                    sheet1, linum, exceltools.getcolnum("A"), val["班级名称"])
                wb.insertcell(
                    sheet1, linum, exceltools.getcolnum("B"), pair[0])
                wb.insertcell(
                    sheet1, linum, exceltools.getcolnum("C"), pair[1])
                wb.insertcell(
                    sheet1, linum, exceltools.getcolnum("E"), int(pair[2]))
                linum = linum + 1
            lend = linum - 1
            linum = linum + 1
            wb.insertcell(sheet1, linum, exceltools.getcolnum("A"), "未排课总课时")
            wb.insertformula(sheet1, linum, exceltools.getcolnum("D"),
                             "=sum(D"+str(lstart)+":D"+str(lend)+")")
            linum = linum + 1
            wb.insertcell(sheet1, linum, exceltools.getcolnum("A"), "已排课总课时")
            wb.insertformula(sheet1, linum, exceltools.getcolnum("E"),
                             "=sum(E"+str(lstart)+":E"+str(lend)+")")
            linum = linum + 1
            linum = linum + 1
        sheet2 = wb.insertsheet("教师课时统计")
        keylist = ["教师", "已安排课时"]
        wb.insertrow(sheet2, "A1", keylist)
        linum = 2
        for key, value in teacherdict.items():
            wb.insertcell(sheet2, linum, exceltools.getcolnum("A"), key)
            wb.insertcell(sheet2, linum, exceltools.getcolnum("B"), value)
            linum = linum + 1


datalist = []
data = parse(exceltools.readdata("keshi.xlsx", "专业班级分别", 1, 12, 'A', 'AQ'))
datalist = datalist + data
data = parse(exceltools.readdata("keshi.xlsx", "专业班级分别", 13, 21, 'A', 'BC'))
datalist = datalist + data
data = parse(exceltools.readdata("keshi.xlsx", "专业班级分别", 22, 25, 'A', 'AE'))
datalist = datalist + data
summary(datalist)
