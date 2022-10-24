import xlwt
import xlrd


def delByName(col_num, sheet_num, source, target, d):
    result = xlwt.Workbook(encoding="ascii")
    whsheet = result.add_sheet("a")
    y = 0
    try:
        workbook = xlrd.open_workbook(source)
        sheet = workbook.sheet_by_index(sheet_num)
        nrowssum = sheet.nrows
        for i in range(0, nrowssum):
            is_valid = 1
            data = sheet.row(i)
            aaa = str(data[col_num])
            for k in range(0, len(d)):
                keyword = str(d[k])
                if aaa.find(keyword) > 0:
                    is_valid = 0
            if is_valid:
                y = y + 1
                for j in range(len(data)):
                    whsheet.write(y, j, sheet.cell_value(i, j))
        result.save(target)
    except Exception as e:
        print(e)
    result.save(target)


def selByList(col_num, sheet_num, source, target, d):
    jieguo = xlwt.Workbook(encoding="ascii")
    wsheet = jieguo.add_sheet("a")
    y = 0
    try:
        workbook = xlrd.open_workbook(source)
        sheet = workbook.sheet_by_index(sheet_num)
        nrowssum = sheet.nrows
        for i in range(0, nrowssum):
            data = sheet.row(i)
            aaa = str(data[col_num])
            for k in range(0, len(d)):
                keyword = str(d[k])
                if aaa.find(keyword) > 0:
                    y = y + 1
                    for j in range(len(data)):
                        wsheet.write(y, j, sheet.cell_value(i, j))
        jieguo.save(target)
    except Exception as e:
        print(e)
    jieguo.save(target)


if __name__ == '__main__':
    del_word = ["生源", "户籍"]
    key_word = ["计算机", "电子", "软件", "信息", "工学", "工科"]
    delByName(22, 0, "sheet4_1.xls", "sheet4.xls", del_word)
    # selByList(15, 0, "sheet4.xls", "sheet4_1.xls", del_word)
