import math


# (1, 2) to 'B1'
# 1 B
# 0 1
# (1, 1) to 'A1'
import openpyxl


def _tuple_canceler(tp):
    lst = []
    if type(tp) is tuple:
        for x in tp:
            lst.extend(tuple_canceler(x))
    else:
        if type(tp) is list:
            lst.extend(tp)
        else:
            lst.append(tp)

    return lst


def tuple_canceler(tp):
    return tuple(_tuple_canceler(tp))


def pyxl_xy(row: int, col):
    if type(col) is int:
        # temporary from A-Z
        A2Z = [str(chr(x)) for x in range(0x41, 0x5A + 1)]
        if col < 0 or col > len(A2Z):
            raise "Col out of range"
        if row < 0:
            raise "Row out of range"
        COL = str(A2Z[col-1])
        ROW = str(row)
        return COL + ROW
    elif type(col) is str:
        if row < 0:
            raise "Row out of range"
        return col + str(row)


def sheet_range(num):
    return range(1, num)


def float_eq(a, b):
    return math.fabs(a - b) < 10e-7


def write_to_xl(dct, xlname):
    workbook = openpyxl.Workbook()
    sheet = workbook.create_sheet('汇总结果', 0)

    if type(dct) is dict:
        line_contents = ( tuple_canceler( (x, dct[x] ) for x in dct.keys ) )
        for x in line_contents:
            sheet.append(x)

        workbook.save(xlname)
    elif type(dct) is list:
        lst = dct

        for x in lst:
            sheet.append(list(tuple_canceler(x)))

        workbook.save(xlname)
