import openpyxl

import misc
from SelectorLoader import SelectorsLoading
from GroupByColumn import group_by_columns
from DoToColumn import sum_to_a_column_of_cells


def _fold_left_ops(sheet, applying_col, rowset, ops):
    accumulated_items = [None for i in range(len(ops))]
    for row in rowset:
        idx = 0
        for op in ops:
            accumulated = accumulated_items[idx]
            n = sheet[misc.pyxl_xy(row, applying_col)].value
            accumulated = op(accumulated, n)
            accumulated_items[idx] = accumulated
            idx += 1
    return tuple(accumulated_items)


def _sum(accumulated, n, add_op=None):
    import sys
    if accumulated is None:
        accumulated = 0.0
    if add_op is None:
        float_n = 0.0
        try:
            float_n = float(n)
        except ValueError:
            #raise ValueError('Not corrected format. do not omit')
            print('Not corrected format. do not omit. value is ' + str(n), file=sys.stderr)
        except TypeError:
            #raise TypeError('Not corrected number format. do not omit')
            print('Not corrected number format. do not omit. value is None', file=sys.stderr)

        accumulated += float_n

        return accumulated
    else:
        accumulated = add_op(accumulated, n)
        return accumulated


def _count(accumulated, n, add_op=None):
    if accumulated is None:
        accumulated = 0.0
    if add_op is None:
        accumulated += 1
        return accumulated
    else:
        accumulated = add_op(accumulated, n)
        return accumulated


def __sum(sheet, group_cols, sum_col, rowset):
    groups = group_by_columns(sheet, group_cols, rowset)
    sum_pairs = []
    for left in groups:
        filterted_rowset = set(rowset) & set(groups[left])
        li, ns = sum_to_a_column_of_cells(sheet, sum_col, filterted_rowset)
        sum_pairs.append((left, ns))
    return sum_pairs


def _get_rows_set(sheet, conf_dir=None, sheet_name=None):
    if conf_dir is None:
        rowset = range(1, sheet.max_row + 1)
        return rowset
    else:
        selectors = SelectorsLoading.load_all_selectors(conf_dir)
        selector = selectors[sheet_name]
        rowset = selector.filter(sheet)
        return rowset


class Works:
    _op_funcs = {
        'sum': _sum,
        'count': _count
    }

    def __init__(self, conf_dir, sheet_name, ops, file_name, group_cols, dest_col, xl=0):
        self._conf_dir = conf_dir
        self._sheet_name = sheet_name
        self._ops = ops
        self._file_name = file_name

        self._group_cols = group_cols
        self._dest_col = dest_col
        self._xl = xl
        self._workbook = None

    def __del__(self):
        if self._workbook is not None:
            self._workbook.close()
            del self._workbook

    def do(self):
        workbook = openpyxl.load_workbook(self._file_name, data_only=True)
        self._workbook = workbook
        sheet = workbook[self._sheet_name]
        rowset = _get_rows_set(sheet, self._conf_dir, self._sheet_name)
        operators = [Works._op_funcs[x] for x in self._ops]
        groups = group_by_columns(sheet, self._group_cols, rowset)
        res = []
        for x in groups:
            dest_tuple = _fold_left_ops(sheet, self._dest_col, groups[x], operators)
            tp = misc.tuple_canceler( (x, '' ,dest_tuple) )
            res.append(tp)

        if self._xl is None:
            for x in res:
                print(x)
        else:
            misc.write_to_xl(res, self._xl)
