import os
import re
from os import path

import openpyxl

import ColumnSelector
import SheetSelector


# File structure:
#   conf_dir--
#           sheet_name_a --
#                         column configurations
#                         A.txt
#                         B.txt
#                         ...
#           sheet_name_b
#           ...

# File format:
#   logic or every rows
#   Code:rep
#   Code:rep

# Language Specification
# CODE:rep
# 1. am:
# 2. fm:

class FileStructure:

    def __init__(self, conf_dir):
        if not os.path.isdir(conf_dir):
            raise Exception(conf_dir + "is not dir")

        self.__conf_dir = conf_dir
        self.__sheet_names = [x for x in os.listdir(conf_dir) if path.isdir(path.join(conf_dir, x))]
        selectors = {}
        for s in self.__sheet_names:
            ss = [x for x in os.listdir(path.join(conf_dir, s)) if x.endswith('.txt')]
            selectors[s] = ss
        self.__selectors = selectors

    @property
    def conf_dir(self):
        return self.__conf_dir

    @property
    def sheet_names(self):
        return self.__sheet_names

    # selectors: dict, sheet_name -> list of selector names
    @property
    def selectors(self):
        return self.__selectors


class FileFormat:
    def __init__(self, file_path, skip_wrong_lines = True):
        extractor = re.compile(r'^(?P<CODE>\w+):(?P<rep>.*)')
        self.__code_rep = []
        self.__skip_wrong_ling = skip_wrong_lines
        with open(file_path) as file:
            for line in file.readlines():
                m = extractor.match(line)
                if m is None:
                    if not (line.startswith('#') or line.startswith(';')):
                        if self.__skip_wrong_ling:
                            continue
                        err = "Wrong file format at " + file_path
                        raise Exception(err)
                    continue
                code, rep = m.group('CODE', 'rep')
                self.__code_rep.append((code, rep))

    @property
    def code_rep(self):
        return self.__code_rep


class SelectorFactory:
    __code_to_func = {
        'am': lambda s, rep: s.add_any_match(rep),
        'fm': lambda s, rep: s.add_fullmatch(rep),
        'cf': lambda s, rep: s.add_funcs(rep)
    }

    def get_selector(ff: FileFormat, selector=None):
        if selector is None:
            selector = ColumnSelector.Selector()
        if type(selector) is not ColumnSelector.Selector:
           raise Exception('The selector inputted is not a correct Selector object')
        for x in ff.code_rep:
           SelectorFactory.__code_to_func[x[0]](selector, x[1])
        return selector


class SelectorsLoading:
    def load_all_selectors(conf_dir: str, skip_wrong_lines=True):
        files_struct = FileStructure(conf_dir)
        sheet_selectors = {}
        extr = re.compile(r'(?P<code>.+)\.txt')
        for sheet_name in files_struct.sheet_names:
            col_select_files = files_struct.selectors[sheet_name]
            ss = SheetSelector.SheetSelector()
            for col_select_file in col_select_files:
                filepath = os.path.join(str(conf_dir), sheet_name, col_select_file)
                ff = FileFormat(filepath, skip_wrong_lines)
                # extract code from code file, like A.txt, C.txt
                code = extr.fullmatch(col_select_file).group('code')

                cs = ColumnSelector.ColumnSelector(code, SelectorFactory.get_selector(ff))
                ss.add_column_selector(cs)
            sheet_selectors[sheet_name] = ss
        return sheet_selectors


def test_selectors_loading():
    # ss = SelectorsLoading.load_all_selectors('assets/强电')
    conf_dir = 'assets/强电'
    fs = FileStructure(conf_dir)
    assert(fs.conf_dir == 'assets/强电')
    print('-------------')
    ff = FileFormat('assets/强电/电气计算1/C.txt')
    print(ff.code_rep)
    ff = FileFormat('assets/强电/电气计算1/B.txt')
    print(ff.code_rep)

    try:
        ff = FileFormat('assets/强电/电气计算1/A.txt', False)
    except Exception:
        assert(True)

    sel = SelectorFactory.get_selector(ff)
    sheet_selectors = SelectorsLoading.load_all_selectors(conf_dir)
    print(sheet_selectors)

    sheet_selector = sheet_selectors['电气计算1']
    workbook = openpyxl.load_workbook('assets/强电 - 副本07.xlsx')
    sheet = workbook['电气计算1']
    filtered_set = sheet_selector.filter(sheet)
    print(sorted(list(filtered_set)))
