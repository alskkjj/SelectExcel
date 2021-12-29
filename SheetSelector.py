import ColumnSelector


class SheetSelector:
    def __init__(self, column_selectors: set = None):
        if column_selectors is None:
            column_selectors = set()

        self.__column_selectors = column_selectors

    def filter(self, sheet):
        filtered_set = set()
        for selector in self.__column_selectors:
            res = selector.single_col_filter(sheet)
            if len(filtered_set) == 0:
                filtered_set = res
            else:
                filtered_set &= res
        return filtered_set

    # set
    @property
    def column_selectors(self):
        return self.__column_selectors

    def add_column_selector(self, column_selector: ColumnSelector.ColumnSelector):
        self.__column_selectors.add(column_selector)
