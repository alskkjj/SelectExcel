import argparse, os

from SelectActionXWhereYGroupByC import Works


class App:

    def _args_parser(self):
        parser = argparse.ArgumentParser(description='process a column of data in excel using specific rows')
        parser.add_argument('--conf_dir', nargs='?', help='configuration directory')
        parser.add_argument('--sheet_name', help='directory under conf_dir')
        parser.add_argument('--ops', nargs='+', choices=['sum', 'count'])
        parser.add_argument('-d')
        parser.add_argument('-g', nargs='+', help='The list of columns\' name')
        parser.add_argument('--file_name')
        parser.add_argument('-o', nargs='?', default=None)
        return parser

    def __init__(self, args_v=None):
        if args_v is None:
            parser = self._args_parser()
            args = vars(parser.parse_args())
            #args = vars(parser.parse_args())
            self._works = Works(args['conf_dir'], args['sheet_name'], args['ops'], args['file_name'],
                            args['g'], args['d'], args['o'])
        elif type(args_v) is list:
            parser = self._args_parser()
            args = vars(parser.parse_args(args_v))
            self._works = Works(args['conf_dir'], args['sheet_name'], args['ops'], args['file_name'],
                                args['g'], args['d'], args['o'])

    def start(self):
        self._works.do()
        del self._works


def test_app():
    app = App('--conf_dir=assets/强电 --sheet_name=电气计算 --ops sum count --file_name=assets/强电-副本07.xlsx -g D E H I -d=Q'.split())
    app.start()


def test_app1():
    app = App('--conf_dir=assets/强电 --sheet_name=电气计算 --ops sum count --file_name=assets/强电-副本07.xlsx -g D E H I -d=Q -o output.xlsx'.split())
    app.start()


def test_app2():
    app = App('--file_name=assets/强电-副本07.xlsx --sheet_name=电气计算 --ops sum count -g D E H I -d=Q -o output1.xlsx'.split())
    app.start()

