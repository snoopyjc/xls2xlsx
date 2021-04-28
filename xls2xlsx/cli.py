"""Console script for xls2xlsx."""
import argparse
import sys
import os
from .xls2xlsx import XLS2XLSX
from xlrd.compdoc import CompDocError
def main():
    """Console script for xls2xlsx."""
    parser = argparse.ArgumentParser(usage='xls2xlsx [-v] file.xls ... - converts and generates ./file.xlsx .... File may be a local file or a url.')
    parser.add_argument("-v", "--verbose", help="print the input and output filenames",
                    action="store_true")
    parser.add_argument('_', nargs='+')
    parser.add_argument("-i","--ignore",help="ignore workbook corruption",
                    action="store_true")
    args = parser.parse_args()
    

    status = 0
    for arg in args._:
        try:
            ignore_w_c = False
            if args.ignore:
                ignore_w_c = True
            x2x = XLS2XLSX(arg,ignore_workbook_corruption=ignore_w_c)
            filename = os.path.splitext(os.path.split(arg)[-1])[0]+'.xlsx'
            x2x.to_xlsx(filename=filename)
            if args.verbose:
                print(f'Converted {arg} to {filename}')
        except CompDocError as e:
            print(f'convert failed:{e}\n you can try to use -i or --ignore to ignore this error')
            status = 1
        except Exception as e:
            print(f'Exception converting {arg}: {e}: skipping!')
            status = 1


    return status


if __name__ == "__main__":
    sys.exit(main())  # pragma: no cover
