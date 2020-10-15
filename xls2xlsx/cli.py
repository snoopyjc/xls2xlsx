"""Console script for xls2xlsx."""
import argparse
import sys
import os
from .xls2xlsx import XLS2XLSX

def main():
    """Console script for xls2xlsx."""
    parser = argparse.ArgumentParser(usage='xls2xlsx [-v] file.xls ... - converts and generates ./file.xlsx .... File may be a local file or a url.')
    parser.add_argument("-v", "--verbose", help="print the input and output filenames",
                    action="store_true")
    parser.add_argument('_', nargs='+')
    args = parser.parse_args()

    status = 0
    for arg in args._:
        try:
            x2x = XLS2XLSX(arg)
            filename = os.path.splitext(os.path.split(arg)[-1])[0]+'.xlsx'
            x2x.to_xlsx(filename=filename)
            if args.verbose:
                print(f'Converted {arg} to {filename}')
        except Exception as e:
            print(f'Exception converting {arg}: {e}: skipping!')
            status = 1

    return status


if __name__ == "__main__":
    sys.exit(main())  # pragma: no cover
