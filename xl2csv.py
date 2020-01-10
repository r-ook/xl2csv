'''
Purpose: Convert Excel Sheet to CSV

Confirmed Supported file format: 
    .xlsx, .xls, .xlsm
    other .xl* formats may be supported but untested at this time

Usage: xlsx2csv [-h] [-s] [-d DIRECTORY] [-f] filename [sheetname]

Positional arguments:
  filename              File name of xls[x/m] to be converted
  sheetname             Sheet name to be converted (default: first indexed
                        sheet)

Optional arguments:
  -h, --help            show this help message and exit
  -s, --stdout          Output to stdout
  -d DIRECTORY, --directory DIRECTORY
                        Directory to save the output file
  -f, --force           Optional flag when directory is supplied. Create
                        directory(ies) if directory doesn't exist
'''

# This script needs some restructuring to support the additional argparse, hopefully to be done at a later time...
__author__ = 'Roca Koo'
__last_updated__ = 'Jan 10, 2020'

from os import path, mkdir
import csv
import sys
import xlrd
import argparse

def r_mkdir(pth):
    '''
    Recursive function to create sub-directories if not exist
    '''
    parent, _ = path.split(pth)
    if not path.exists(parent):
        r_mkdir(parent)
    if not path.exists(pth):
        mkdir(pth)
    return pth

def export_xl(fl, sheet: str=None, dir_: str=None, dir_force=False, use_stdout=False, overwrite=False):
    '''
    Main function to export the excel sheet as csv
    Arguments:
        sheet       [str]   - Name of the sheet, case sensitive.  If not provided, first indexed sheet will be used.
        dir_        [str]   - Directory name to save the output file.
                                If not provided, save to current directory
                                If provided, folder(s) must exist.  If not, dir_force should be used.
        dir_force   [bool]  - If dir_ is provided, to determine if path.exists will checked.
        use_stdout  [bool]  - Use stdout for piping instead of physical file.
        overwrite   [bool]  - By default, the physical file will always be saved as a new one.
                                If flagged, existing file with the same name will be overwritten.
    '''

    # Currently, to ensure exit code is respected, returning the actual error if occurred to raise at the end of the script.
    # This might be updated in the future to better respect the syntax.

    # validate file exists
    try:
        assert path.exists(fl)
    except AssertionError:
        return FileNotFoundError('{} not found'.format(fl))

    if dir_ and not dir_force:
        if not path.exists(dir_):
            return FileNotFoundError('{} does not exist.  If directory(s) is to be created, use -f or --force option.'.format(dir_))

    # validate file can be opened by xlrd module
    try:    
        wb = xlrd.open_workbook(fl)
    except xlrd.biffh.XLRDError:
        return NotImplementedError('{} extension is not supported'.format(path.splitext(fl)[-1]))

    if sheet:
        try:    # validate sheet exists
            sh = wb.sheet_by_name(sheet)
        except xlrd.biffh.XLRDError:
            wb.release_resources()
            return ValueError('"{}" not found in workbook (case sensitive)'.format(sheet))
    else:
        sh = wb.sheet_by_index(0)

    # Check if using stdout or spit out physical CSV
    if use_stdout:
        sys.stdout.writelines(
            '\n'.join(','.join(sh.row_values(r)) for r in range(sh.nrows)))
        wb.release_resources()
    else:
        if dir_: 
            fl = path.join(r_mkdir(dir_), path.basename(fl))

        new_file = get_new_name(fl, overwrite)
        with open(new_file, 'w', newline='') as f:
            c = csv.writer(f)
            for r in range(sh.nrows):
                c.writerow(sh.row_values(r))
            wb.release_resources()
            print('{} is created'.format(new_file))
    return None

def get_new_name(f, override=False):
    '''
    Fetch a unique file name from the name generater
    override=True will override existing file without warning
    '''
    new_file = csv_name(f)
    if not override:
        name = new_name_generator(new_file)
        while path.exists(new_file):    # if file exists create a new file
            new_file = next(name)
    return new_file

def new_name_generator(f):
    '''Generate a new unique file name with (n)'''
    i = 1
    while True:
        yield ' ({})'.format(i).join(path.splitext(f))
        i += 1

def csv_name(f):
    '''Change file ext to csv'''
    return path.splitext(f)[0] + '.csv'

def show_help():
    '''Show console help'''
    print(__doc__)
    sys.exit()

if __name__ == '__main__':
    helper_text = '''
Purpose: Convert Excel Sheet to CSV.

Confirmed Supported file format: 
    .xlsx, .xls, .xlsm
    other .xl* formats may be supported but untested at this time'''
    parser = argparse.ArgumentParser(
        prog='xlsx2csv',
        description=helper_text
        )
    parser.add_argument('filename', help='File name of xls[x/m] to be converted')
    parser.add_argument('sheetname', nargs='?', default=None, help='Sheet name to be converted (default: first indexed sheet).')
    parser.add_argument('-s', '--stdout', action='store_true', help='Output to stdout for piping.')
    parser.add_argument('-d', '--directory', help='Directory to save the output file.')
    parser.add_argument('-f', '--force', action='store_true', help="Optional flag when directory is supplied.  Create directory(ies) if directory doesn't exist.")
    parser.add_argument('-o', '--overwrite', action='store_true', help="Overwrite csv file if it already exists.  By default, a new unique csv file will be created each time.")

    # If no arguments are provided, default the display help text
    if len(sys.argv) == 1:
        sys.argv.append('-h')

    args = parser.parse_args()    
    error = export_xl(
        fl=args.filename,
        sheet=args.sheetname,
        dir_=args.directory,
        dir_force=args.force,
        use_stdout=args.stdout,
        overwrite=args.overwrite
        )
    if error is not None:
        #pylint: disable=raising-bad-type
        # Error will always be not None with the condition
        # and only error types are returned, so this error can be suppressed.
        raise error     
