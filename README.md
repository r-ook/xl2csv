#xl2csv

Simple CLI using `xlrd` package to convert Excel Sheets to CSV files

#Purpose:

Convert Excel Sheet to CSV

#Confirmed Supported file format: 

`.xlsx`, `.xls`, `.xlsm`
other `.xl*` formats may be supported but untested at this time

#Usage:

    xlsx2csv [-h] [-s] [-d DIRECTORY] [-f] filename [sheetname]`

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
