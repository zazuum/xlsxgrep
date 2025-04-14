#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import argparse
import re
import warnings
import logging
import pyexcel as p
from pathlib import Path

__version__ = '0.0.30'


def main():

    example_text = '''examples:\n  xlsxgrep -i "foo" foobar.xlsx\n  xlsxgrep -c -H "(?i)foo|bar" /folder
			\n'''
    parser = argparse.ArgumentParser(prog='xlsxgrep',
                                     epilog=example_text,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument('-h','--help', action="help", help="show this help message and exit.")
    parser.add_argument('pattern', help="use PATTERN as the pattern to search for.",
                        type=str)
    parser.add_argument("-V", '--version', help="display version information and exit.",
                        action='version', version='%(prog)s ' + str(__version__))
    parser.add_argument('path', help="file or folder location",
                        nargs="+", action="append")
    parser.add_argument("-P", '--python-regex',
                        help="PATTERN is a Python regular expression. This is the default.", required=False,
                        action="store_true", default=False)
    parser.add_argument("-F", '--fixed-strings',
                        help="interpret PATTERN as fixed strings, not regular expressions.", required=False,
                        action="store_true", default=False)
    parser.add_argument("-i", '--ignore-case', help="ignore case distinctions.",
                        required=False, action="store_true")
    parser.add_argument("-w", '--word-regexp', help="force PATTERN to match only whole words.",
                        required=False, action="store_true")
    parser.add_argument("-c", '--count', help="print only a count of matches per file.",
                        required=False, action="store_true")
    parser.add_argument("-r", '--recursive', help="search directories recursively.",
                        required=False, action="store_true")
    parser.add_argument("-H", '--with-filename',
                        help="print the file name for each match.", required=False,
                        action="store_true")
    parser.add_argument("-N", '--with-sheetname', help="print the sheet name for each match.",
                        required=False, action="store_true")
    parser.add_argument("-l", '--files-with-match', help="print only names of FILEs with match pattern.",
                        required=False, action="store_true")
    parser.add_argument("-L", '--files-without-match', help="print only names of FILEs with no match pattern.",
                        required=False, action="store_true")
    parser.add_argument('-S', "--separator",
                        help="define custom list separator for output, the default is TAB",
                        required=False, default="\t", type=str)
    parser.add_argument("-Z", '--null', help="output a zero byte (the ASCII NUL character) instead of the usual newline.",
                        required=False, action="store_true")

    if len(sys.argv) == 1:
        parser.print_usage(sys.stderr)
        print("Type 'xlsxgrep --help' for more information.")
        sys.exit(1)

    args = parser.parse_args()

    ### CLI Arguments ###
    Query = args.pattern
    PythonRegex = args.python_regex
    FixedString = args.fixed_strings
    WordRegexp = args.word_regexp
    IgnoreCase = args.ignore_case
    Count = args.count
    Filename = args.with_filename
    Recursive = args.recursive
    Delimiter = str(args.separator)
    Files_with_match = args.files_with_match
    Files_without_match = args.files_without_match
    ShowFileAndSheetName = args.with_sheetname
    ZeroByte = args.null


    def DisableWarningsAndLoggings():
        ## - Supress unsupported file extensions warnings.
        ##   'UserWarning: Data Validation extension is not supported and will be removed'. (module=openpyxl)
        ##   'UserWarning: Unknown extension is not supported and will be removed'.         (module=openpyxl)
        warnings.filterwarnings("ignore", category=UserWarning, message="Unknown extension is not supported and will be removed")
        warnings.filterwarnings("ignore", category=UserWarning, message='Data Validation extension is not supported and will be removed')
        ## - Ignore deprecated python regex warnings.
        ##   'DeprecationWarning: Flags not at the start of the expression 'foo|(?i)bar'.   (module=re)
        warnings.filterwarnings("ignore", category=DeprecationWarning,message='.*Flags not at the start of the expression*.')
        ## - Supress Conditional Formatting extension not supported and Cannot parse header or footer warning.
        warnings.filterwarnings("ignore", category=UserWarning,message='Conditional Formatting extension is not supported and will be removed')
        warnings.filterwarnings("ignore", category=UserWarning, message="Cannot parse header or footer so it will be ignored")
        ## - Disable all warnings in openpyxl
        #warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
        ## - Disable all logging warnings
        logging.disable(logging.WARNING)

    DisableWarningsAndLoggings()

# Valid Python Regex Check ( Optional Argument -P, --python-regex)

    def Check_Python_Regex():
        if (FixedString or IgnoreCase or WordRegexp):
            if args.python_regex == True:
                sys.exit("xlsxgrep: --python-regex cannot be used together with: -F, -w or -i")
            else:
                args.python_regex = False
                return args.python_regex

        else:
            try:
                args.python_regex = True
                re.compile(Query)
                is_valid = True  # Used for debug test
                pass
            except re.error:
                is_valid = False  # Used for debug test
                exit("Error:  Not valid Python Regular Expression. For fixed strings use flag: -F")

    Check_Python_Regex()

# Checking file or folder format and destination

    def File_And_Path_Location():
        File_List = [ ]
        fileTypes = ('.xls', '.XLS','.xlsx', '.XLSX', '.ods', '.ODS',
                     '.csv', '.CSV', '.tsv', '.TSV', '.xlsm', '.XLSM')
        for i in args.path[0]:

            if (Path(i).is_file() is False) and (Path(i).is_dir() is False):
                exit(str(i) + " File or folder not found. ")

            elif (Path(i).is_file() and str(Path(i)).endswith(fileTypes)):
                File_List.append(str(Path(i)))

            elif Path(i).is_dir():
                if Recursive == True:
                    for child in Path(i).rglob("*"):
                        if (str(child).endswith(fileTypes)):
                            File_List.append(str(child))
                else:
                    for child in Path(i).iterdir():
                        if (str(child).endswith(fileTypes)):
                            File_List.append(str(child))

            elif (Path(i).is_file() and str(Path(i)).endswith(fileTypes)) == False:
                # perform file check
                print("Error:   Unsupported file format: ",
                      Path(i), file=sys.stderr)

        SEARCH(File_List)

# Checking pattern optional arguments ("-P", '--python-regex', "-w", '--word-regexp')

    def Check_Optional_Args(val):

        if args.python_regex == True:
            return re.search(r'%s' % Query, str(val))

        elif WordRegexp == True:
            if IgnoreCase == True:
                return str(Query).upper() == (str(val).upper())
            else:
                return Query == (str(val))

        elif WordRegexp == False:
            if IgnoreCase == False:
                return Query == Query in (str(val))
            else:
                return str(Query).upper() in (str(val).upper())
        else:
            return print("...Some Error Occured...(optional arguments!?)")

# Checking output optional arguments ("-H", '--with-filename', "-N", '--with-sheetname')

    def Show_Filename_And_Sheetname(file, active_sheet, linesArray):
        ENDSWITH="\n"
        if ZeroByte:
            ENDSWITH =""

        if Count == True:
            pass

        elif Files_with_match:
            pass

        elif Files_without_match:
            pass

        elif Filename == True:
            if ShowFileAndSheetName == True:
                return print(file + ": " + active_sheet + ': ' + Delimiter + Delimiter.join(map(str, linesArray)),end=ENDSWITH)

            elif ShowFileAndSheetName == False:
                return print(file + ": " + Delimiter + Delimiter.join(map(str, linesArray)),end=ENDSWITH)

        elif Filename == False:
            if ShowFileAndSheetName == True:
                return print(active_sheet + ': ' + Delimiter + Delimiter.join(map(str, linesArray)),end=ENDSWITH)

            else:
                print(*linesArray, sep=Delimiter,end=ENDSWITH)


# Iterate over rows and columns and append matches count to array.

    SumOfROW , SumOfCELL , SumOfSTR = [] , [] , []
    def Iterate_Over_Cells(book, file):
        ROWcount, CELLcount, STRcount = [0] , [0] ,[0]
        for key, item in book.items():
             for line in item:
                AuxFlag = False
                for cell in line:
                 if Check_Optional_Args(cell):
                    if Count:
                        AuxFlag = True
                        CELLcount[0] = CELLcount[0] + 1
                        reESCapedQuery = re.escape( str(Query).upper() )
                        STRcell = str(cell).upper()
                        if args.python_regex == False:
                         for x in re.findall(reESCapedQuery , STRcell):
                            STRcount[0] = STRcount[0] + 1

                        else:
                            for x in re.findall(str(Query), str(cell)):
                                STRcount[0] = STRcount[0] + 1
                    else:
                        AuxFlag = True
                        ROWcount[0] = ROWcount[0] - 1

                if AuxFlag == True:
                    ROWcount[0] = ROWcount[0] + 1

                    Show_Filename_And_Sheetname(file, key, line)


        if ROWcount[0] > 0:
          ENDSWITH="\n"
          ROWS, CELLS, STRINGS = ROWcount , CELLcount, STRcount
          if args.null:
            ENDSWITH =""
          if ShowFileAndSheetName or Filename:
            print(file,":",ROWS[0] , "Rows, ",CELLS[0], "Cells, ", STRINGS[0], "Strings" ,end=ENDSWITH )
          SumOfCELL.extend(CELLcount)
          SumOfSTR.extend(STRcount)
          SumOfROW.extend(ROWcount)



    # Check files-with-match and files-without-match arguments

    def HyphenlAndHyphenLCheck(book,file):
        ENDSWITH="\n"
        if args.null:
            ENDSWITH =""

        if Files_with_match:
            if Check_Optional_Args(book):
                 return print(file, end=ENDSWITH)

        elif Files_without_match:
            if not Check_Optional_Args(book):
                return print(file, end=ENDSWITH)

        else:

            Iterate_Over_Cells(book,file)


# Count matches. Rows, cells and strings.

    def SumOfRowsCellsAndStrings():
         ROWS, CELLS, STRINGS = sum(SumOfROW) , sum(SumOfCELL), sum(SumOfSTR)
         print("Search results: ", ROWS , "Rows, ",CELLS, "Cells, ", STRINGS, "Strings"  )


# Opening files, start searching

    def SEARCH(File_List):
        for file in File_List:
            try:
                if file.endswith((".xlsx",".XLSX",".xlsm",".XLSM")):
                    book = p.get_book_dict(file_name=file, skip_hidden_row_and_column=False)

                else:
                    book = p.get_book_dict(file_name=file,)


                HyphenlAndHyphenLCheck(book, file)


            except KeyboardInterrupt:

                sys.exit(0)

            except:
                print(f"Error:\tUnsupported format, password protected or corrupted file: {file}", file=sys.stderr)
                pass


        if Count:
            if Files_with_match or Files_without_match:
                pass

            else:
                SumOfRowsCellsAndStrings()

    File_And_Path_Location()


if __name__ == "__main__":
    main()


