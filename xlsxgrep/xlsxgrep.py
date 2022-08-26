#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import argparse
import re
import warnings
import pyexcel as p
from collections import Counter
from pathlib import Path

__version__ = '0.0.28'


def main():

    example_text = '''examples:\n  xlsxgrep -i "foo" foobar.xlsx\n  xlsxgrep -c -H "(?i)foo|bar" /folder 
			\n'''
    parser = argparse.ArgumentParser(prog='xlsxgrep',
                                     epilog=example_text,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
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
    parser.add_argument('-sep', "--separator",
                        help="define custom list separator for output, the default is TAB",
                        required=False, default="\t", type=str)

    if len(sys.argv) == 1:
        parser.print_usage(sys.stderr)
        print("Type 'xlsxgrep --help' for more information.")
        sys.exit(1)

    args = parser.parse_args()

##         a bunch of variables        ##
    ## CLI Arguments
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
    ## Arrays
    CountMatchesArr, FilesWithMatch, FilesWithoutMatch, StrMatches, File_List = [ ],[ ] ,[ ],[ ],[ ]


    # - Supress unsupported file extensions warnings.
    #   'UserWarning: Data Validation extension is not supported and will be removed'. (module=openpyxl)
    #   'UserWarning: Unknown extension is not supported and will be removed'.         (module=openpyxl)
    warnings.filterwarnings("ignore", category=UserWarning, message="Unknown extension is not supported and will be removed")
    warnings.filterwarnings("ignore", category=UserWarning, message='Data Validation extension is not supported and will be removed')
    # - Ignore deprecated python regex warnings.
    #   'DeprecationWarning: Flags not at the start of the expression 'foo|(?i)bar'.   (module=re)
    warnings.filterwarnings("ignore", category=DeprecationWarning,message='.*Flags not at the start of the expression*.')



# Valid Python Regex Check ( Optional Argument -P, --python-regex)

    def Check_Python_Regex():
        if (FixedString or IgnoreCase or WordRegexp):
            if args.python_regex == True:
                sys.exit("xlsxgrep: --python-regex cannot be used together with: -F, -w or -i")
            else:
                PythonRegex = False
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
        fileTypes = ('.xls', '.XLS','.xlsx', '.XLSX', '.ods', '.ODS',
                     '.csv', '.CSV', '.tsv', '.TSV')
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

        SEARCH()

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
        if Count == True:
            pass

        elif Files_with_match:
            pass

        elif Files_without_match:
            pass

        elif Filename == True:
            if ShowFileAndSheetName == True:
                return print(file + ": " + active_sheet + ': ' + Delimiter + Delimiter.join(map(str, linesArray)))

            elif ShowFileAndSheetName == False:
                return print(file + ": " + Delimiter + Delimiter.join(map(str, linesArray)))

        elif Filename == False:
            if ShowFileAndSheetName == True:
                return print(active_sheet + ': ' + Delimiter + Delimiter.join(map(str, linesArray)))

            else:
                print(*linesArray, sep=Delimiter)

# Iterate over rows and columns and append matches to array

    def Iterate_Over_Cells(book, file):
        FilesWithoutMatch.append(file)
        for key, item in book.items():
            for line in item:
                AuxFlag = False
                for cell in line:
                    if Check_Optional_Args(cell):
                        AuxFlag = True
                        CountMatchesArr.append(cell)
                        reESCapedQuery = re.escape( str(Query).upper() )
                        STRcell = str(cell).upper()
                        if args.python_regex == False:
                            [StrMatches.append(cell) for x in re.findall(reESCapedQuery , STRcell)]
                        else:
                            [StrMatches.append(cell) for x in re.findall(str(Query), str(cell))]

                if AuxFlag == True:
                    FilesWithMatch.append(file)
                    if file in FilesWithoutMatch:
                        FilesWithoutMatch.remove(file)
                    Show_Filename_And_Sheetname(file, key, line)

# Count matches. Rows, cells and strings.

    def Count_Matches():
            if ShowFileAndSheetName or Filename == True:
                for x in Counter(FilesWithMatch):
                    d = Counter(FilesWithMatch)
                    print(str(x) + ": " + str(d[x]) + " Rows")
            else:
                pass

            ROWS, CELLS, STRINGS = len(FilesWithMatch) , len(CountMatchesArr) ,len(StrMatches)
            return print("Search results: ", ROWS , "Rows, ",CELLS, "Cells, ", STRINGS, "Strings"  )

# Opening files, start searching

    def SEARCH():
        for file in File_List:
            try:
                book = p.get_book_dict(file_name=file)
                Iterate_Over_Cells(book, file)

            except KeyboardInterrupt:
                # print('KeyboardInterrupt exception is caught')
                sys.exit(0)

            except:
                print(f"Error:\tUnsupported format, password protected or corrupted file: {file}", file=sys.stderr)
                pass


        if Count == True:
            Count_Matches()

        elif Files_with_match:
            MYset = list(set(FilesWithMatch))
            MYset.sort()
            return [print(fx) for fx in MYset]

        elif Files_without_match:
            MYset = list(set(FilesWithoutMatch))
            MYset.sort()
            return [print(fx) for fx in MYset]

    File_And_Path_Location()


if __name__ == "__main__":
    main()
