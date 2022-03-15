#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import argparse
import re
import warnings
import pyexcel as p
from collections import Counter
from pathlib import Path


def main():

    example_text = '''example:\n\txlsxgrep "PATTERN" -H -N -sep=";" -r /path/to/folder
                               \n'''
    parser = argparse.ArgumentParser(prog='xlsxgrep',
                                     epilog=example_text,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument('pattern', help="use PATTERN as the pattern to search for.",
                        type=str)
    parser.add_argument("-V", '--version', help="display version information and exit.",
                        action='version', version="xlsxgrep  0.0.26")
    parser.add_argument('path', help="file or folder location",
                        nargs="+", action="append")
    parser.add_argument("-i", '--ignore-case', help="ignore case distinctions.",
                        required=False, action="store_true")
    parser.add_argument("-P", '--python-regex',
                        help="PATTERN is a Python regular expression.", required=False,
                        action="store_true")
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
                        help="define custom list separator for output, default is TAB",
                        required=False, default="\t", type=str)

    if len(sys.argv) == 1:
        parser.print_usage(sys.stderr)
        print("Type 'xlsxgrep --help' for more information.")
        sys.exit(1)

    args = parser.parse_args()

##         a bunch of variables        ##
    # CLI Arguments
    query = args.pattern
    pythonRegexp = args.python_regex
    wordRegexp = args.word_regexp
    ignoreCase = args.ignore_case
    count = args.count
    filename = args.with_filename
    recursive = args.recursive
    delimiter = args.separator
    files_with_match = args.files_with_match
    files_without_match = args.files_without_match
    showFileAndSheetName = args.with_sheetname
    # Arrays
    countMatches = []
    matchFILES = []
    NONmatchFILES = []
    strMatches = []
    fList = []
    # Supress unsupported file extensions warnings
    Warnings = warnings.filterwarnings("ignore", category=UserWarning ) 


# Valid Python Regex Check ( Optional Argument -P, --python-regex)

    def checkPythonRegex():
        if pythonRegexp == True:
            try:
                re.compile(query)
                is_valid = True  # Used for debug test
                pass
            except re.error:
                is_valid = False  # Used for debug test
                exit("Error:    Not valid Python Regular Expression")

    checkPythonRegex()

# Checking file or folder format and destination

    def locationPath():
        fileTypes = ('.xlsx', '.xls', '.XSLX', '.XLS', '.ods', '.ODS')
        for i in args.path[0]:

            if (Path(i).is_file() is False) and (Path(i).is_dir() is False):
                exit(str(i) + " File or folder not found. ")

            elif (Path(i).is_file() and str(Path(i)).endswith(fileTypes)):
                fList.append(str(Path(i)))

            elif Path(i).is_dir():
                if recursive == True:
                    for child in Path(i).rglob("*"):
                        if (str(child).endswith(fileTypes)):
                            fList.append(str(child))
                else:
                    for child in Path(i).iterdir():
                        if (str(child).endswith(fileTypes)):
                            fList.append(str(child))

            elif (Path(i).is_file() and str(Path(i)).endswith(fileTypes)) == False:
                # perform file check
                print("Error:   Unsupported file format: ",
                      Path(i), file=sys.stderr)

        search()

# Checking pattern optional arguments ("-P", '--python-regex', "-w", '--word-regexp')

    def checkArgs(val):
        if pythonRegexp == True:
            return re.search(r'%s' % query, str(val))

        elif wordRegexp == True:
            if ignoreCase == True:
                return str(query).upper() == (str(val).upper())
            else:
                return query == (str(val))

        elif wordRegexp == False:
            if ignoreCase == False:
                return query == query in (str(val))
            else:
                return str(query).upper() in (str(val).upper())
        else:
            return print("...Some Error Occured...(optional arguments!?)")

# Checking output optional arguments ("-H", '--with-filename', "-N", '--with-sheetname')

    def showFileNameAndSheet(file, active_sheet, linesArray):
        if count == True:
            pass

        elif files_with_match:
            pass

        elif files_without_match:
            pass

        elif filename == True:
            if showFileAndSheetName == True:
                return print(file + ": " + active_sheet + ': ' + str(delimiter) + str(delimiter).join(map(str, linesArray)))

            elif showFileAndSheetName == False:
                return print(file + ": " + str(delimiter) + str(delimiter).join(map(str, linesArray)))

        elif filename == False:
            if showFileAndSheetName == True:
                return print(active_sheet + ': ' + str(delimiter) + str(delimiter).join(map(str, linesArray)))

            else:
                print(*linesArray, sep=delimiter)

# Iterate over rows and columns and append matches to array

    def iterateOverCells(book, file):
        NONmatchFILES.append(file)
        for key, item in book.items():
            for line in item:
                AuxFlag = False
                for cell in line:
                    if checkArgs(cell):
                        AuxFlag = True
                        countMatches.append(cell)
                        [strMatches.append(cell) for x in re.findall(str(query.upper()), str(cell).upper())]

                if AuxFlag == True:
                    matchFILES.append(file)
                    if file in NONmatchFILES:
                        NONmatchFILES.remove(file)
                    showFileNameAndSheet(file, key, line)

# Opening files, start searching

    def search():
        for file in fList:
            try:
                book = p.get_book_dict(file_name=file)
                iterateOverCells(book, file)
            
            except KeyboardInterrupt:
                # print('KeyboardInterrupt exception is caught')
                sys.exit(0)
            
            except:
                print(f"Error:\tUnsupported format, password protected or corrupted file: {file}", file=sys.stderr)
                pass
            
                   
        if count == True:
            print("Total matches: ", len(countMatches),
                  "Cells, ", len(strMatches), "Strings")

            if showFileAndSheetName or filename == True:
                for x in Counter(matchFILES):
                    d = Counter(matchFILES)
                    print(str(x) + ": " + str(d[x]) + " Rows")
            else:
                pass

        elif files_with_match:
            MYset = list(set(matchFILES))
            MYset.sort()
            [print(fx) for fx in MYset]

        elif files_without_match:
            MYset = list(set(NONmatchFILES))
            MYset.sort()
            [print(fx) for fx in MYset]

    locationPath()


if __name__ == "__main__":
    main()
