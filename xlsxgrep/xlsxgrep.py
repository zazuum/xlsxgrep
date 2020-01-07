#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import argparse
import xlrd
import re
from collections import Counter
from pathlib import Path


def main():

    parser = argparse.ArgumentParser()

    parser.add_argument('pattern', help="Use PATTERN as the pattern to search for.",
                        type=str)
    parser.add_argument('path', help="file or folder location",
                        nargs="+", action="append")
    parser.add_argument("-i", '--ignore-case', help="Ignore case distinctions.",
                        required=False, action="store_true")
    parser.add_argument("-P", '--python-regex',
                        help="PATTERN is a Python regular expression.", required=False,
                        action="store_true")
    parser.add_argument("-w", '--word-regexp', help="Force PATTERN to match only whole words.",
                        required=False, action="store_true")
    parser.add_argument("-H", '--with-filename',
                        help="Print the file name for each match.", required=False,
                        action="store_true")
    parser.add_argument("-c", '--count', help="Print only a count of matches per file",
                        required=False, action="store_true")
    parser.add_argument("-N", '--with-sheetname', help="Print the sheet name for each match.",
                        required=False, action="store_true")
    parser.add_argument("-r", '--recursive', help="Search directories recursively.",
                        required=False, action="store_true")
    parser.add_argument("-V", '--version', help="Display version information and exit.", 
                        action='version', version="xlsxgrep  0.0.22")
    parser.add_argument('-sep', "--separator",
                        help="Define custom list separator for output, default is TAB", 
                        required=False, default="\t", type=str)
    
    args = parser.parse_args()



    ###### a bunch of variables #####

    fList = []
    query = args.pattern
    pythonRegexp = args.python_regex
    countMatches = []
    wordRegexp = args.word_regexp
    ignoreCase = args.ignore_case
    showFileAndSheetName = args.with_sheetname
    count = args.count
    filename = args.with_filename
    recursive = args.recursive
    matchFILES = []
    delimiter = args.separator
   


    def checkPythonRegex():
        if pythonRegexp == True:
            try:
                re.compile(query)
                is_valid = True # Used for debug test 
                pass
            except re.error:
                is_valid = False # Used for debug test
                exit("Error: Not valid Python Regular Expression")


    checkPythonRegex()



    def locationPath():
        for i in args.path[0]:

            if (Path(i).is_file() is False) and (Path(i).is_dir() is False):
                exit(str(i) + " File or folder not found. ")

            elif (Path(i).is_file() and str(Path(i)).endswith(('.xlsx', '.xls', '.XSLX', '.XLS'))):
                fList.append(str(Path(i)))

            elif Path(i).is_dir():
                if recursive == True:
                    for child in Path(i).rglob("*"):
                        if (str(child).endswith(('.xlsx', '.xls', '.XSLX', '.XLS'))):
                            fList.append(str(child))
                else:
                    for child in Path(i).iterdir():
                        if (str(child).endswith(('.xlsx', '.xls', '.XSLX', '.XLS'))):
                            fList.append(str(child))

            elif (Path(i).is_file() and str(Path(i)).endswith(('.xlsx', '.xls', '.XSLX', '.XLS'))) == False:
                # perform file check
                print("--> Error. Unsupported format, or corrupt file: ", Path(i))
                exit(0)

        search()



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



    def showFileNameAndSheet(file, active_sheet, linesArray):
        if count == True:
            pass

        elif filename == True:
            if showFileAndSheetName == True:
                return print(file + ": " + active_sheet.name + ': ' + str(delimiter) + str(delimiter).join(map(str, linesArray)))

            elif showFileAndSheetName == False:
                return print(file + ": " + str(delimiter) + str(delimiter).join(map(str, linesArray)))

        elif filename == False:
            if showFileAndSheetName == True:
                return print(active_sheet.name + ': ' + str(delimiter) + str(delimiter).join(map(str, linesArray)))
             
            else:
                print(*linesArray, sep=delimiter)



    def iterate(book, file):
        for eachSheet in range(0, len(book.sheet_names())):
            active_sheet = book.sheet_by_index(eachSheet)

            for i in range(0, active_sheet.nrows):
                cells = active_sheet.row_slice(
                    rowx=i, start_colx=0, end_colx=active_sheet.nrows)

                for cell in cells:
                    if checkArgs(cell.value):
                        countMatches.append(str(cell.value))

                        for col in range(0, active_sheet.ncols):
                            mLine = active_sheet.row_slice(
                                rowx=i, start_colx=0, end_colx=active_sheet.ncols)
                            linesArray = []

                            for col in mLine:
                                linesArray.append(col.value)
                        matchFILES.append(file)
                        showFileNameAndSheet(file, active_sheet, linesArray)

                    else:
                        pass


    def search():
        try:
            for file in fList:
                book = xlrd.open_workbook(file)
                iterate(book, file)

            if count == True:
                print("Total matches: ", len(countMatches))

                if showFileAndSheetName or filename == True:
                    for x in Counter(matchFILES):
                        d = Counter(matchFILES)
                        print(str(x) + ": " + str(d[x]))
                else:
                    pass
        except:
            print("Error. Unsupported format, or corrupt file: ", file)


    locationPath()




if __name__ == "__main__":
     main()
