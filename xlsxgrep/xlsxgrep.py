#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import argparse
import re
import warnings
import logging
import pyexcel as p
from pathlib import Path
import locale
from textwrap import dedent

__license__ = "MIT"
__version__ = "0.0.31"
__author__ = "Ivan Cvitic"
__email__ = "cviticivan@gmail.com"
VERSION_INFO = [
    "xlsxgrep version: {0}".format(__version__),
    "Python version: {0}".format(
        " ".join(line.strip() for line in sys.version.splitlines())
    ),
    "Locale: {0}".format(".".join(str(s) for s in locale.getlocale())),
]


def main():

    help_text = """positional arguments:
  PATTERN                    use PATTERN as the pattern to search for.
  FILE                       file or path to folder

options:
  -h, --help                 show this help message and exit.
  -V, --version              display version information and exit.
  -P, --python-regex         PATTERN is a Python regular expression. This is the default.
  -F, --fixed-strings        interpret PATTERN as fixed strings, not regular expressions.
  -i, --ignore-case          ignore case distinctions.
  -w, --word-regexp          force PATTERN to match only whole words.
  -c, --count                print only a count of matches per file.
  -r, --recursive            search directories recursively.
  -H, --with-filename        print the file name for each match.
  -N, --with-sheetname       print the sheet name for each match.
  -l, --files-with-match     print only names of FILEs with match pattern.
  -L, --files-without-match  print only names of FILEs with no match pattern.
  -S, --separator SEPARATOR  define custom list separator for output, the default is TAB.
  -Z, --null                 output a zero byte (the ASCII NUL character) instead of the 
                             usual newline.

examples:
    xlsxgrep -i "foo" foobar.xlsx
    xlsxgrep -c -H "(?i)foo|bar" /folder"""
    parser = argparse.ArgumentParser(
        add_help=False,  # epilog=example_text,
        description=dedent(help_text),
        prog="xlsxgrep",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        usage=dedent(
            """
	    xlsxgrep [-h] [-V] [-P] [-F] [-i] [-w] [-c] [-r] [-H] [-N] [-l] [-L] [-S SEPARATOR] 
                [-Z] [-d] PATTTERN FILE [FILE ...]


            """
        ).strip(),
    )
    parser.add_argument(
        "-h", "--help", action="help", help=argparse.SUPPRESS
    )
    parser.add_argument(
        "PATTERN", help=argparse.SUPPRESS, type=str
    )
    parser.add_argument(
        "-V",
        "--version",
        help=argparse.SUPPRESS,
        # help="display version information and exit.",
        action="version",
        version=dedent("\n".join(VERSION_INFO) + "\n"),
    )
    parser.add_argument(
        "FILE", help=argparse.SUPPRESS, nargs="+", action="append",
    )
    parser.add_argument(
        "-P",
        "--python-regex",
        help=argparse.SUPPRESS,
        # help="PATTERN is a Python regular expression. This is the default.",
        required=False,
        action="store_true",
        default=False,
    )
    parser.add_argument(
        "-F",
        "--fixed-strings",
        help=argparse.SUPPRESS,
        # help="interpret PATTERN as fixed strings, not regular expressions.",
        required=False,
        action="store_true",
        default=False,
    )
    parser.add_argument(
        "-i",
        "--ignore-case",
        help=argparse.SUPPRESS,
        # help="ignore case distinctions.",
        required=False,
        action="store_true",
    )
    parser.add_argument(
        "-w",
        "--word-regexp",
        help=argparse.SUPPRESS,
        # help="force PATTERN to match only whole words.",
        required=False,
        action="store_true",
    )
    parser.add_argument(
        "-c",
        "--count",
        help=argparse.SUPPRESS,
        # help="print only a count of matches per file.",
        required=False,
        action="store_true",
    )
    parser.add_argument(
        "-r",
        "--recursive",
        help=argparse.SUPPRESS,
        # help="search directories recursively.",
        required=False,
        action="store_true",
    )
    parser.add_argument(
        "-H",
        "--with-filename",
        help=argparse.SUPPRESS,
        # help="print the file name for each match.",
        required=False,
        action="store_true",
    )
    parser.add_argument(
        "-N",
        "--with-sheetname",
        help=argparse.SUPPRESS,
        # help="print the sheet name for each match.",
        required=False,
        action="store_true",
    )
    parser.add_argument(
        "-l",
        "--files-with-match",
        help=argparse.SUPPRESS,
        # help="print only names of FILEs with match pattern.",
        required=False,
        action="store_true",
    )
    parser.add_argument(
        "-L",
        "--files-without-match",
        help=argparse.SUPPRESS,
        # help="print only names of FILEs with no match pattern.",
        required=False,
        action="store_true",
    )
    parser.add_argument(
        "-S",
        "--separator",
        # help="define custom list separator for output, the default is TAB.",
        help=argparse.SUPPRESS,
        required=False,
        default="\t",
        type=str,
    )
    parser.add_argument(
        "-Z",
        "--null",
        help=argparse.SUPPRESS,
        # help="output a zero byte (the ASCII NUL character) instead of the usual newline.",
        required=False,
        action="store_true",
    )
    parser.add_argument(
        "-d"
        "--debug",
        help=argparse.SUPPRESS,
        required=False,
        default=False,
        action="store_true",
    )

    if len(sys.argv) == 1:
        parser.print_usage(sys.stderr)
        print("Type 'xlsxgrep --help' for more information.")
        sys.exit(1)

    args = parser.parse_args()

    def ActivateDebug():
        if args.debug == False:
            # Some debug options
            # - Supress unsupported file extensions warnings.
            # 'UserWarning: Data Validation extension is not supported and will be removed'. (module=openpyxl)
            # 'UserWarning: Unknown extension is not supported and will be removed'.         (module=openpyxl)
            warnings.filterwarnings(
                "ignore",
                category=UserWarning,
                message="Unknown extension is not supported and will be removed",
            )
            warnings.filterwarnings(
                "ignore",
                category=UserWarning,
                message="Data Validation extension is not supported and will be removed",
            )
            # - Ignore deprecated python regex warnings.
            # 'DeprecationWarning: Flags not at the start of the expression 'foo|(?i)bar'.   (module=re)
            warnings.filterwarnings(
                "ignore",
                category=DeprecationWarning,
                message=".*Flags not at the start of the expression*.",
            )
            # - Supress Conditional Formatting extension not supported and Cannot parse header or footer warning.
            warnings.filterwarnings(
                "ignore",
                category=UserWarning,
                message="Conditional Formatting extension is not supported and will be removed",
            )
            warnings.filterwarnings(
                "ignore",
                category=UserWarning,
                message="Cannot parse header or footer so it will be ignored",
            )
            # - Disable all warnings in openpyxl
            warnings.filterwarnings(
                "ignore", category=UserWarning, module="openpyxl")
            # - Disable all logging warnings
            logging.disable(logging.WARNING)
        else:
            print("--version info: "+" ".join(VERSION_INFO))

            pass

    ActivateDebug()

    # Valid Python Regex Check ( Optional Argument -P, --python-regex)

    def Check_Python_Regex():
        if args.fixed_strings or args.ignore_case or args.word_regexp:
            if args.python_regex == True:
                sys.exit(
                    "xlsxgrep: --python-regex cannot be used together with: -F, -w or -i"
                )
            else:
                args.python_regex = False
                return args.python_regex

        else:
            try:
                args.python_regex = True
                re.compile(args.PATTERN)
                pass
            except re.error:
                exit(
                    "Error:  Not valid Python Regular Expression. For fixed strings use flag: -F"
                )

    Check_Python_Regex()

    # Checking file or folder format and destination

    def File_And_Path_Location():
        File_List = []
        fileTypes = (
            ".xls",
            ".XLS",
            ".xlsx",
            ".XLSX",
            ".ods",
            ".ODS",
            ".csv",
            ".CSV",
            ".tsv",
            ".TSV",
            ".xlsm",
            ".XLSM",
        )
        for i in args.FILE[0]:

            if (Path(i).is_file() is False) and (Path(i).is_dir() is False):
                exit(str(i) + " File or folder not found. ")

            elif Path(i).is_file() and str(Path(i)).endswith(fileTypes):
                File_List.append(str(Path(i)))

            elif Path(i).is_dir():
                if args.recursive == True:
                    for child in Path(i).rglob("*"):
                        if str(child).endswith(fileTypes):
                            File_List.append(str(child))
                else:
                    for child in Path(i).iterdir():
                        if str(child).endswith(fileTypes):
                            File_List.append(str(child))

            elif (Path(i).is_file() and str(Path(i)).endswith(fileTypes)) == False:
                # perform file check
                print("Error:   Unsupported file format: ",
                      Path(i), file=sys.stderr)

        SEARCH(File_List)

    # Checking pattern optional arguments ("-P", '--python-regex', "-w", '--word-regexp')

    def Check_Optional_Args(val):

        if args.python_regex == True:
            return re.search(r"%s" % args.PATTERN, str(val))

        elif args.word_regexp == True:
            if args.ignore_case == True:
                return str(args.PATTERN).upper() == (str(val).upper())
            else:
                return args.PATTERN == (str(val))

        elif args.word_regexp == False:
            if args.ignore_case == False:
                return args.PATTERN == args.PATTERN in (str(val))
            else:
                return str(args.PATTERN).upper() in (str(val).upper())
        else:
            return print("...Some Error Occured...(optional arguments!?)")

    # Checking output optional arguments ("-H", '--with-filename', "-N", '--with-sheetname')

    def Show_Filename_And_Sheetname(file, active_sheet, linesArray):
        ENDSWITH = "\n"
        if args.null:
            ENDSWITH = ""

        if args.count == True:
            pass

        elif args.files_with_match:
            pass

        elif args.files_without_match:
            pass

        elif args.with_filename == True:
            if args.with_sheetname == True:
                return print(
                    file
                    + ": "
                    + active_sheet
                    + ": "
                    + str(args.separator)
                    + str(args.separator).join(map(str, linesArray)),
                    end=ENDSWITH,
                )

            elif args.with_sheetname == False:
                return print(
                    file + ": " + str(args.separator) +
                    str(args.separator).join(map(str, linesArray)),
                    end=ENDSWITH,
                )

        elif args.with_filename == False:
            if args.with_sheetname == True:
                return print(
                    active_sheet
                    + ": "
                    + str(args.separator)
                    + str(args.separator).join(map(str, linesArray)),
                    end=ENDSWITH,
                )

            else:
                print(*linesArray, sep=str(args.separator), end=ENDSWITH)

    # Iterate over rows and columns and append matches count to array.

    SumOfROW, SumOfCELL, SumOfSTR = [], [], []

    def Iterate_Over_Cells(book, file):
        ROWcount, CELLcount, STRcount = [0], [0], [0]
        for key, item in book.items():
            for line in item:
                AuxFlag = False
                for cell in line:
                    if Check_Optional_Args(cell):
                        if args.count:
                            AuxFlag = True
                            CELLcount[0] = CELLcount[0] + 1
                            reESCapedQuery = re.escape(
                                str(args.PATTERN).upper())
                            STRcell = str(cell).upper()
                            if args.python_regex == False:
                                for x in re.findall(reESCapedQuery, STRcell):
                                    STRcount[0] = STRcount[0] + 1

                            else:
                                for x in re.findall(str(args.PATTERN), str(cell)):
                                    STRcount[0] = STRcount[0] + 1
                        else:
                            AuxFlag = True
                            ROWcount[0] = ROWcount[0] - 1

                if AuxFlag == True:
                    ROWcount[0] = ROWcount[0] + 1

                    Show_Filename_And_Sheetname(file, key, line)

        if ROWcount[0] > 0:
            ENDSWITH = "\n"
            ROWS, CELLS, STRINGS = ROWcount, CELLcount, STRcount
            if args.null:
                ENDSWITH = ""
            if args.with_sheetname or args.with_filename:
                print(
                    file,
                    ":",
                    ROWS[0],
                    "Rows, ",
                    CELLS[0],
                    "Cells, ",
                    STRINGS[0],
                    "Strings",
                    end=ENDSWITH,
                )
            SumOfCELL.extend(CELLcount)
            SumOfSTR.extend(STRcount)
            SumOfROW.extend(ROWcount)

    # Check files-with-match and files-without-match arguments

    def HyphenlAndHyphenLCheck(book, file):
        ENDSWITH = "\n"
        if args.null:
            ENDSWITH = ""

        if args.files_with_match:
            if Check_Optional_Args(book):
                return print(file, end=ENDSWITH)

        elif args.files_without_match:
            if not Check_Optional_Args(book):
                return print(file, end=ENDSWITH)

        else:

            Iterate_Over_Cells(book, file)

    # Count matches. Rows, cells and strings.

    def SumOfRowsCellsAndStrings():
        ROWS, CELLS, STRINGS = sum(SumOfROW), sum(SumOfCELL), sum(SumOfSTR)
        print("Search results: ", ROWS, "Rows, ",
              CELLS, "Cells, ", STRINGS, "Strings")

    # Opening files, start searching

    def SEARCH(File_List):
        for file in File_List:
            try:
                if args.debug == True:
                    warnings.resetwarnings()
                    print("-- debug mode: " + file)

                if file.endswith((".xlsx", ".XLSX", ".xlsm", ".XLSM")):
                    book = p.get_book_dict(
                        file_name=file, skip_hidden_row_and_column=False
                    )

                else:
                    book = p.get_book_dict(
                        file_name=file,
                    )

                HyphenlAndHyphenLCheck(book, file)

            except KeyboardInterrupt:

                sys.exit(0)

            except:
                print(
                    f"Error:\tUnsupported format, password protected or corrupted file: {file}",
                    file=sys.stderr,
                )
                pass

        if args.count:
            if args.files_with_match or args.files_without_match:
                pass

            else:
                SumOfRowsCellsAndStrings()

    File_And_Path_Location()


if __name__ == "__main__":
    main()
