#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import argparse
import re
import warnings
import logging
import os
import functools
from concurrent.futures import ProcessPoolExecutor
import pyexcel as p
from pathlib import Path
import locale
from textwrap import dedent

__license__ = "MIT"
__version__ = "0.0.34"
__author__ = "Ivan Cvitic"
__email__ = "cviticivan@gmail.com"
VERSION_INFO = [
    "xlsxgrep version: {0}".format(__version__),
    "Python version: {0}".format(
        " ".join(line.strip() for line in sys.version.splitlines())
    ),
    "Locale: {0}".format(".".join(str(s) for s in locale.getlocale())),
]


def check_optional_args(opts, val):
    pattern = opts["PATTERN"]
    if opts["python_regex"]:
        return re.search(r"%s" % pattern, str(val))
    elif opts["word_regexp"]:
        if opts["ignore_case"]:
            return str(pattern).upper() == str(val).upper()
        else:
            return str(pattern) == str(val)
    elif not opts["word_regexp"]:
        if not opts["ignore_case"]:
            return str(pattern) in str(val)
        else:
            return str(pattern).upper() in str(val).upper()
    return None


def count_matching_strings(opts, cell, STRcount):
    pattern = opts["PATTERN"]
    re_escaped = re.escape(str(pattern).upper())
    str_cell = str(cell).upper()
    if not opts["python_regex"]:
        for x in re.findall(re_escaped, str_cell):
            STRcount[0] += 1
    else:
        for x in re.findall(str(pattern), str(cell)):
            STRcount[0] += 1


def format_line_ending(opts):
    return "" if opts["null"] else "\n"


def format_single_match(opts, file, active_sheet, value):
    endswith = format_line_ending(opts)
    if opts["count"] or opts["files_with_match"] or opts["files_without_match"]:
        return ""

    if opts["with_filename"] and opts["with_sheetname"]:
        return f"{file}: {active_sheet}: {value}{endswith}"
    elif opts["with_filename"]:
        return f"{file}: {value}{endswith}"
    elif opts["with_sheetname"]:
        return f"{active_sheet}: {value}{endswith}"
    else:
        return f"{value}{endswith}"


def format_filename_and_sheetname(opts, file, active_sheet, linesArray):
    endswith = format_line_ending(opts)
    sep = str(opts["separator"])

    if opts["count"] or opts["files_with_match"] or opts["files_without_match"]:
        return ""

    if opts["with_filename"]:
        if opts["with_sheetname"]:
            return (
                file
                + ": "
                + active_sheet
                + ": "
                + sep
                + sep.join(map(str, linesArray))
                + endswith
            )
        else:
            return (
                file
                + ": "
                + sep
                + sep.join(map(str, linesArray))
                + endswith
            )
    else:
        if opts["with_sheetname"]:
            return (
                active_sheet
                + ": "
                + sep
                + sep.join(map(str, linesArray))
                + endswith
            )
        else:
            return sep.join(map(str, linesArray)) + endswith


def process_single_file(file, opts):
    stdout_lines = []
    stderr_lines = []
    SumOfROW, SumOfCELL, SumOfSTR = [0], [0], [0]

    if not opts["debug"]:
        warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
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
        warnings.filterwarnings(
            "ignore",
            category=DeprecationWarning,
            message=".*Flags not at the start of the expression*.",
        )
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
        logging.disable(logging.WARNING)
    else:
        warnings.resetwarnings()
        stdout_lines.append(f"-- debug mode: {file}\n")

    try:
        if file.endswith((".xlsx", ".XLSX", ".xlsm", ".XLSM")):
            book = p.get_book_dict(
                file_name=file, skip_hidden_row_and_column=False
            )
        else:
            book = p.get_book_dict(file_name=file)
    except KeyboardInterrupt:
        raise
    except Exception:
        stderr_lines.append(
            f"Error:\tUnsupported format, password protected or corrupted file: {file}\n"
        )
        return {
            "file": file,
            "stdout": stdout_lines,
            "stderr": stderr_lines,
            "counts": (0, 0, 0),
        }

    endswith = format_line_ending(opts)

    if opts["files_with_match"]:
        if check_optional_args(opts, book):
            stdout_lines.append(f"{file}{endswith}")
        return {
            "file": file,
            "stdout": stdout_lines,
            "stderr": stderr_lines,
            "counts": (0, 0, 0),
        }
    elif opts["files_without_match"]:
        if not check_optional_args(opts, book):
            stdout_lines.append(f"{file}{endswith}")
        return {
            "file": file,
            "stdout": stdout_lines,
            "stderr": stderr_lines,
            "counts": (0, 0, 0),
        }

    if opts["column"]:
        COLcount, CELLcount, STRcount = [0], [0], [0]
        for key, item in book.items():
            if not item:
                continue
            max_cols = max(len(row) for row in item)
            for col_idx in range(max_cols):
                column = [
                    row[col_idx] if col_idx < len(row) else ""
                    for row in item
                ]
                col_has_match = any(
                    check_optional_args(opts, cell) for cell in column
                )
                if not col_has_match:
                    continue

                COLcount[0] += 1
                if opts["count"]:
                    for cell in column:
                        if check_optional_args(opts, cell):
                            CELLcount[0] += 1
                            count_matching_strings(opts, cell, STRcount)
                else:
                    for cell in column:
                        out = format_single_match(opts, file, key, cell)
                        if out:
                            stdout_lines.append(out)

        if opts["count"] and COLcount[0] > 0:
            if opts["with_sheetname"] or opts["with_filename"]:
                stdout_lines.append(
                    f"{file} : {COLcount[0]} Columns,  {CELLcount[0]} Cells,  {STRcount[0]} Strings{endswith}"
                )
            SumOfROW[0] = COLcount[0]
            SumOfCELL[0] = CELLcount[0]
            SumOfSTR[0] = STRcount[0]
    else:
        ROWcount, CELLcount, STRcount = [0], [0], [0]
        for key, item in book.items():
            for line in item:
                AuxFlag = False
                for cell in line:
                    if check_optional_args(opts, cell):
                        if opts["count"]:
                            AuxFlag = True
                            CELLcount[0] += 1
                            count_matching_strings(opts, cell, STRcount)
                        else:
                            AuxFlag = True
                            ROWcount[0] -= 1

                if AuxFlag:
                    ROWcount[0] += 1
                    out = format_filename_and_sheetname(opts, file, key, line)
                    if out:
                        stdout_lines.append(out)

        if opts["count"] and ROWcount[0] > 0:
            if opts["with_sheetname"] or opts["with_filename"]:
                stdout_lines.append(
                    f"{file} : {ROWcount[0]} Rows,  {CELLcount[0]} Cells,  {STRcount[0]} Strings{endswith}"
                )
            SumOfROW[0] = ROWcount[0]
            SumOfCELL[0] = CELLcount[0]
            SumOfSTR[0] = STRcount[0]

    return {
        "file": file,
        "stdout": stdout_lines,
        "stderr": stderr_lines,
        "counts": (SumOfROW[0], SumOfCELL[0], SumOfSTR[0]),
    }


def SEARCH(File_List, opts):
    SumOfROW, SumOfCELL, SumOfSTR = [], [], []
    jobs = opts.get("jobs", 1)
    if jobs <= 0:
        jobs = os.cpu_count() or 1

    def process_result(res):
        for err in res["stderr"]:
            sys.stderr.write(err)
            sys.stderr.flush()
        for out in res["stdout"]:
            sys.stdout.write(out)
            sys.stdout.flush()
        r, c, s = res["counts"]
        if opts["count"] and (r > 0 or c > 0 or s > 0):
            SumOfROW.append(r)
            SumOfCELL.append(c)
            SumOfSTR.append(s)

    if jobs > 1 and len(File_List) > 1:
        max_workers = jobs
        worker_fn = functools.partial(process_single_file, opts=opts)
        try:
            with ProcessPoolExecutor(max_workers=max_workers) as executor:
                results = executor.map(worker_fn, File_List)
                for res in results:
                    process_result(res)
        except KeyboardInterrupt:
            sys.exit(0)
    else:
        for file in File_List:
            try:
                res = process_single_file(file, opts)
                process_result(res)
            except KeyboardInterrupt:
                sys.exit(0)

    if opts["count"]:
        if not (opts["files_with_match"] or opts["files_without_match"]):
            GROUPS, CELLS, STRINGS = sum(SumOfROW), sum(SumOfCELL), sum(SumOfSTR)
            group_label = "Columns" if opts["column"] else "Rows"
            print(
                "Search results: ",
                GROUPS,
                group_label + ", ",
                CELLS,
                "Cells, ",
                STRINGS,
                "Strings",
            )


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
  -j, --jobs JOBS            number of CPU cores/processes to use for search (default: 1).
      --row                  search rows and print matching rows (default).
      --column               search columns and print whole matching columns vertically.

examples:
    xlsxgrep -i "foo" foobar.xlsx
    xlsxgrep -c -H "(?i)foo|bar" /folder

For more details refer to man page.
"""
    parser = argparse.ArgumentParser(
        add_help=False,  # epilog=example_text,
        description=dedent(help_text),
        prog="xlsxgrep",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        usage=dedent(
            """
	    xlsxgrep [-h] [-V] [-P] [-F] [-i] [-w] [-c] [-r] [-H] [-N] [-l] [-L] [-S SEPARATOR] 
                [-Z] [-j JOBS] [--row | --column] [-d] PATTTERN FILE [FILE ...]


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
        "-j",
        "--jobs",
        help=argparse.SUPPRESS,
        required=False,
        default=1,
        type=int,
    )
    parser.add_argument(
        "-d",
        "--debug",
        help=argparse.SUPPRESS,
        required=False,
        default=False,
        action="store_true",
    )
    search_mode = parser.add_mutually_exclusive_group()
    search_mode.add_argument(
        "--row",
        help=argparse.SUPPRESS,
        action="store_true",
    )
    search_mode.add_argument(
        "--column",
        help=argparse.SUPPRESS,
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

        opts = {
            "PATTERN": args.PATTERN,
            "python_regex": args.python_regex,
            "fixed_strings": args.fixed_strings,
            "ignore_case": args.ignore_case,
            "word_regexp": args.word_regexp,
            "count": args.count,
            "recursive": args.recursive,
            "with_filename": args.with_filename,
            "with_sheetname": args.with_sheetname,
            "files_with_match": args.files_with_match,
            "files_without_match": args.files_without_match,
            "separator": args.separator,
            "null": args.null,
            "debug": args.debug,
            "row": args.row,
            "column": args.column,
            "jobs": args.jobs,
        }

        SEARCH(File_List, opts)

    File_And_Path_Location()


if __name__ == "__main__":
    main()
