## Owerview

*xlsxgrep* is a CLI tool to search text in XLSX, XLS, XLSM, CSV, TSV and ODS files. It works similarly to Unix/GNU Linux *grep*.

## Features

- Grep compatible: xlsxgrep tries to be compatible with Unix/Linux grep, where it makes sense.
  Some of grep options are supported (such as `-r`, `-i`  or `-c`).

- Search many XLSX, XLS, XLSM, CSV, TSV and ODS files at once, even recursively in directories.

- Regular expressions: Python regex and POSIX extended regular expressions (-E).

- Supported file types: csv, ods, tsv, xls, xlsx, xlsm.

- Works on all major platforms: Windows, macOS, BSD and Linux,

## Usage:
```

usage: xlsxgrep [-h] [-V] [-P] [-E] [-F] [-i] [-w] [-c] [-r] [-H] [-N] [-l] [-L] [-S SEPARATOR]
                [-Z] [-j JOBS] [--row | --column] [-d] PATTTERN FILE [FILE ...]

positional arguments:
  PATTERN                    use PATTERN as the pattern to search for.
  FILE                       file or path to folder

options:
  -h, --help                 show this help message and exit.
  -V, --version              display version information and exit.
  -P, --python-regex         PATTERN is a Python regular expression.
  -E, --extended-regexp      PATTERN is a POSIX extended regular expression (Default).
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
```

## Examples:
```sh
xlsxgrep -i "foo" foobar.xlsx
```
```sh
xlsxgrep -c -H "(?i)foo|bar" /folder
```
## Installation

```sh
pip install xlsxgrep
```
