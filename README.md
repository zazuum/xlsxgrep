## Owerview

*xlsxgrep* is a CLI tool to search text in XLSX, XLS, XLSM, CSV, TSV and ODS files. It works similarly to Unix/GNU Linux *grep*.

## Features

- Grep compatible: xlsxgrep tries to be compatible with Unix/Linux grep, where it makes sense.
  Some of grep options are supported (such as `-r`, `-i`  or `-c`).

- Search many XLSX, XLS, XLSM, CSV, TSV and ODS files at once, even recursively in directories.

- Parallel execution: multi-process search via `-j` / `--jobs`, using all
  available CPU cores by default.

- Format-aware spreadsheet output: XLSX/XLSM and XLS numeric cells are rendered
  using their spreadsheet number format. ODS cells use their stored display text.
  CSV and TSV values are printed exactly as stored in the source file.

- Regular expressions: Python regex and POSIX extended regular expressions (-E).

- Search modes: row mode (`-R` / `--row`, default) prints matching rows;
  column mode (`-C` / `--column`) prints matching columns vertically. Count
  mode (`-c`) reports matching rows, columns, cells, and strings.

- Usable as a module: import `xlsxgrep` into Python code to run searches
  programmatically and get structured results.

- Works on all major platforms: Windows, macOS, BSD and Linux.

## Usage:
```

usage: xlsxgrep [-h] [-V] [-P] [-E] [-F] [-i] [-w] [-c] [-r] [-H] [-N]
                [-l] [-L] [-S SEPARATOR] [-Z] [-j JOBS] [-p] [-R | -C] 
                [-d] PATTTERN FILE [FILE ...]

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
  -j, --jobs JOBS            number of CPU cores to use for search (default: all cores).
  -p, --progress             display a progress bar during the search.
  -R, --row                  search rows and print matching rows (default).
  -C, --column               search columns and print whole matching columns vertically.
```

## Examples:
```sh
xlsxgrep -i "foo" foobar.xlsx
```
```sh
xlsxgrep -c -H "(?i)foo|bar" /folder
```

#### For more details and options, see xlsxgrep(1) or run 'man xlsxgrep'.
<br>

## Installation

```sh
pip install xlsxgrep
```

Or download a standalone installer or portable binary for Windows, macOS, or
Linux (`.deb`/`.rpm`/`.pkg.tar.zst`) from the [GitHub releases page](https://github.com/zazuum/xlsxgrep/releases).
