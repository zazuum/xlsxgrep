## Owerview

*xlsxgrep* is a command-line tool to search text in XLSX and XLS files. It works similary to Unix/Linux *grep*.


## Features

- Grep compatible: xlsxgrep tries to be compatible with Unix/Linux grep,
    where it makes sense. Some of grep options are supported (such as `-r`, `-i`  or `-c`).

- Search many XLSX and XLS files at once, even recursively in directories.

- Regular expressions: Python regex.

## Usage:
```
usage: xlsxgrep.py [-h] [-i] [-P] [-w] [-H] [-c] [-N] [-r] [-V]
                   [-sep SEPARATOR]
                   pattern path [path ...]

positional arguments:
  pattern               Use PATTERN as the pattern to search for.
  path                  file or folder location

optional arguments:
  -h, --help            show this help message and exit
  -i, --ignore-case     Ignore case distinctions.
  -P, --python-regex    PATTERN is a Python regular expression.
  -w, --word-regexp     Force PATTERN to match only whole words.
  -H, --with-filename   Print the file name for each match.
  -c, --count           Print only a count of matches per file
  -N, --with-sheetname  Print the sheet name for each match.
  -r, --recursive       Search directories recursively.
  -V, --version         Display version information and exit.
  -sep SEPARATOR, --separator SEPARATOR
                        Define custom list separator for output, default is
                        TAB
```

## Example

```sh
     $ xlsxgrep "myPATTERN" --with-filename --with-sheetname -sep=";" Document.xlsx
   Document.xlsx: Sheet1:   ;column1;column2;myPATTERN;column3;column4;column5;column6 
   
```
## Installation

```
 pip install xlsxgrep
 ```
 
## Windows compiled download
```
https://github.com/zazuum/pool/blob/master/xlsxgrep-compiled-exe/xlsxgrep.exe

MD5 hash of .\xlsxgrep.exe:
274d9d9aafdb7ad97d10bab1a873ab4c

SHA1 hash of .\xlsxgrep.exe:
6eea9d25739c9d4bd339d4a1fc809c6bd645c6e6

```





 
 


