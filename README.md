## Owerview

*xlsxgrep* is a tool to search text in XLSX and XLS files. It works similary to *grep*.


This is a very early development version. While it technically works, much work needs to be done to better
organize the code and implement additional features. See "To-do" below.


## Features

- Grep compatible: xlsxgrep tries to be compatible with GNU grep,
    where it makes sense. Some of your favorite grep options are
    supported (such as `-r`, `-i`  or `-c`).

- Search many XLSX and XLS files at once, even recursively in directories.

- Regular expressions: Python regex.

## Usage:
```
usage: xlsxgrep [-h] [-i] [-P] [-w] [-H] [-c] [-N] [-r] pattern path [path ...]

positional arguments:
  pattern               Use PATTERN as the pattern to search for.
  path                  File or folder location

optional arguments:
  -h, --help            Show this help message and exit
  -i, --ignore-case     Ignore case distinctions.
  -P, --python-regex    PATTERN is a Python regular expression
  -w, --word-regexp     Force PATTERN to match only whole words
  -H, --with-filename   Print the file name for each match.
  -c, --count           Print only a count of matches per file
  -N, --with-sheetname  Print the sheet name for each match.
  -r, --recursive       Search directories recursively.
```

## Example

```sh
     $ xlsxgrep pattern --with-filename --with-sheetname Document.xlsx
   Document.xlsx: Sheet1: Line that match your pattern.
   
```
## Installation

```
 pip install xlsxgrep
 ```
 
 
 ### Windows binary download
 
 https://github.com/zazuum/pool/blob/master/xlsxgrep-compiled-exe/xlsxgrep.exe
```
MD5SUM 95f396801e663077f6061336340ec857 xlsxgrep.exe

SHA1SUM 67293471496dda6629a6eeeb922fd65c74d15fe5 xlsxgrep.exe

SHA256SUM 2a79c93ce376b5c6dc9e608fc64dea547bb0a09b3a956b31178df6e387ab18aa xlsxgrep.exe

SHA512SUM 3c97cd575628074d89c1102221dc47d78a4d9cdd6f20f2f82a4ded24493d66a6b4c7c33d89ebed9667a3b
6ffc69e74df8ea9faf6eedb943286975398b24f0281 xlsxgrep.exe
```

## TODO

- Rewrite the whole thing from the scratch. :-D  
- Add new options.


