## Owerview

*xlsxgrep* is a CLI tool to search text in XLSX and XLS files. It works similarly to Unix/GNU Linux *grep*.


## Features

- Grep compatible: xlsxgrep tries to be compatible with Unix/Linux grep,
    where it makes sense. Some of grep options are supported (such as `-r`, `-i`  or `-c`).

- Search many XLSX and XLS files at once, even recursively in directories.

- Regular expressions: Python regex.

## Usage:
```

usage: xlsxgrep [-h] [-i] [-P] [-w] [-H] [-c] [-N] [-r] [-V] [-sep SEPARATOR]
                pattern path [path ...]

positional arguments:
  pattern               use PATTERN as the pattern to search for.
  path                  file or folder location

optional arguments:
  -h, --help            show this help message and exit
  -i, --ignore-case     ignore case distinctions.
  -P, --python-regex    PATTERN is a Python regular expression.
  -w, --word-regexp     force PATTERN to match only whole words.
  -H, --with-filename   print the file name for each match.
  -c, --count           print only a count of matches per file
  -N, --with-sheetname  print the sheet name for each match.
  -r, --recursive       search directories recursively.
  -V, --version         display version information and exit.
  -sep SEPARATOR, --separator SEPARATOR
                        define custom list separator for output, default is
                        TAB
        
```

## Example:

```sh

     xlsxgrep "PATTERN" -H -N --sep=";" -r /path/to/folder
  
```
## Installation

```

 pip install xlsxgrep
 
 ```

 
 





 
 


