# batchreplace
simple python script to generate a series of documents by replacing 
strings in a template with each row of an excel file. 

### references
using [openpyxl](https://openpyxl.readthedocs.io/en/stable/) for excel parsing.

find an unknown string between two delimiters by [using regular expressions](https://stackoverflow.com/questions/3368969/find-string-between-two-substrings).

string replacement inspired by
[this](https://www.geeksforgeeks.org/how-to-search-and-replace-text-in-a-file-in-python/) 
and [this](https://www.geeksforgeeks.org/python-program-to-replace-text-in-a-file/) examples.

### requirements
$ pip install openpyxl
