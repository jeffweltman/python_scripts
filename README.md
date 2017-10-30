# python_scripts
Miscellaneous Python scripts to standardize and simplify
## *Tested and verified for Python 3.6.2 on Windows 10*

### SetFilenameAsWorksheetName.py
 *Required packages: os, openpyxl*

This script will iterate through Excel workbooks in a specified directory and will set the worksheet name to match the filename. 
If the worksheet name already matches, it will continue without changing. 
Optional print commmands are included to output the filename and whether or not the worksheet was named correctly.
  
### RenameColumns.py
 *Required packages: os, openpyxl*
  
This script will iterate through Excel workbooks in a specified directory and replace specified column names with a target column name.

### ConvertXLSXtoCSV.py
 *Required packages: os, openpyxl, csv*

This script will iterate through Excel workbooks in a specified directory and convert the XLSX to a CSV. I had not found a great Python3 script to accomplish this, so I created one.
