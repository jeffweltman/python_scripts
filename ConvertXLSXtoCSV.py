#!/usr/bin/python3
#ConvertXLSXtoCSV.py
#Created by Jeff Weltman
#https://github.com/jeffweltman/python_scripts
#v1.1 11/02/2017
#Requires openpyxl package

def main():
    import os
    import openpyxl
    import csv
    target_dir = ('Documents') # set target directory
    for file in os.listdir(target_dir):
        filename = os.fsdecode(file)
        if filename.endswith('.xlsx'): #looking only at Excel xlsx docs
            book = filename.rstrip('.xlsx') # this strips the extension
            fullpath = target_dir + "/" + filename
            success = '{} converted to CSV.'.format(book)
            from openpyxl import load_workbook
            wb = load_workbook(fullpath)
            sh = wb.get_active_sheet()
            with open(fullpath.rstrip('.xlsx') + '.csv', 'w') as f:
                c = csv.writer(f)
                for r in sh.rows:                
                    c.writerow([cell.value for cell in r])
                print(success) # prints successful conversions to console for each file   
                f.close()

if __name__ == "__main__": main()
