#!/usr/bin/python3
#ConvertXLSXtoCSV.py
#created by Jeff Weltman
#https://github.com/jeffweltman/python_scripts
#v1.0 10/30/2017
#Requires openpyxl package

def main():
    import os
    import openpyxl
    import csv
    target_dir = ('Documents') # set target directory
    for file in os.listdir(target_dir):
        filename = os.fsdecode(file)
        if filename.endswith('.xlsx'): #looking only at Excel docs
            book = filename.rstrip('.xlsx') # this strips the extension
            fullpath = target_dir + "/" + filename
            #print(book) 
            success = '{} converted to CSV.'.format(book)
            from openpyxl import load_workbook
            wb = load_workbook(fullpath)
            sh = wb.get_active_sheet()
            #print(ws)
            #print(sh.sheetnames)
            with open(fullpath.rstrip('.xlsx') + '.csv', 'w') as f:
                c = csv.writer(f)
                for r in sh.rows:
                # here, you can also filter non-ascii characters
                    c.writerow([cell.value for cell in r])
                print(success)    
                f.close()

if __name__ == "__main__": main()
