#!/usr/bin/python3
#SetFilenameAsWorksheetName.py
#created by Jeff Weltman
#https://github.com/jeffweltman/python_scripts
#v1.0 8/29/2017

def main():
    import os
    import openpyxl
    target_dir = ('Documents') # set target directory
    for file in os.listdir(target_dir):
        filename = os.fsdecode(file)
        if filename.endswith('.xlsx'): #looking only at Excel docs
            book = filename.rstrip('.xlsx') # this strips the extension
            #print(book) 
            badname = 'Worksheet for {} was not correct.'.format(book)
            goodname = 'Worksheet for {} is correct.'.format(book)
            from openpyxl import load_workbook
            wb = load_workbook(filename)
            ws = wb.active
            #print(ws)
            #print(wb.sheetnames)
            for sheet in wb:
                if ws.title != book:
                    ws.title = book
                    print(badname)
                else:
                  print(goodname)
                  continue
                wb.save(filename)

if __name__ == "__main__": main()
