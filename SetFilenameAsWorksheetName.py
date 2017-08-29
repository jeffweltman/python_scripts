#!/usr/bin/python3
#SetFilenameAsWorksheetName.py
#created by Jeff Weltman
#v1.0 8/29/2017

def main():
    import os
    import openpyxl
    target_dir = ('C:\\Users\\administrator\\Documents\\PythonTest') # set target directory
    pathiter = (os.path.join(root, filename)
        for root, _, filenames in os.walk(target_dir)
        for filename in filenames
    )
    for path in pathiter:
        filename = path.lstrip(target_dir) # this gets the name and extension of the current workbook
        book = filename.rstrip('.xlsx') # this strips the extension
        #print(book) 
        dollar = '$'
        newtitle = book + dollar
        badname = 'Workheet for {} was not correct.'.format(book)
        goodname = 'Workheet for {} is correct.'.format(book)
        from openpyxl import load_workbook
        wb = load_workbook(path)
        ws = wb.active
        #print(ws)
        #print(wb.sheetnames)
        for sheet in wb:
            if ws.title != newtitle:
                ws.title = newtitle
                print(badname)
            else:
                print(goodname)
                continue
        wb.save(path)

if __name__ == "__main__": main()
