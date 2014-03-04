excelpy
=======
* Excel 2010 Library for Python 3.3
* You can add sheets, rename sheets, copy sheets, delete sheets and edit string or number type data.
* excelpy needs json data for edit(See Usage)


License
-------
MIT License


Require
-------
* Python 3.3
* lxml 3.3.0


Usage
-----
    from excelpy import ExcelPy
    excel = ExcelPy(test_excel.xlsx)
    
    excel.addSheet('Next')
    excel.addSheet('Revolution')
    
    excel.deleteSheet('Sheet2')
    
    excel.copySheet('Slayers')
    # copy as 'Slasyers_copy'
    
    excel.copySheet('Slayers', '슬레이어즈')
    
    excel.renameSheet('Next', 'Try')
    
    changed = json.dumps({'A1': 'Title', 'C8': 'Diary', 'D9': 'Title', 'A12': '슬레이어즈'}, ensure_ascii=False)
    excel.edit(sheetname='Slayers', changed_data_json=changed)
    
	excel.save()


To Do
-----
* Support Python 2.7
* Change the count property value of <sst> Element in 'xl/sharedString.xml' file when cell changed, add/delete/copy sheets.
* change default sheet name when not given copied sheetname
* Sheet ordering
* Add testcase
* Refactoring
* Et Ceteras