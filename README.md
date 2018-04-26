# robotframework-excel10library

Robot Framework Excel library compatible with .xlsx files
=========================================================

Excel10Library is a library for Robot Framework that supports the latest versions of Excel document format. 
This library is heavily depending on xlwings https://www.xlwings.org/

This library is compatible (as much as possible) with robotframework-excellibrary at https://github.com/NaviNet/robotframework-excellibrary with some minor differences:
- row/column are 1 based  (cell A1 is 1,1) in this library while robotframework-excellibrary has a row/column 0 base (cell A1 is 0,0)
- date formating is a made a bit more explicit: reading and editing existing cells with date values allows you to provide a 
date format so the change of date recognition is increased.
- detecting last row/last column with keywords: Get Row Count and Get Column Count has to be done by scanning the sheet. xlwing library does not provide methods or properies for retrieving the sheet working area. Optional paramters are added to rougly indicate the scanning boundries. Be carefull with high values. It will slow down the keywords and maybe result in a timeout.
- additional keyword: Close Workbook - allows you to close the workbook after saving or without saving.


Requirements
------------
robotframework-excel10library requires the following :

* Python 2.7.4 or newer. Older and newer versions are not tested but are likely to work (if its not to far back)
* Robot Framework 3.0.4 or newer

Python libraries:
* xlwings. Recommended to use: pip install xlwings
* natsort. Recommended to use: pip install natsort

Note: xlwings requires on Windows win32api (pip install pywin32) and comtypes (pip install comtypes). 
These libraries were not automatically installed during pip install xlwings (probably due to platform compatibility) so pywin32 and comtypes need to installed manually
See for more install info on xlwings site: http://docs.xlwings.org/en/stable/installation.html

For more details on how to use the Robot Framework see http://robotframework.org/

Important to know
------------------
- xlwings is a library that uses the excel program in the background. The benefit of this is that the calculations are performed when needed and the value of a cell with formulas can be read and used. I ran into problems with formulas when I tried to write a library with openpyxl.
Downside of this is that you cannot open a sheet in excel and in robot framework at the same time. You have to close the sheet in Excel before running your testscript.
- The Save Excel keyword generates a Excel overwrite warning when you overwrite an existing excelsheet. I think it originates from Excel, not from xlwings or Excel10Library. I hope the xlwings developers find a way to suppress the warning. I noticed this on Windows. Not sure how it behaves on MacOSX.

