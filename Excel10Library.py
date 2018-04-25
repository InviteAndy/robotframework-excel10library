#!/usr/bin/env python


#  Copyright 2018 A.C. Hasper
#
#  Licensed under the Apache License, Version 2.0 (the "License");
#  you may not use this file except in compliance with the License.
#  You may obtain a copy of the License at
#
#      http://www.apache.org/licenses/LICENSE-2.0
#
#  Unless required by applicable law or agreed to in writing, software
#  distributed under the License is distributed on an "AS IS" BASIS,
#  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
#  See the License for the specific language governing permissions and
#  limitations under the License.

import os
import natsort
import re
from operator import itemgetter
from datetime import datetime, timedelta
import xlwings as xw
from version import VERSION
import atexit

_version_ = VERSION


class Excel10Library:
    """
    This library provides keywords to allow basic control
    over Excel10 (xlsx) files from Robot Framework.

    It is made to be compatible with ExcelLibrary for Robot Framework but
    some keyword are added to provide extra functionality

    This library depends heavily on xlwings
    """

    ROBOT_LIBRARY_SCOPE = 'GLOBAL'
    ROBOT_LIBRARY_VERSION = VERSION

    def __init__(self):
        self.wb = None
        self.tb = None
        self.sheetNum = None
        self.sheetNames = None
        self.fileName = None
        self.xa = xw.App(False)
        if os.name is "nt":
            self.tmpDir = "Temp"
        else:
            self.tmpDir = "tmp"
        atexit.register(self._exit)

    def _exit(self):
        self.xa.quit()

    def open_excel(self, filename, useTempDir=False):
        """
        Opens the Excel file from the path provided in the file name parameter.
        If the boolean useTempDir is set to true, depending on the operating system of the computer running the test the file will be opened in the Temp directory if the operating system is Windows or tmp directory if it is not.

        Arguments:
                |  File Name (string)                      | The file name string value that will be used to open the excel file to perform tests upon.                                  |
                |  Use Temporary Directory (default=False) | The file will not open in a temporary directory by default. To activate and open the file in a temporary directory, pass 'True' in the variable. |
        Example:

        | *Keywords*           |  *Parameters*                                      |
        | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |

        """

        if useTempDir is True:
            print 'Opening file at %s' % filename
            self.wb = xw.books.open(os.path.join("/", self.tmpDir, filename))
        else:
            self.wb = xw.books.open(filename)
        self.fileName = filename
        sheetnames = []
        for sh in self.wb.sheets:
            sheetnames.append(sh.name)
        self.sheetNames = sheetnames

    def open_excel_current_directory(self, filename):
        """
        Opens the Excel file from the current directory using the directory the test has been run from.

        Arguments:
                |  File Name (string)  | The file name string value that will be used to open the excel file to perform tests upon.  |
        Example:

        | *Keywords*           |  *Parameters*        |
        | Open Excel           |  ExcelRobotTest.xls  |

        """
        workdir = os.getcwd()
        print 'Opening file at %s' % filename
        self.wb = xw.books.open(os.path.join(workdir, filename), read_only=False, keep_vba=False)
        sheetnames = []
        for sh in self.wb.sheets:
            sheetnames.append(sh.name)
        self.sheetNames = sheetnames

    def get_sheet_names(self):
        """
        Returns the names of all the worksheets in the current workbook.

        Example:

        | *Keywords*              |  *Parameters*                                      |
        | Open Excel              |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |
        | Get Sheets Names        |                                                    |

        """

        sheetNames = self.sheetNames
        return sheetNames

    def get_number_of_sheets(self):
        """
        Returns the number of worksheets in the current workbook.

        Example:

        | *Keywords*              |  *Parameters*                                      |
        | Open Excel              |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |
        | Get Number of Sheets    |                                                    |

        """
        sheetNum = len(self.sheetNames)
        return sheetNum

    def get_column_count(self, sheetname, scanmaxrows=150):
        """
        Returns the specific number of columns of the sheet name specified.

        Arguments:
                |  Sheet Name (string)  | The selected sheet that the column count will be returned from. |
                |  scanmaxrows (default=150)    | The number of rows to scan for the right most column.   |
        Example:

        | *Keywords*          |  *Parameters*                                      |
        | Open Excel          |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |
        | Get Column Count    |  TestSheet1                                        |

        """
        #  XLwings provides no property to use so a scan is needed.
        sheet = self.wb.sheets[sheetname]
        most_right = 0
        for r in range(1,int(scanmaxrows)):
            last_col =  sheet.range((16384,r)).end('left')
            if last_col.column > most_right:
                most_right = last_col.column
        return most_right


    def get_row_count(self, sheetname, scanmaxcolumns = 100):
        """
        Returns the specific number of rows of the sheet name specified.

        Arguments:
                |  Sheet Name (string)  | The selected sheet that the row count will be returned from. |
                |  scanmaxcolumns (default=100)    | The number of columns to scan for the lowest row.   |
        Example:

        | *Keywords*          |  *Parameters*                                      |
        | Open Excel          |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |
        | Get Row Count       |  TestSheet1                                        |

        """

        # XLwings provides no property to use so a scan is needed.

        sheet = self.wb.sheets[sheetname]
        lowest = 0
        for col in range(1, int(scanmaxcolumns)):
            last_row =  sheet.range((1048575,col)).end('up')
            if last_row.row > lowest:
                lowest = last_row.row

        return lowest


    def get_column_values(self, sheetname, column, includeEmptyCells=True):
        """
        Returns the specific column values of the sheet name specified.

        Arguments:
                |  Sheet Name (string)                 | The selected sheet that the column values will be returned from.                                                            |
                |  Column (int)                        | The column integer value that will be used to select the column from which the values will be returned.                     |
                |  Include Empty Cells (default=True)  | The empty cells will be included by default. To deactivate and only return cells with values, pass 'False' in the variable. |
        Example:

        | *Keywords*           |  *Parameters*                                          |
        | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |   |
        | Get Column Values    |  TestSheet1                                        | 0 |

        """
        sheet = self.wb.sheets[sheetname]
        data = {}
        last_row = sheet.range((1048575, int(column))).end('up')

        for cell in sheet.range((1,int(column)), last_row):
            addr = cell.get_address(False,False)
            data[addr] = cell.value
        if includeEmptyCells is True:
            sorted_data = natsort.natsorted(data.items(), key=itemgetter(0))
            return sorted_data
        else:
            data = dict([(k, v) for (k, v) in data.items() if v])
            ordered_data = natsort.natsorted(data.items(), key=itemgetter(0))
            return ordered_data

    def get_row_values(self, sheetname, row, includeEmptyCells=True):
        """
        Returns the specific row values of the sheet name specified.

        Arguments:
                |  Sheet Name (string)                 | The selected sheet that the row values will be returned from.                                                               |
                |  Row (int)                           | The row integer value that will be used to select the row from which the values will be returned.                           |
                |  Include Empty Cells (default=True)  | The empty cells will be included by default. To deactivate and only return cells with values, pass 'False' in the variable. |
        Example:

        | *Keywords*           |  *Parameters*                                          |
        | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |   |
        | Get Row Values       |  TestSheet1                                        | 0 |

        """
        sheet = self.wb.sheets[sheetname]
        data = {}
        last_column =  sheet.range("XFD"+str(row)).end('left')
        for cell in sheet.range((int(row), 1), last_column):
            addr = cell.get_address(False, False)
            data[addr] = cell.value
        if includeEmptyCells is True:
            sorted_data = natsort.natsorted(data.items(), key=itemgetter(0))
            return sorted_data
        else:
            data = dict([(k, v) for (k, v) in data.items() if v])
            ordered_data = natsort.natsorted(data.items(), key=itemgetter(0))
            return ordered_data

    def get_sheet_values(self, sheetname, includeEmptyCells=True):
        """
        Returns the values from the sheet name specified.

        Arguments:
                |  Sheet Name (string)                 | The selected sheet that the cell values will be returned from.                                                              |
                |  Include Empty Cells (default=True)  | The empty cells will be included by default. To deactivate and only return cells with values, pass 'False' in the variable. |
        Example:

        | *Keywords*           |  *Parameters*                                      |
        | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |
        | Get Sheet Values     |  TestSheet1                                        |

        """
        sheet = self.wb.sheets[sheetname]
        data = {}
        last_row = self.get_column_count(sheetname)
        last_col = self.get_row_count(sheetname)
        for cell in sheet.range((1, 1), (last_row, last_col)):
            data[cell.get_address(False, False)] = cell.value
        if includeEmptyCells is True:
            sorted_data = natsort.natsorted(data.items(), key=itemgetter(0))
            return sorted_data
        else:
            data = dict([(k, v) for (k, v) in data.items() if v])
            ordered_data = natsort.natsorted(data.items(), key=itemgetter(0))
            return ordered_data

    def get_workbook_values(self, includeEmptyCells=True):
        """
        Returns the values from each sheet of the current workbook.

        Arguments:
                |  Include Empty Cells (default=True)  | The empty cells will be included by default. To deactivate and only return cells with values, pass 'False' in the variable. |
        Example:

        | *Keywords*           |  *Parameters*                                      |
        | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |
        | Get Workbook Values  |                                                    |

        """

        #Only included for compatibility reasons.  Always failed to see the use of this function
        self.get_sheet_values(self.wb.sheets.active, includeEmptyCells=includeEmptyCells)

    def read_cell_data_by_name(self, sheetname, cell_name):
        """
        Uses the cell name to return the data from that cell.

        Arguments:
                |  Sheet Name (string)  | The selected sheet that the cell value will be returned from.  |
                |  Cell Name (string)   | The selected cell name that the value will be returned from.   |
        Example:

        | *Keywords*           |  *Parameters*                                             |
        | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |      |
        | Get Cell Data        |  TestSheet1                                        |  A2  |

        """
        sheet = self.wb.sheets[sheetname]
        cell = sheet.range(cell_name)
        return cell.value


    def read_cell_data_by_coordinates(self, sheetname, column, row):
        """
        Uses the column and row to return the data from that cell.

        Arguments:
                |  Sheet Name (string)  | The selected sheet that the cell value will be returned from.         |
                |  Column (int)         | The column integer value that the cell value will be returned from.   |
                |  Row (int)            | The row integer value that the cell value will be returned from.      |
        Example:

        | *Keywords*     |  *Parameters*                                              |
        | Open Excel     |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |   |   |
        | Read Cell      |  TestSheet1                                        | 0 | 0 |

        """
        #import pdb, sys
        #pdb.Pdb(stdout=sys.__stdout__).set_trace()
        value = self.read_cell_data_by_name(sheetname, (int(row),int(column)))
        return value

    def check_cell_type(self, sheetname, column, row):
        """
        Checks the type of value that is within the cell of the sheet name selected.

        Arguments:
                |  Sheet Name (string)  | The selected sheet that the cell type will be checked from.          |
                |  Column (int)         | The column integer value that will be used to check the cell type.   |
                |  Row (int)            | The row integer value that will be used to check the cell type.      |
        Example:

        | *Keywords*           |  *Parameters*                                              |
        | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |   |   |
        | Check Cell Type      |  TestSheet1                                        | 0 | 0 |

        """
        sheet = self.wb.sheets[sheetname]
        cell = sheet.range((int(row),int(column))).value

        if type(cell) is float:
            celltype="number"
        elif type(cell) is str:
            celltype = "string"
        elif type(cell) is datetime:
            celltype = "date"
        elif type(cell) is bool:
            celltype = "boolean"
#        elif cell.data_type is XL_CELL_ERROR:
#            print "The cell value has an error"
        elif cell.value == "":
            celltype = "blank"
        elif cell.value is None:
            celltype = "empty"
        else:
            celltype = "Unknown"
        print "The cell value is a " + celltype
        return  celltype

    def put_value_to_cell(self, sheetname, column, row, value):
        """
        Using the sheet name the value of the indicated cell is set to be the number given in the parameter.

        Arguments:
                |  Sheet Name (string) | The selected sheet that the cell will be modified from.                                           |
                |  Column (int)        | The column integer value that will be used to modify the cell.                                    |
                |  Row (int)           | The row integer value that will be used to modify the cell.                                       |
                |  Value               | Any value that will set to the indicated cell. Type will be determines automatically              |
        Example:

        | *Keywords*           |  *Parameters*                                                          |
        | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls   |     |     |      |
        | Put Value To Cell    |  TestSheet1                                         |  0  |  0  |  34  |
        """
        sheet = self.wb.sheets[sheetname]
        #import pdb, sys
        #pdb.Pdb(stdout=sys.__stdout__).set_trace()
        if value.isdigit():
            sheet.range((int(row), int(column))).value = int(value)
        elif value.replace('.','',1).isdigit():
            sheet.range((int(row), int(column))).value = float(value)
        else:
            sheet.range((int(row), int(column))).value = value

    def put_number_to_cell(self, sheetname, column, row, value):
        """
        Using the sheet name the value of the indicated cell is set to be the number given in the parameter.

        Arguments:
                |  Sheet Name (string) | The selected sheet that the cell will be modified from.                                           |
                |  Column (int)        | The column integer value that will be used to modify the cell.                                    |
                |  Row (int)           | The row integer value that will be used to modify the cell.                                       |
                |  Value (int)         | The integer value that will be added to the specified sheetname at the specified column and row.  |
        Example:

        | *Keywords*           |  *Parameters*                                                         |
        | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xlsx |     |     |      |
        | Put Number To Cell   |  TestSheet1                                        |  0  |  0  |  34  |

        """
        self.put_value_to_cell(sheetname, column, row, value)

    def put_string_to_cell(self, sheetname, column, row, value):
        """
        Using the sheet name the value of the indicated cell is set to be the string given in the parameter.

        Arguments:
                |  Sheet Name (string) | The selected sheet that the cell will be modified from.                                           |
                |  Column (int)        | The column integer value that will be used to modify the cell.                                    |
                |  Row (int)           | The row integer value that will be used to modify the cell.                                       |
                |  Value (string)      | The string value that will be added to the specified sheetname at the specified column and row.   |
        Example:

        | *Keywords*           |  *Parameters*                                                           |
        | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xlsx |     |     |        |
        | Put String To Cell   |  TestSheet1                                        |  0  |  0  |  Hello |

        """
        sheet = self.wb.sheets[sheetname]
        if value.replace('.','',1).isdigit():
            sheet.range((int(row), int(column))).value = '\''+str(value)
        else:
            sheet.range((int(row), int(column))).value = value

    def put_date_to_cell(self, sheetname, column, row, value, informat='%d-%m-%Y'):
        """
        Using the sheet name the value of the indicated cell is set to be the date given in the parameter.

        Arguments:
                |  Sheet Name (string)               | The selected sheet that the cell will be modified from.                                                            |
                |  Column (int)                      | The column integer value that will be used to modify the cell.                                                     |
                |  Row (int)                         | The row integer value that will be used to modify the cell.                                                        |
                |  Value (int)                       | The integer value containing a date that will be added to the specified sheetname at the specified column and row. |
        Example:

        | *Keywords*           |  *Parameters*                                                               |
        | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |     |     |            |
        | Put Date To Cell     |  TestSheet1                                        |  0  |  0  |  12.3.1999 |

        """
        sheet = self.wb.sheets[sheetname]
        #import pdb, sys
        #pdb.Pdb(stdout=sys.__stdout__).set_trace()
        if format != '%d-%m-%Y':
            dateparts = re.split('(\d+)',value)
            if dateparts[2] != "-":
                value = value.replace(dateparts[2], "-")
        dt=datetime.strptime(value,informat)
        sheet.range((int(row), int(column))).options(dates=dt.date).value = dt


    def modify_cell_with(self, sheetname, column, row, op, val):
        """
        Using the sheet name a cell is modified with the given operation and value.

        Arguments:
                |  Sheet Name (string)  | The selected sheet that the cell will be modified from.                                                  |
                |  Column (int)         | The column integer value that will be used to modify the cell.                                           |
                |  Row (int)            | The row integer value that will be used to modify the cell.                                              |
                |  Operation (operator) | The operation that will be performed on the value within the cell located by the column and row values.  |
                |  Value (int)          | The integer value that will be used in conjuction with the operation parameter.                          |
        Example:

        | *Keywords*           |  *Parameters*                                                               |
        | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |     |     |     |      |
        | Modify Cell With     |  TestSheet1                                        |  0  |  0  |  *  |  56  |

        """
        sheet = self.wb.sheets[sheetname]
        cell = sheet.range((int(row), int(column)))
        if cell.value is None:
                cell.value = 0
        if val.replace('.','',1).isdigit():
            curval = cell.value
            cell.value = eval(str(curval)+str(op)+str(val))

    def add_to_date(self, sheetname, column, row, numdays):
        """
        Using the sheet name the number of days are added to the date in the indicated cell.

        Arguments:
                |  Sheet Name (string)             | The selected sheet that the cell will be modified from.                                                                          |
                |  Column (int)                    | The column integer value that will be used to modify the cell.                                                                   |
                |  Row (int)                       | The row integer value that will be used to modify the cell.                                                                      |
                |  Number of Days (int)            | The integer value containing the number of days that will be added to the specified sheetname at the specified column and row.   |
        Example:

        | *Keywords*           |  *Parameters*                                                        |
        | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |     |     |     |
        | Add To Date          |  TestSheet1                                        |  0  |  0  |  4  |

        """
        sheet = self.wb.sheets[sheetname]
        cell = sheet.range((int(row), int(column)))
        #import pdb, sys
        #pdb.Pdb(stdout=sys.__stdout__).set_trace()
        if type(cell.value) is datetime:
            dt=cell.value
            dated = timedelta(days=int(numdays))
            cell.value = dt + dated

    def subtract_from_date(self, sheetname, column, row, numdays):
        """
        Using the sheet name the number of days are subtracted from the date in the indicated cell.

        Arguments:
                |  Sheet Name (string)             | The selected sheet that the cell will be modified from.                                                                                 |
                |  Column (int)                    | The column integer value that will be used to modify the cell.                                                                          |
                |  Row (int)                       | The row integer value that will be used to modify the cell.                                                                             |
                |  Number of Days (int)            | The integer value containing the number of days that will be subtracted from the specified sheetname at the specified column and row.   |
        Example:

        | *Keywords*           |  *Parameters*                                                        |
        | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |     |     |     |
        | Subtract From Date   |  TestSheet1                                        |  0  |  0  |  7  |

        """

        self.add_to_date(sheetname, column, row, -int(numdays))

    def save_excel(self, filename, useTempDir=False):
        """
        Saves the Excel file indicated by file name, the useTempDir can be set to true if the user needs the file saved in the temporary directory.
        If the boolean useTempDir is set to true, depending on the operating system of the computer running the test the file will be saved in the Temp directory if the operating system is Windows or tmp directory if it is not.

        Arguments:
                |  File Name (string)                      | The name of the of the file to be saved.  |
                |  Use Temporary Directory (default=False) | The file will not be saved in a temporary directory by default. To activate and save the file in a temporary directory, pass 'True' in the variable. |
        Example:

        | *Keywords*           |  *Parameters*                                      |
        | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |
        | Save Excel           |  NewExcelRobotTest.xls                             |

        """
        if useTempDir is True:
            print '*DEBUG* Got fname %s' % filename
            self.wb.save(os.path.join("/", self.tmpDir, filename))
        else:
            self.wb.save(filename)

    def save_excel_current_directory(self, filename):
        """
        Saves the Excel file from the current directory using the directory the test has been run from.

        Arguments:
                |  File Name (string)    | The name of the of the file to be saved.  |
        Example:

        | *Keywords*                     |  *Parameters*                                      |
        | Open Excel                     |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |
        | Save Excel Current Directory   |  NewTestCases.xls                                  |

        """
        workdir = os.getcwd()
        print '*DEBUG* Got fname %s' % filename
        self.wb.save(os.path.join(workdir, filename))

    def add_new_sheet(self, newsheetname):
        """
        Creates and appends new Excel worksheet using the new sheet name to the current workbook.

        Arguments:
                |  New Sheet name (string)  | The name of the new sheet added to the workbook.  |
        Example:

        | *Keywords*           |  *Parameters*                                      |
        | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |
        | Add New Sheet        |  NewSheet                                          |

        """
        self.wb.sheets.add(newsheetname)

    def create_excel_workbook(self, newsheetname):
        """
        Creates a new Excel workbook

        Arguments:
                |  New Sheet Name (string)  | The name of the new sheet added to the new workbook.  |
        Example:

        | *Keywords*           |  *Parameters*                                      |
        | Open Excel           |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |
        | Create Excel         |  NewExcelSheet                                     |

        """
        self.wb=xw.books.add()
        self.add_new_sheet(newsheetname)

    def close_excel_workbook(self):
        """
        Closes current Excel workbook in memory. Good to use in Suite Teardown.

        Example:

        | *Keywords*              |  *Parameters*                                      |
        | Open Excel              |  C:\\Python27\\ExcelRobotTest\\ExcelRobotTest.xls  |
        | Close Excel Workbook    |                                                    |


        """
        self.wb.close()
