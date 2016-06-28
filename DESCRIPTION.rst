xl2dict
=======

xl2dict is a python module to convert spreadsheets in to a dictionary. The input is a spreadsheet (xls or xlsx)
and the output is a list of dictionaries. The first row in the spreadsheet is treated as the header rows and each of the
 cells in the first row assumes the keys in the output dictionary. This module will also enable the user
 to seamlessly search for a data row in the speadsheet by specifying keyword / keywords . All the data rows containing
 the specified keyword in any of their cells will be returned. This behavior is extremely useful in implementing
 data driven and keyword driven tests and also in implementing object repositories for most opensource test automation
 tools.This module will also enable the users to write data in to spreadsheet rows matching a
 specified keyword / keywords, a feature that can be used to store dynamic data between dependent tests.