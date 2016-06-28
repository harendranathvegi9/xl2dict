Introduction
============

xl2dict is a python module to convert spreadsheets in to python dictionary. The input is a spreadsheet (xls or xlsx)
and the output is a list of dictionaries. The first row in the spreadsheet is treated as the header rows and each of the
cells in the first row assumes the keys in the output python dictionary. This python module will also enable the user
to seamlessly search for a data row in the speadsheet by specifying keyword / keywords . All the data rows containing
the specified keyword in any of their cells will be returned. This behavior is extremely useful in implementing
data driven and keyword driven tests and also in implementing object repositories for most opensource test automation
tools.This module will also enable the users to write data in to spreadsheet rows matching a
specified keyword / keywords, a feature that can be used to store dynamic data between dependent tests.

Installation
============

To install xl2dict, type the following command in the command line

.. code-block:: bash

    $ pip install xl2dict

Quickstart
==========

**1. convert_sheet_to_dict()**

This method will convert excel sheets to dict. The input is path to the excel file or a sheet object.
if file_path is None, sheet object must be provded. This method will convert only the first sheet.
If you need to convert multiple sheets, please use the method fetch_data_by_column_by_sheet_name_multiple() and
fetch_data_by_column_by_index_multiple().If you need to filter data by a specific keyword, specify the dict in
filter_variables_dict like {column name : keyword} . Any rows that matches the keyword in the specified column
will be returned. Multiple keywords can be specified.

Usage example::

    myxlobject= XlToDict()
    myxlobject.convert_sheet_to_dict(file_path="Users/xyz/Desktop/myexcel.xls", sheet="First Sheet",
                                     filter_variables_dict={"User Type" : "Admin", "Environment" : "Dev"})


**2. fetch_data_by_column_by_sheet_name()**

This method will convert the specified sheet in the excel file to dict. The input is path to the excel file .
If sheet_name is not provided, this method will convert only the first sheet.
If you need to convert multiple sheets, please use the method fetch_data_by_column_by_sheet_name_multiple() or
fetch_data_by_column_by_sheet_index_multiple(). If you need to filter data by a specific keyword,
specify the dict in filter_variables_dict like {column name : keyword} . Any rows that matches the keyword in
the specified column will be returned. Multiple keywords can be specified.

Usage example::

    myxlobject= XlToDict()
    myxlobject.fetch_data_by_column_by_sheet_name(file_path="Users/xyz/Desktop/myexcel.xls",
                                                  sheet_name="First Sheet",
                                                  filter_variables_dict={"User Type" : "Admin", "Environment" : "Dev"})

**3. fetch_data_by_column_by_sheet_index()**

This method will convert the specified sheet in the excel file to dict. The input is path to the excel file .
If sheet_index is not provided, this method will convert only the first sheet.
If you need to convert multiple sheets, please use the method fetch_data_by_column_by_sheet_name_multiple() or
fetch_data_by_column_by_sheet_index_multiple(). If you need to filter data by a specific keyword,
specify the dict in filter_variables_dict like {column name : keyword} . Any rows that matches the keyword in
the specified column will be returned. Multiple keywords can be specified.

Usage example::

    myxlobject= XlToDict()
    myxlobject.fetch_data_by_column_by_sheet_index(file_path="Users/xyz/Desktop/myexcel.xls",
                                                   sheet_index=1,
                                                   filter_variables_dict={"User Type" : "Admin", "Environment" : "Dev"})

**4. fetch_data_by_column_by_sheet_name_multiple()**

This method will convert multiple sheets in the excel file to dict. The input is path to the excel file .
If sheet_names is not provided, this method will convert ALL the sheets.If you need to filter data by a specific
keyword / keywords, specify the dict in filter_variables_dict like {column name : keyword} .
Any rows that matches the keyword  in the specified column will be returned. Multiple keywords can be specified.

Usage example::

    myxlobject= XlToDict()
    myxlobject.fetch_data_by_column_by_sheet_name_multiple(file_path="Users/xyz/Desktop/myexcel.xls",
                                                           sheet_names=["First Sheet","Some other sheet"],
                                                           filter_variables_dict={"User Type" : "Admin", "Environment" : "Dev"})

**5. fetch_data_by_column_by_sheet_index_multiple()**

This method will convert multiple sheets in the excel file to dict. The input is path to the excel file .
If sheet_indices is not provided, this method will convert ALL the sheets.If you need to filter data by a
specific keyword / keywords, specify the dict in filter_variables_dict like {column name : keyword} .
Any rows that matches the keyword  in the specified column will be returned. Multiple keywords can be specified.

Usage example::

    myxlobject= XlToDict()
    myxlobject.fetch_data_by_column_by_sheet_index_multiple(file_path="Users/xyz/Desktop/myexcel.xls",
                                                            sheet_indices=[0,1,4,7],
                                                            filter_variables_dict={"User Type" : "Admin", "Environment" : "Dev"})

**6. fetch_matching_data_row_indices()**

This method will fetch all the rows matching the specified filter. The input is path to the excel file .
If sheet_name_index is not provided, this method will search the first sheet sheet. If you need to filter data
by a specific keyword / keywords, specify the dict in filter_variables_dict like {column name : keyword} .
All the row indices that matches the keyword  in the specified column will be returned. Multiple keywords can be
specified.

Usage example::

    myxlobject= XlToDict()
    myxlobject.fetch_matching_data_row_indices(file_path="Users/xyz/Desktop/myexcel.xls",
                                               sheet_name_index="First Sheet",
                                               filter_variables_dict={"User Type" : "Admin", "Environment" : "Dev"})

    myxlobject.fetch_matching_data_row_indices(file_path="Users/xyz/Desktop/myexcel.xls",
                                               sheet_name_index=5,
                                               filter_variables_dict={"User Type" : "Admin", "Environment" : "Dev"})

**7. write_data_to_column()**

This method will write data in to the specified column of all the rows matching the specified filter. The input
is path to the excel file .If sheet_name is not provided, this method will write data in to the specified column
in the first sheet sheet. If you need to write data  in to rows by a specific keyword / keywords, specify the
dict in filter_variables_dict like {column name : keyword} .The specified data will be written in the specified
column in all rows that matches the keyword. Multiple keywords can be specified.


Usage example::

    myxlobject= XlToDict()
    myxlobject.write_data_to_column(file_path="Users/xyz/Desktop/myexcel.xls",column_name="Workorder Number",
                                    data="999999999", sheet_name="First Sheet",
                                    filter_variables_dict={"Test Case" : "Create Work Order", "Environment" : "Dev"})

