=====
Usage
=====

To use xls2xlsx from the command line:

.. code-block:: console

    $ xls2xlsx [-v] file.xls ...

This will create ``file.xlsx`` in the current folder.  ``file.xls`` can be any ``.xls``, ``.htm``, or ``.mht`` file and can also be a URL.  The ``-v`` flag will print the input and output filename.

To use xls2xlsx in a project:

.. code:: python

    from xls2xlsx import XLS2XLSX
    x2x = XLS2XLSX("spreadsheet.xls")
    x2x.to_xlsx("spreadsheet.xlsx")

Alternatively:

.. code:: python

    from xls2xlsx import XLS2XLSX
    x2x = XLS2XLSX("spreadsheet.xls")
    wb = x2x.to_xlsx()

The xls2xlsx.to_xlsx method returns the filename given.  If no filename is provided, the method returns the openpyxl workbook.

The input file can be in any of the following formats:

* Excel 97-2003 workbook (``.xls``)
* Web page (``.htm``, ``.html``), optionally including a _Files folder
* Single file web page (``.mht``, ``.mhtml``)

The input specified can also be any of the following:

* A filename / pathname
* A url
* A file-like object (opened in Binary mode for ``.xls`` and either Binary or Text mode otherwise)
* The contents of a ``.xls`` file as a ``bytes`` object
* The contents of a ``.htm`` or ``.mht`` file as a ``str`` object

Note: The file format is determined by examining the file contents, *not* by looking at the file extension.

