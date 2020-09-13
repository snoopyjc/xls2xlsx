********
xls2xlsx
********


.. image:: https://img.shields.io/pypi/v/xls2xlsx.svg
        :target: https://pypi.python.org/pypi/xls2xlsx

.. image:: https://img.shields.io/travis/snoopyjc/xls2xlsx.svg
        :target: https://travis-ci.com/snoopyjc/xls2xlsx

.. image:: https://readthedocs.org/projects/xls2xlsx/badge/?version=latest
        :target: https://xls2xlsx.readthedocs.io/en/latest/?badge=latest
        :alt: Documentation Status




Convert xls file to xlsx


* Free software: MIT license
* Documentation: https://xls2xlsx.readthedocs.io.


========
Features
========

* Convert ``.xls`` files to ``.xlsx`` using xlrd and openpyxl.
* Convert ``.htm`` and ``.mht`` files containing tables or excel contents to ``.xlsx`` using beautifulsoup4 and openpyxl.

We attempt to support anything that the underlying packages used will support.  For example, the following are supported for both input types:

* Multiple worksheets
* Text, Numbers, Dates/Times, Unicode
* Fonts, text color, bold, italic, underline, double underline, strikeout
* Solid and Pattern Fills with color
* Borders: Solid, Hair, Thin, Thick, Double, Dashed, Dotted; with color
* Alignment: Horizontal, Vertical, Rotated, Indent, Shrink To Fit
* Number Formats, including unicode currency symbols
* Hidden Rows and Columns
* Merged Cells
* Hyperlinks (only 1 per cell)
* Comments

These features are additionally supported by the ``.xls`` input format:

* Freeze panes

These features are additional supported by the ``.htm`` and ``.mht`` input formats:

* Images

Not supported by either format:

* Conditional Formatting (the current stylings are preserved)
* Formulas (the calculated values are preserved)
* Charts (the image of the chart is handled by ``.htm`` and ``.mht`` input formats)
* Pivot tables (the current data is preserved)
* Text boxes (converted to an image by ``.htm`` and ``.mht`` input formats)
* Shapes and Clip Art (converted to an image by ``.htm`` and ``.mht`` input formats)
* Autofilter (the current filtered out rows are preserved)
* Rich text in cells (openpyxl doesn't support this: only styles applied to the entire cell are preserved)

============
Installation
============

To install xls2xlsx, run this command in your terminal:

.. code-block:: console

    $ pip install xls2xlsx

This is the preferred method to install xls2xlsx, as it will always install the most recent stable release.

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


============
Dependencies
============

Python >= 3.6 is required.

These packages are also required: ``xlrd, openpyxl, requests, beautifulsoup4, Pillow, python-dateutil, cssutils, webcolors, currency-symbols, fonttools, PyYAML``.

====================
Implementation Notes
====================

The ``.htm`` and ``.mht`` input format conversion uses ImageFont from Pillow to measure the size (width and height) of cell contents.  The first time you use it, it will look for font files in standard places on your system and create a Font Name to filename mapping.  If the proper font files are not found on your system corresponding to the fonts used in the input file, then as a backup, an estimation algorithm is used.

If passed a ``.mht`` file (or url), the temporary folder name specified in the file will be used to unpack the contents for processing, then this folder will be removed when done.

=======
Credits
=======

Development Lead
----------------

* Joe Cool <snoopyjc@gmail.com>

Contributors
------------

None yet. Why not be the first?

================
Acknowledgements
================

A portion of the code is based on the work of John Ricco (johnricco226@gmail.com), Apr 4, 2017:
https://johnricco.github.io/2017/04/04/python-html/

This package was created with Cookiecutter_ and the `audreyr/cookiecutter-pypackage`_ project template.

.. _Cookiecutter: https://github.com/audreyr/cookiecutter
.. _`audreyr/cookiecutter-pypackage`: https://github.com/audreyr/cookiecutter-pypackage
