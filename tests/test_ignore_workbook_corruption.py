#!/usr/bin/env python

"""Tests for `xls2xlsx` feature of ignoring workbook corruption error"""
"""This feature are depended on xlrd(>=2.0.0)"""
"""Highly recommend setting `ignore_workbook_corruption=True` """
from xls2xlsx import XLS2XLSX
import pytest

# By using of `XLS2XLSX` default set of `ignore_workbook_corruption`
# It will try to replace the old XLS2XLSX object by a new that are set `ignore_workbook_corruption=True`
# if raise `xlrd.compdoc.CompDocError: Workbook corruption: seen[2] == 4`
@pytest.mark.parametrize("bad_file_path", [
    "tests/cinputs/workbook_corruption.xls",
    "tests/cinputs/corrupted_error.xls",
])
def test_ignore_workbook_corruption(bad_file_path):
    test_passed = False
    try:
        excel = XLS2XLSX(f=bad_file_path)
        excel.to_xlsx('tests/outputs/workbook_corruption.xlsx')
    except Exception as e:
        test_passed = True

    assert test_passed

# Manually set `ignore_workbook_corruption=True`
# It will ignore `xlrd.compdoc.CompDocError: Workbook corruption
def test_ignore_workbook_corruption_with_set():
    bad_file_path = 'tests/cinputs/corrupted_error.xls'
    excel = XLS2XLSX(f=bad_file_path,ignore_workbook_corruption=True)
    assert(excel.to_xlsx('tests/outputs/corrupted_error.xlsx'))

# Manually set `ignore_workbook_corruption=True`
# It will ignore `xlrd.compdoc.CompDocError: Workbook corruption:
# this one is so corrupt it can't be read at all
def test_ignore_workbook_corruption_with_set():
    bad_file_path = 'tests/cinputs/workbook_corruption.xls'
    excel = XLS2XLSX(f=bad_file_path,ignore_workbook_corruption=True)
    assert(excel.to_xlsx('tests/outputs/workbook_corruption.xlsx') == None)

