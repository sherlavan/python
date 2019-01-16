# -*- coding: cp1251 -*-
import os
import sys
from openpyxl import *
import types
import datetime
import time
import win32com.client as win32


# ==============search func===========
def search_row(list2d, unicodeSStrngs):
    res = []
    for i in range(len(list2d)):
        print "\rSearching: ", unicodeSStrngs[0],
        for j in range(len(list2d[i])):
            if type(list2d[i][j]) is unicode or type(list2d[i][j]) is str:
                for k in unicodeSStrngs:
                    if list2d[i][j].startswith(k):
                        res.append([i, j])
    print ''
    return res


# =========build list from sheet
def build_list(sheet, mc = 20):
    print "Building matrix.."
    rows = []
    c = 0
    # for row in sheet.rows:
    for row in sheet.iter_rows(max_col=mc):
        cells = []
        for cell in row:
            if cell.value:
                if type(cell.value) is unicode:
                    cells.append(unicode.strip(cell.value))
                else:
                    cells.append(cell.value)
        c += 1
        print '\rrow #', c,
        rows.append(cells)
    print ""
    print "Done building"
    return rows

# convert xls to xslx

def convertXLStoXLSX(fname):
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(fname)
    wb.SaveAs(fname+"x", FileFormat=51)    #FileFormat = 51 is for .xlsx extension
    wb.Close()                               #FileFormat = 56 is for .xls extension
    excel.Application.Quit()
    os.remove(fname)
