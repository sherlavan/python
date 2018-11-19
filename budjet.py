# -*- coding: cp1251 -*- 
import os
import sys
from openpyxl import *
import codecs
import types
import datetime
import time
from funct import build_list, search_row, convertXLStoXLSX


start_time = time.time()

fcount = 0
# create new xl doc
new_work_book = Workbook()
new_sheet = new_work_book.worksheets[0]
new_work_book_row = 1
new_work_book_col = 1
# add table head
table_head = [u'��� �����', u'������', u'���������', u'������ ������', u'�� �1', u'��-��� �1', u'��� �1',
              u'��������� �1', u'�� �1', u'�� �1', u'�� � �� �1', u'�� � 0,98', u'��-��� � 0,98', u'��� � 0,98',
              u'��������� � 0,98', u'�� � 0,98', u'�� � 0,98', u'�� � �� � 0,98', u'�� ����������',
              u'��-��� ����������', u'��� ����������', u'��������� ����������', u'�� ����������', u'�� ����������',
              u'�� � �� ����������']

for i in range(len(table_head)):
    new_sheet.cell(new_work_book_row, i + 1).value = table_head[i]

for file_name in os.listdir("./"):
    if file_name.endswith("xls"):
        path = os.getcwd() + '\\' + file_name
        path = path.replace('\'', '\\') + '\\'
        path = path.rstrip('\\')
        print u'������������ ����', path.decode('cp1251')
        convertXLStoXLSX(path)
new_work_book_row += 1
for file_name in os.listdir("./"):
    if all([file_name.endswith("xlsx"), file_name[0] != '~', file_name != 'budjet.xlsx', file_name != 'mat.xlsx',
            file_name != 'meh.xlsx']):

        wb = load_workbook(file_name)
        sheet = wb.worksheets[0]
        obj = ''
        per = ''
        flag = ''
        work_type = ''
        found_data = False
        jobs_done = 'XZ'
        data37 = []
        fcount += 1
        print u'parsing', file_name.decode('cp1251')
        rows = build_list(sheet)

        new_sheet.cell(new_work_book_row, 1).value = file_name.decode('cp1251')
        coordinates = search_row(rows, [u'������'])[0]
        new_sheet.cell(new_work_book_row, 2).value = rows[coordinates[0]][coordinates[1] + 1]

        coordinates = search_row(rows, [u"��������� �", u"������ �����������", u'������ ��������� �'])
        if coordinates:
            new_sheet.cell(new_work_book_row, 3).value = rows[coordinates[0][0]][coordinates[0][1]]
        else:
            new_sheet.cell(new_work_book_row, 3).value = u'����������'

        try:
            coordinates = search_row(rows, [u'���� �����������'])[0]
            new_sheet.cell(new_work_book_row, 4).value = rows[coordinates[0] + 2][-1]
        except:
            new_sheet.cell(new_work_book_row, 4).value = u'����������'
        # �������� �������� �������
        coordinates = search_row(rows, [u'�������� �������� �������'])
        try:
            new_sheet.cell(new_work_book_row, 5).value = rows[coordinates[0][0]][coordinates[0][1] + 1]
        except:
            new_sheet.cell(new_work_book_row, 5).value = u'����������'
        try:
            new_sheet.cell(new_work_book_row, 12).value = rows[coordinates[1][0]][coordinates[1][1] + 1]
        except:
            new_sheet.cell(new_work_book_row, 12).value = u'����������'
        # EMZPM1 = EM1 - ZPM1
        # � �.�. �������� ����������
        coordinates = search_row(rows, [u'� �.�. �������� ����������'])
        try:
            new_sheet.cell(new_work_book_row, 7).value = rows[coordinates[0][0]][coordinates[0][1] + 1]
            ZPM = rows[coordinates[0][0]][coordinates[0][1] + 1]
        except:
            new_sheet.cell(new_work_book_row, 7).value = u'����������'
            ZPM = 0
        try:
            new_sheet.cell(new_work_book_row, 14).value = rows[coordinates[1][0]][coordinates[1][1] + 1]
            ZPM09 = rows[coordinates[1][0]][coordinates[1][1] + 1]
        except:
            new_sheet.cell(new_work_book_row, 14).value = u'����������'
            ZPM09 = 0

        coordinates = search_row(rows, [u'������������ �����'])
        try:
            new_sheet.cell(new_work_book_row, 6).value = rows[coordinates[0][0]][coordinates[0][1] + 1] - ZPM
        except:
            new_sheet.cell(new_work_book_row, 6).value = u'����������'
        try:
            new_sheet.cell(new_work_book_row, 13).value = rows[coordinates[1][0]][coordinates[1][1] + 1] - ZPM09
        except:
            new_sheet.cell(new_work_book_row, 13).value = u'����������'
        # //em-zpm
        # materials
        coordinates = search_row(rows, [u'���������'])
        try:
            new_sheet.cell(new_work_book_row, 8).value = rows[coordinates[0][0]][coordinates[0][1] + 1]
        except:
            new_sheet.cell(new_work_book_row, 8).value = u'����������'
        try:
            new_sheet.cell(new_work_book_row, 15).value = rows[coordinates[1][0]][coordinates[1][1] + 1]
        except:
            new_sheet.cell(new_work_book_row, 15).value = u'����������'

        # ��������� ������� �� ��
        coordinates = search_row(rows, [u'��������� ������� �� ��'])
        try:
            new_sheet.cell(new_work_book_row, 9).value = rows[coordinates[0][0]][coordinates[0][1] + 1]
        except:
            new_sheet.cell(new_work_book_row, 9).value = u'����������'
        try:
            new_sheet.cell(new_work_book_row, 16).value = rows[coordinates[1][0]][coordinates[1][1] + 1]
        except:
            new_sheet.cell(new_work_book_row, 16).value = u'����������'
        # ������� ������� �� ��
        coordinates = search_row(rows, [u'������� ������� �� ��'])
        try:
            new_sheet.cell(new_work_book_row, 10).value = rows[coordinates[0][0]][coordinates[0][1] + 1]
        except:
            new_sheet.cell(new_work_book_row, 10).value = u'����������'
        try:
            new_sheet.cell(new_work_book_row, 17).value = rows[coordinates[1][0]][coordinates[1][1] + 1]
        except:
            new_sheet.cell(new_work_book_row, 17).value = u'����������'
        # �� � �� �� ���
        coordinates = search_row(rows, [u'�� � �� �� ���'])
        try:
            new_sheet.cell(new_work_book_row, 11).value = rows[coordinates[-2][0]][coordinates[-2][1] + 1]
        except:
            new_sheet.cell(new_work_book_row, 11).value = u'����������'
        try:
            new_sheet.cell(new_work_book_row, 18).value = rows[coordinates[-1][0]][coordinates[-1][1] + 1]
        except:
            new_sheet.cell(new_work_book_row, 18).value = u'����������'
        # =L2-E2*0,05

        new_sheet.cell(new_work_book_row, 19).value = '=L' + str(new_work_book_row) + '-E' + str(
            new_work_book_row) + '*0.05'
        new_sheet.cell(new_work_book_row, 20).value = '=M' + str(new_work_book_row) + '-F' + str(
            new_work_book_row) + '*0.05'
        new_sheet.cell(new_work_book_row, 21).value = '=N' + str(new_work_book_row) + '-G' + str(
            new_work_book_row) + '*0.05'
        new_sheet.cell(new_work_book_row, 22).value = '=H' + str(new_work_book_row) + '*0.05'

        new_sheet.cell(new_work_book_row, 23).value = '=P' + str(new_work_book_row) + '-I' + str(
            new_work_book_row) + '*0.05'
        new_sheet.cell(new_work_book_row, 24).value = '=Q' + str(new_work_book_row) + '-J' + str(
            new_work_book_row) + '*0.05'
        new_sheet.cell(new_work_book_row, 25).value = '=R' + str(new_work_book_row) + '-K' + str(
            new_work_book_row) + '*0.05'

        new_work_book_row += 1

new_work_book.save('./budjet.xlsx')
print fcount, 'files parsed for ',
print("--- %s seconds --- " % (time.time() - start_time))
