# !python.exe
# coding: cp1251
#

from __future__ import with_statement
import os
import xlrd
from OkvedAssemblyClass import *

class ListOkved(object):
    def __init__ (self, ToPrint, ListOkvedDir):
        def getListOkved(self, nameFileList, rAss, grcol, sbcol, shrow):
#            self.listOk = {}
            self.listOk = []
            if os.access(ListOkvedDir + nameFileList, os.F_OK):
                self.rb = xlrd.open_workbook(ListOkvedDir + nameFileList,formatting_info=True,encoding_override="cp1251")
                self.sh = self.rb.sheet_by_index(0)
                i = 0
                for rownum in range(shrow, self.sh.nrows):
                    self.grOkved = str(self.sh.cell_value(rowx=rownum, colx=grcol).encode('cp1251')).strip()
                    if not self.grOkved == '':
#                        self.listOk[self.grOkved] = self.sh.cell_value(rowx=rownum, colx=sbcol).encode('cp1251').strip()
                        self.listOk.append([self.grOkved, self.sh.cell_value(rowx=rownum, colx=sbcol).encode('cp1251').strip(), rAss[self.grOkved]])
#                        print rownum, self.grOkved, self.listOk[self.grOkved]
                        print rownum+1, i+1, self.listOk[i][0], self.listOk[i][1], self.listOk[i][2]
                        
                        i += 1
                del self.rb
            else:
                ToPrint('�� ������!! ���� ������ ����� ' + nameFileList)
            return self.listOk
            
        ToPrint('������������� ������ ������ ������� �� �����. ')
#        self.listOkvedmal = {}
#        self.listOkvedmic = {}
        self.listOkvedmal = []
        self.listOkvedmic = []
        self.OkRul = OkvedAss(ToPrint, ListOkvedDir)
#        self.nameFileListOkvedmal = '\\total2009\\tabl_33_1401(1)_1121100010001.xls'
#        self.nameFileListOkvedmic = '\\total2009\\��� �����-1121100010001.xls'
        self.nameFileListOkvedmal = '\\����� �� 2007\\tab_33(01-09)_list1.xls'
        self.nameFileListOkvedmic = '\\����� �� 2007\\tab_33(01-09)_list1.xls'
        self.listOkvedmal = getListOkved(self, self.nameFileListOkvedmal, self.OkRul.rulesAssPlus, 0, 1, 6)
        self.listOkvedmic = getListOkved(self, self.nameFileListOkvedmic, self.OkRul.rulesAssPlus, 0, 1, 6)
