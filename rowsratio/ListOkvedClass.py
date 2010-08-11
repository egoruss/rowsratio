# !python.exe
# coding: cp1251
#

from __future__ import with_statement
import os
import xlrd

class ListOkved(object):
    def __init__ (self, ToPrint, ListOkvedDir):
        def getListOkved(self, nameFileList, grcol, sbcol, shrow):
            self.listOk = {}
            if os.access(ListOkvedDir + nameFileList, os.F_OK):
                self.rb = xlrd.open_workbook(ListOkvedDir + nameFileList,formatting_info=True,encoding_override="cp1251")
                self.sh = self.rb.sheet_by_index(0)
                for rownum in range(shrow, self.sh.nrows):
                    self.grOkved = str(self.sh.cell_value(rowx=rownum, colx=grcol).encode('cp1251')).strip()
                    if not self.grOkved == '':
                        self.listOk[self.grOkved] = self.sh.cell_value(rowx=rownum, colx=sbcol).encode('cp1251').strip()
#                        print rownum, self.grOkved, self.listOk[self.grOkved]
                del self.rb
            else:
                ToPrint('Не найден!! файл списка ОКВЭД ' + nameFileList)
            return self.listOk
            
        ToPrint('Инициализация Класса списка разреза по ОКВЭД. ')
        self.listOkvedmal = {}
        self.listOkvedmic = {}
        self.nameFileListOkvedmal = '\\total2009\\tabl_33_1401(1)_1121100010001.xls'
        self.nameFileListOkvedmic = '\\total2009\\СЧР всего-1121100010001.xls'
        self.listOkvedmal = getListOkved(self, self.nameFileListOkvedmal, 0, 1, 6)
        self.listOkvedmic = getListOkved(self, self.nameFileListOkvedmic, 0, 1, 6)
