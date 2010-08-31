# !python.exe
# coding: cp1251
#

from __future__ import with_statement
import os
import xlrd

class OkvedAss(object):
    def __init__ (self, ToPrint, rulesDir):
        def insertGr(self, gr, listKod):
            listRet = []
#            print 'Проверяем ', listKod
            for ig, ingr in enumerate(listKod):               # Расширение подгруппы
                if ingr == gr:                               # Не сам из себя!
                    listRet = listRet + [ingr]   # Не расширяем, переносим как есть
                elif ingr in self.rulesAss:                   # Это подгруппа и Есть правило сборки
#                    print 'Найдена подгруппа', ingr, self.rulesAss[ingr]
                    listRet = listRet + insertGr(self, ingr, self.rulesAss[ingr])
                    
                else:                                       # Это самостоятельный код
                    listRet = listRet + [ingr]
#                print 'rulesAssPlus', gr, listRet
            return listRet
            
#        os.chdir(rulesDir)
        ToPrint('Инициализация Класса правил сборки ОКВЭД. ')
        self.nameFileRules = '\\okved_assembly.xls'
        self.rulesAss = {}
        self.rulesAssPlus = {}
        if os.access(rulesDir + self.nameFileRules, os.F_OK):
            self.rb = xlrd.open_workbook(rulesDir + self.nameFileRules,formatting_info=True,encoding_override="cp1251")
            self.sh = self.rb.sheet_by_index(0)
            self.grcol = 0
            self.sbcol = 1
            for rownum in range(self.sh.nrows):
                self.grOkved = str(self.sh.cell_value(rowx=rownum, colx=self.grcol).encode('cp1251')).split('=',1)[0].strip()
                if not self.grOkved == '':
                    self.rulesAss[self.grOkved] = str(self.sh.cell_value(rowx=rownum, colx=self.sbcol)).encode('cp1251').split('+')
#                    print '-', self.grOkved, '-', self.rulesAss[self.grOkved]
            del self.rb
            for gr, strOk in sorted(self.rulesAss.iteritems()):
                
#                if gr.split('.', 1)[0] in  ['DB', '17', '18']:
                    
#                    print 'rulesAss', gr, strOk
                    self.rulesAssPlus[gr] = [] + insertGr(self, gr, strOk)             # Расширение группы
                    print 'Собран ', gr, self.rulesAssPlus[gr]
        else:
            ToPrint('Не найден!! файл правил сборки ОКВЭД ' + rulesDir + self.nameFileRules)
        del self.rulesAss


