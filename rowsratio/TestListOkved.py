# !python.exe
# coding: cp1251
#

from __future__ import with_statement
import os
from datetime import datetime, date, time
from time import *
from types import *
import xlrd

_newDir         = 'e:\\Tmp\\Spss\\rowsratio\\'

#os.chdir(_newDir)

def ToPrintLog (sMess):
    print str(datetime.now().strftime("%d.%m.%Y %H:%M:%S ")) + str(sMess)

from ListOkvedClass import *

LiO = ListOkved(ToPrintLog, _newDir)
i = 0
for re in LiO.listOkvedmal:
    print i, re[0], re[1], re[2]
    i += 1