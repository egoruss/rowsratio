# !python.exe
# coding: cp1251
#

from __future__ import with_statement
import os
from datetime import datetime, date, time
from time import *
from types import *
import xlrd

stro = 'e:\\Tmp\\Spss\\rowsratio\\okved_assembly.xls'
s = '2.0'
print stro.split('\\',0)
print stro.split('\\',1)
print stro.split('\\',2)

print stro.rsplit('\\',0)
print stro.rsplit('\\',1)[1]
print stro.rsplit('\\',2)


if len(s.split('.',1)[0]) == 1:
    s = '0'+s
print s
