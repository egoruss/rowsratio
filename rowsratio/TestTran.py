# !python.exe
# coding: cp1251
#
from PyQt4 import QtCore
class A(QtCore.QObject):
    def hello(self):
        return QtCore.QCoreApplication.translate("A", 'Привет!')

a = A()
print a.hello()
