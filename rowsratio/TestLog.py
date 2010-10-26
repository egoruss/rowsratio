# !python.exe
# coding: cp1251
#

from __future__ import with_statement
import os
from datetime import datetime, date, time
from time import *
from types import *
import sip
sip.setapi('QVariant', 2)
from PyQt4 import QtCore, QtGui 

class Logger(object):
    def __init__(self, output, lCursor):
        self.output = output
        self.lC = lCursor
    def write(self, string):
        if not (string == "\n"):
            trstring = QtGui.QApplication.translate("MainWindow", string.replace("\n",'').decode('cp1251').strip(), None, QtGui.QApplication.UnicodeUTF8)
            self.output.append(trstring)
            self.output.moveCursor(self.lC.End, mode=self.lC.MoveAnchor)

class myQThread(QtCore.QThread):
    def __init__(self, output):
        QtCore.QThread.__init__(self)
        self.output = output
 
    def run(self):
        pass
        import RowsRatioD

class ProgressBar(QtGui.QWidget):
    def __init__(self, parent=None):
        QtGui.QWidget.__init__(self, parent)
        self.pbar = QtGui.QProgressBar()


class MainWindow(QtGui.QMainWindow):
    def __init__(self, fileName=None):
        super(MainWindow, self).__init__()
        self.setWindowTitle(u'Вычисление скорректированных итогов')
        self.setAttribute(QtCore.Qt.WA_DeleteOnClose)
        self.isUntitled = True
#        self.main_widget = QtGui.QWidget()
#        self.setCentralWidget(self.main_widget)

        self.progress = QtGui.QProgressBar(self)
        self.progress.setMaximum(100)
        self.downDock = QtGui.QDockWidget()
        self.downDock.setAllowedAreas(QtCore.Qt.BottomDockWidgetArea)
        self.downDock.setWidget(self.progress);
        self.addDockWidget(QtCore.Qt.BottomDockWidgetArea, self.downDock)

#        self.vbox = QtGui.QVBoxLayout(self)
#        self.vbox.addWidget(self.logText)
#        self.vbox.addWidget(self.progress)
#        self.main_widget.setLayout(self.vbox)
#        self.setLayout(self.vbox)
        
        self.logText = QtGui.QTextEdit()
        self.setCentralWidget(self.logText)
        # даем виджету свойство read-only
        self.logText.setReadOnly(True)
        self.logText.setCurrentFont(QtGui.QFont('Courier New', 9))
        # делаем полосу вертикальной прокрутки видимой всегда
        self.logText.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOn)
        self.logCursor = QtGui.QTextCursor()
        # можно как перенаправлять все в один виджет, так и разделить ошибки и 
        # стандартный вывод
        self.logger = Logger(self.logText, self.logCursor)
        self.errors = Logger(self.logText, self.logCursor)
        sys.stdout = self.logger
#        sys.stderr = self.errors
#        sys.stderr = self.logger
       
        self.readSettings() 
        self.createStatusBar()
        print 'Создание окна завершено.'
        # начальное сообщение
        self.logText.append(".:: Старт лога ::.".decode('cp1251'))
        self.progress.setValue(1)
        
    @QtCore.pyqtSignature("")
    def threadFinished(self):
        self.logger.write("Поток завершен!")

        
    def closeEvent(self, event):
        self.writeSettings()
        event.accept() 
 
    def createStatusBar(self):
        self.statusBar().showMessage(u"Готово")

    def readSettings(self):
        print os.getcwd()
        settings = QtCore.QSettings('e:\\Tmp\\Spss\\rowsratio'+'\\rowsratio.ini', 1)
#        settings = QtCore.QSettings('StepService', 'RowsRatio')
#        settings.setPath( )
        pos = settings.value('pos', QtCore.QPoint(200, 200))
        size = settings.value('size', QtCore.QSize(400, 400))
        self.move(pos)
        self.resize(size)

    def writeSettings(self):
        settings = QtCore.QSettings('e:\\Tmp\\Spss\\rowsratio'+'\\rowsratio.ini', 1)
#        settings = QtCore.QSettings('StepService', 'RowsRatio')
        settings.setValue('pos', self.pos())
        settings.setValue('size', self.size()) 
 


if __name__ == '__main__':

    import sys

    app = QtGui.QApplication(sys.argv)
#    app.qRegisterMetaType(QtGui.QTextCursor())
    mainWin = MainWindow()
    mainWin.show()
    logTread = myQThread(mainWin.logger)
    print 'Перевод вывода нового потока в окно.'
    logTread.start()

    sys.exit(app.exec_()) 
    