***************************************************************************
*  TestSamples.py
***************************************************************************
* Проверка файлов выборок на полноту показателей, весов и данных
*  *  *  *
*  
*
*
***************************************************************************
* Последняя модификация: 15 июня 2010, Сергей Степанов. 
***************************************************************************

DEFINE @WORKDIR () 		'e:\Tmp\Spss\rowsratio\' 	!ENDDEFINE.
DEFINE @SAV ()  		'.sav'			!ENDDEFINE.
DEFINE @ListTOGS ()  		'\ListTOGS'		!ENDDEFINE.


SET PRINTBACK BOTH.  
SET MPRINT OFF.
set MXLOOPS = 1200.
set MITERATE = 20000.
* set WORKSPACE = 1000000.
set RNG = MT MTINDEX = 362.
set decimal dot.
*preserve. 
CD @WORKDIR.
DATASET CLOSE ALL. 
GET file = @WORKDIR + @ListTOGS + @SAV. 
DATASET NAME TOGS.
* NEW FILE. 

BEGIN PROGRAM PYTHON.
import __future__
import SpssClient
import spss 
import os

newDir         = 'e:\Tmp\Spss\rowsratio\'
nameListTOGS   = 'TOGS'
sav            = '.sav'
mal            = 'mal'
mic            = 'micro'
y2008          = '2008'

SpssClient.StartClient()
SpssClient.SetCurrentDirectory(newDir)
os.chdir(newDir)
print 'Текущая директория (Python): ', os.curdir, os.getcwd()
print 'Текущая директория (Spss  ): ', SpssClient.GetCurrentDirectory() 

spss.StartDataStep()
# Таблица описаний ТОГС
# SpssClient.LogToViewer('Чтение таблицы описаний ТОГС.')
datasetListTOGS = spss.Dataset(name=nameListTOGS)

print os.listdir(os.curdir)
print 'Склеивание имён файлов.'
sokato = '00'
for i in range(len(datasetListTOGS.cases)):
   if datasetListTOGS.cases[i,4][0] <= 99:
      sokato = '%02i' % datasetListTOGS.cases[i,4][0]
      nfoMal2008 = sokato + mal + y2008 + sav
      print i+1, datasetListTOGS.cases[i,4][0], nfoMal2008

spss.EndDataStep()
SpssClient.StopClient()

END PROGRAM.
*restore.
DATASET CLOSE ALL.
