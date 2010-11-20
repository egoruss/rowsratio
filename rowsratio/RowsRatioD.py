# !python.exe
# coding: cp1251
# 
#C:\Program Files (x86)\SPSSInc\PASWStatistics18\paswstat.com -production e:\Tmp\Spss\Rows\StartMain.spj
########################################
#  
#  To Do: 
#         + ����� �� ����������� �����.
#         + ��������� ������� ��� 2009 � 2007 �����.
#         + ������� ������ � XLS.
#         + ��������� ������ Excel-����� ����� ����� �������.
#         + ������ ����� �����.
#         + ���������� ������ 2007 �� �����.
#         + ����������������� �������� �� ����� � �����.
#         + ���� �� ��� ������� �����������.
#         + ������� ����� ������ �����
#         + ������� �������� �� ����� (�� �����, ������� ������ 2009 ���� ��� 2007)
#         + ��������� �������� ������ �� �������� �����
#         + ���������� � ����������� ���������� �����������
#         + ����� �� �����
#         + �������� �� ������ 2008 ���� ������� 2007 ���� - ����������� ��
#           �����������, � ������ �����
#         + ��������� � ���������� ����� ���������������� � 2007 ���� �� ������� �����
#         + ������� ������ � Excel-����� �� ����������� � ������� �� �����-2007
#         - ������� ��� 2008 ���� ������������������
#           + ���������� �����. ��������� ��� ������ ������ 2009 ���� � �������� 2008 ����
#           + ����� ����� 2008 ���� (��������� ������)
#           - ���������� ����������������� �������� ������ 2008 ���� �� �������� ������ 2008 ����� �����.
#           - ������� ������ � Excel-����� �� ����������� 2008 ���� � ������� �����-2007 (������������)
#         
#         
#         
############################################
#       �������� �.�.   15.11.2010 20:57:45
############################################
from __future__ import with_statement
import os
from datetime import datetime, date, time
from time import *
from types import *
from xlwt import *
import xlrd

newDir         = 'e:\\Tmp\\Spss\\rowsratio'
dataDir        = newDir + '\\spssData'
dataDirXls     = newDir + '\\xlsDataAll'
totalDir       = '\\Total'
totalDir2007   = newDir + '\\����� �� 2007\\'
totalDir2008   = newDir + '\\total2008\\'
ratioDir       = newDir + '\Ratio\\'
ListTOGS       = '\\ListTOGS'
nameListTOGS   = 'TOGS'
sav            = '.sav'
#mal            = '_��'
#mic            = '_��(�����)'
mal            = 'p'
mic            = 'mi'
y2008          = '2008'
sMesFileNo     = '����������� ���� � �������: '
sMesFileYes    = '���� ���� � ������� : '

os.chdir(newDir)
sys.path.append(newDir)

_numvesmal        = 33          # ������ ���� � ����� = ����� ���������� - 1
_numvesmic        = 30          # ������ ���� � �����

#---  Defining Classes  ---
from OkvedAssemblyClass import *
from ListOkvedClass import *
#from cTerrClass import *
class cTerr(object):

    def __init__ (self, dataDir, mal, mic, nO, sT, kO, sO):
        #print globals()
        self.OKATO = "%02i" % nO        # ����� 
        self.sTerr = sT.strip()         # �������� ����������
        self.nOkrug = "%02i" % kO       # ��� ������
        self.sOkrug = sO                # �������� ������
        self.malFile = dataDir + '\\' + self.OKATO + mal + '.sav'    # ��� ����� ������ ����� �����������.
        self.micFile = dataDir + '\\' + self.OKATO + mic + '.sav'    # ��� ����� ������ ����������������.
                                             # �.� �� --- (03)������������� ����_1.xls
                                             # �.� ��(�����) --- (87)���������� ����_1.xls
        self.malFileXls = dataDir + '\\' + self.OKATO+mal+'.xls'    # ��� ����� ������ ����� �����������.
        self.micFileXls = dataDir + '\\' + self.OKATO+mic+'.xls'    # ��� ����� ������ ����������������.
        self.ismalFile = False
        self.ismicFile = False
        self.nameDataSet = 'D' + self.OKATO                       # ��� DATASET

        self.malTot2088 = [0,0,0,0,0,0,0,0,0,0,0,0.0]                 # ����� �� 2008 ���� �� �������� 2008
        self.malTot2008 = [0,0,0,0,0,0,0,0,0,0,0,0.0]                 # ����� �� 2008 ���� �� �������� 2007
        self.malTot2009 = [0,0,0,0,0,0,0,0,0,0,0,0.0]                 # ����� �� 2009 ���� �� �������� 2009
        self.malTot2007 = [0,0,0,0,0,0,0,0,0,0,0,0.0]                 # ����� �� 2007 ����
        self.malTos2007 = [0,0,0,0,0,0,0,0,0,0,0,0.0]                 # ����� �� 2007 ���� ����������������� 
        self.malTot2__8 = [0,0,0,0,0,0,0,0,0,0,0,0.0]                 # ����� �� 2008 ����
        self.malTos2__8 = [0,0,0,0,0,0,0,0,0,0,0,0.0]                 # ����� �� 2008 ���� ����������������� 
        self.malcoeff   = [1,1,1,1,1,1,1,1,1,1,1,1.0]                 # ������������ ��������� � 2009/2007
        self.malcoeff8   = [1,1,1,1,1,1,1,1,1,1,1,1.0]                # ������������ ��������� � 2009/2008

        self.malTot2888Okv = []                 # ����� ��� �� 2008 ���� ������ �� �����
        self.malTot2008Okv = []                 # ����� ��� �� 2008 ���� ������ �� ����� �� �������� 2007
        self.malTot2009Okv = []                 # ����� ��� 2008 �� �������� 2009 ���� ������ �� �����
        self.malTot2007Okv = []                 # ����� �� 2007 ���� ������ �� ����� ��������
        self.malTos2007Okv = []                 # ����� �� 2007 ���� ����������������� ������ �� ����� 
        self.malTot2088Okv = []                 # ����� �� 2008 ���� ������ �� ����� ��������
        self.malTos2088Okv = []                 # ����� �� 2008 ���� ����������������� ������ �� ����� 
        self.malcoeffOkv   = []                 # ������������ ��������� ������ �� ����� � 2009/2007
        self.malcoeff8Okv   = []                # ������������ ��������� ������ �� ����� � 2009/2008

        self.malnnab     = 0                                      # ���������� ����������
        self.malknab     = 0                                      # ���������� ���������� ����������
        self.malrnab     = 0                                      # ���������� ��������� ����������
        self.numvesmal   = 33           # ������ ���� � ����� = ����� ���������� - 1
        self.numvesmalXls   = 15           # ������ ���� � ����� = ����� ���������� - 1
        self.malCHT      = 5            # ������ ������� ����������� � �������� ����� SPSS
        self.malVIR      = 8            # ������ ������� � �������� ����� SPSS
        self.malnCHT      = 2            # ������ ������� ����������� � �����. ������� ��������� ������
        self.malnVIR      = 3            # ������ ������� � �����. ������� ��������� ������
        self.malCHTres   = 100          # ����������� ������� ����������� �����
        self.malVIRres   = 400000000    # ����������� ������� �����

        self.micRatio2007 = [0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0]             # ���� ���������������� �� 2007 ����
        self.malRatio2007 = [0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0]             # ���� ����� ����������� �� 2007 ����
        self.micRatio2007Okv = []             # ���� ���������������� �� 2007 ���� ������ �� �����
        self.malRatio2007Okv = []             # ���� ����� ����������� �� 2007 ���� ������ �� �����

        self.micTot2088 = [0,0,0,0,0,0,0,0,0,0,0.0]                 # ����� �� 2008 ���� �� �������� 2008
        self.micTot2008 = [0,0,0,0,0,0,0,0,0,0,0.0]                 # ����� �� 2008 ���� �� �������� 2007
        self.micTot2009 = [0,0,0,0,0,0,0,0,0,0,0.0]                 # ����� �� 2009 ���� �� �������� 2009
        self.micTot2007 = [0,0,0,0,0,0,0,0,0,0,0.0]                 # ����� �� 2007 ����
        self.micTos2007 = [0,0,0,0,0,0,0,0,0,0,0.0]                 # ����� �� 2007 ���� �����������������
        self.micTot2__8 = [0,0,0,0,0,0,0,0,0,0,0.0]                 # ����� �� 2008 ����
        self.micTos2__8 = [0,0,0,0,0,0,0,0,0,0,0.0]                 # ����� �� 2008 ���� �����������������
        self.miccoeff   = [1,1,1,1,1,1,1,1,1,1,1.0]                 # ������������ ���������� 2009/2007
        self.miccoeff8   = [1,1,1,1,1,1,1,1,1,1,1,1.0]              # ������������ ��������� � 2009/2008

        self.micTot2888Okv = []                 # ����� ��� �� 2008 ���� ������ �� �����
        self.micTot2008Okv = []                 # ����� ��� �� 2008 ���� ������ �� ����� �� �������� 2007
        self.micTot2009Okv = []                 # ����� ��� 2008 �� �������� 2009 ���� ������ �� �����
        self.micTot2007Okv = []                 # ����� �� 2007 ���� ������ �� ����� ��������
        self.micTos2007Okv = []                 # ����� �� 2007 ���� ����������������� ������ �� ����� 
        self.micTot2088Okv = []                 # ����� �� 2008 ���� ������ �� ����� ��������
        self.micTos2088Okv = []                 # ����� �� 2008 ���� ����������������� ������ �� ����� 
        self.miccoeffOkv   = []                 # ������������ ��������� ������ �� ����� � 2007 �����
        self.miccoeff8Okv   = []                # ������������ ��������� ������ �� ����� � 2008

        self.micnnab     = 0                                      # ���������� ����������
        self.micknab     = 0                                      # ���������� ���������� ����������
        self.micrnab     = 0                                      # ���������� ��������� ����������
        self.numvesmic   = 30          # ������ ���� � ����� = ����� ���������� - 1
        self.numvesmicXls   = 15          # ������ ���� � ����� = ����� ���������� - 1
        self.micCHT      = 5            # ������ ������� ����������� � �������� ����� SPSS
        self.micVIR      = 8            # ������ ������� � �������� ����� SPSS
        self.micnCHT      = 2            # ������ ������� ����������� � �����. ������� ��������� ������
        self.micnVIR      = 3            # ������ ������� � �����. ������� ��������� ������
        self.micCHTres   = 15           # ����������� ������� ����������� ����������������
        self.micVIRres   = 60000000     # ����������� ������� ����������������

    def evalTotal(self, vM, mFile, listOkved, rulOk, \
                        Tot2088, Tot2008, Tot2009, coeff, coeff8, \
                        Tot2888okv, Tot2008okv, Tot2009okv, coeffokv, coeff8okv, \
                        Tot2007okv, Tos2007okv, Ratio2007okv, Tot2088okv, Tos2088okv, \
                        numves, kCHT, kVIR, CHTres, VIRres, npok):

        def getDataXls(self):
            if os.access(mFile, os.F_OK):
                rb = xlrd.open_workbook(mFile,formatting_info=True,encoding_override="cp1251")
                sh = rb.sheet_by_index(0)
                dd = []                                     # �������� ������� ������ � ����� Excel
                nFirstRow = 1                               # ������ ������ � ������� � ����� ����������
                for rownum in range(nFirstRow, sh.nrows):
                    rr = []                                 # ������� ������ �������
                    for i, np in enumerate(npok):
                        if i == 0:
                            s = str(sh.cell_value(rowx=rownum, colx=np[4])).encode('cp1251').strip()
                            if len(s.split('.',1)[0]) == 1:
                                s = '0'+s
                            rr.append(s)
                        else:
                            rr.append(missFloat(sh.cell_value(rowx=rownum, colx=np[4])))  # �������� ������ �������
                    rr.insert(1,missFloat(sh.cell_value(rowx=rownum, colx=numves)))
                    dd.append(rr)                           # ��������� ��������� ������

                del rb
            else:
                ToPrintLog('�� ������!! ���� ������: ' + mFile)
                dd = []
                dd.append(['00.00.00',0,0,0,0,0,0])
            return dd

        stime = time()
        data2008 = getDataXls(self)
#        for k in range(0,len(data2008)):
#            print k, data2008[k]
#            data2008[k][0] = data2008[k][0].strip()
#            print data2008[k][0], type(data2008[k][0]), data2008[k][1], type(data2008[k][1])
# �������: ����� ��� ������� ������� ���������� ������ ��������� ������� ������_���_������� ...
#            0    1     2       3        4         5       6        7             8        
        data  = [rw for rw in data2008 if rw[1] > 0 \
                and ((rw[0] in rulOk['PROM'] and rw[2] <= 100) \
                  or (rw[0] in rulOk['STRO'] and rw[2] <= 100) \
                  or (rw[0] in rulOk['TRAN'] and rw[2] <= 100) \
                  or (rw[0] in rulOk['SECH'] and rw[2] <=  60) \
                  or (rw[0] in rulOk['NTSF'] and rw[2] <=  60) \
                  or (rw[0] in rulOk['OPTO'] and rw[2] <=  50) \
                  or (rw[0] in rulOk['ROTO'] and rw[2] <=  30) \
                  or (rw[0] not in rulOk['PROC'] and rw[2] <= 100) \
                    )]                                                   # �� �������� 2007 ����
        ToPrintLog('����� � ������� 2008 ����                             :'+str(len(data2008))+' �����������')
        ToPrintLog('�������� � ������� 2008 �� �������� 2007 ����         :'+str(len(data))+' �����������')
        ToPrintLog('����� ��������� �� ������� 2008 �� �������� 2007 ���� :'+str(len(data2008)-len(data))+' �����������')
        ToPrintLog('    �� ��� � �������� ������                          :'+str(len([rw for rw in data2008 if rw[1] <= 0])))
#        for rw in data:
#            print rw
        data9 = [rw for rw in data2008 if (rw[1] > 0 \
                and rw[kCHT] <= CHTres and rw[kVIR] <= VIRres)]          # �� �������� 2009 ����

        ToPrintLog('�������� � ������� 2008 �� �������� 2009 ����              :'+str(len(data9))+' �����������')
        ToPrintLog('����� ��������� �� ������� 2008 �� �������� 2009 ����      :'+str(len(data2008)-len(data9))+' �����������')
        ToPrintLog('    �� ��� � �������� ������                               :'+str(len([rw for rw in data2008 if rw[1] <= 0])))

        i = len([rv[0] for rv in data])                                  # ���������� �������� ����������
        i9 = len([rv[0] for rv in data9])                                # ���������� �������� ����������
        Tot2088[0] = '�����'                                             # ��� �����
        Tot2008[0] = '�����'                                             # ��� �����
        Tot2009[0] = '�����'                                             # ��� �����
        Tot2088[1] = sum(rv[1] for rv in data2008)                       # ������ ���������� ����������� �� ���� �� �������� 2008 ����
        Tot2008[1] = sum(rv[1] for rv in data)                           # ������ ���������� ����������� �� ���� �� �������� 2007 ����
        Tot2009[1] = sum(rv[1] for rv in data9)                          # ������ ���������� ����������� �� ���� �� �������� 2009 ����
        
        ToPrintLog('������ �� �������� 2008 ����(��������)        :'+str(Tot2088[1])+' �����������')
        ToPrintLog('������ �� �������� 2007 ����                  :'+str(Tot2008[1])+' �����������')
        ToPrintLog('������ �� �������� 2009 ����                  :'+str(Tot2009[1])+' �����������')
        
        for i in range(2, len(Tot2008)):                                 # �� �����������
            Tot2088[i]  = sum(rv[1] * rv[i] for rv in data2008)
            Tot2008[i]  = sum(rv[1] * rv[i] for rv in data)
            Tot2009[i]  = sum(rv[1] * rv[i] for rv in data9)

        k = 0
        for rokv in listOkved:                                           # ���������� ������ �� ����� (�� 2007 - ����� ������)
#            print rokv[0], rokv[1], rokv[2]
            Tot2888okv.append([0.0 for rw in Tot2088])
            Tot2008okv.append([0.0 for rw in Tot2008])
            Tot2009okv.append([0.0 for rw in Tot2009])
            coeffokv.append([0.0 for rw in coeff])
            coeff8okv.append([0.0 for rw in coeff])
            
            Tot2007okv.append([0.0 for rw in Tot2008])
            Tos2007okv.append([0.0 for rw in Tot2008])
            Tot2088okv.append([0.0 for rw in Tot2008])
            Tos2088okv.append([0.0 for rw in Tot2008])
            Ratio2007okv.append([0.0 for rw in self.micRatio2007])
            
            Tot2008okv[k][0] = rokv[0]
            Tot2009okv[k][0] = rokv[0]
            coeffokv[k][0]   = rokv[0]
            Tot2888okv[k][0] = rokv[0]
            coeff8okv[k][0]   = rokv[0]

            Tot2007okv[k][0] = rokv[0]
            Tos2007okv[k][0] = rokv[0]
            Tot2088okv[k][0] = rokv[0]
            Tos2088okv[k][0] = rokv[0]
            Ratio2007okv[k][0] = rokv[0]

#            print [rv[0] for rv in data if rv[0] in rokv[2]]
            Tot2888okv[k][1] = sum(rv[1] for rv in data2008 if rv[0] in rokv[2])
            Tot2008okv[k][1] = sum(rv[1] for rv in data if rv[0] in rokv[2])
            Tot2009okv[k][1] = sum(rv[1] for rv in data9 if rv[0] in rokv[2])
            for i in range(2, len(Tot2008)):                                      # �� �����������
                Tot2888okv[k][i]  = sum(rv[1] * rv[i] for rv in data2008 if rv[0] in rokv[2])
                Tot2008okv[k][i]  = sum(rv[1] * rv[i] for rv in data if rv[0] in rokv[2])
                Tot2009okv[k][i]  = sum(rv[1] * rv[i] for rv in data9 if rv[0] in rokv[2])

            k += 1
#
#            print Tot2008okv[0], Tot2008okv[1], Tot2008okv[2]
#
        print "�� �����: ", [rv[0] for rv in data if rv[0] not in listOkved[0][2]]
#        for i in range(0, len(Tot2008okv)):
#            print Tot2008okv[i]
#            print Tot2009okv[i]

        
        for itot in range(1, len(Tot2009)):         # ��� 2007 ����
            if Tot2008[itot] > 0:
                coeff[itot] = Tot2009[itot] / Tot2008[itot]
            else:
                coeff[itot] = 1
            if Tot2088[itot] > 0:
                coeff8[itot] = Tot2009[itot] / Tot2088[itot]
            else:
                coeff8[itot] = 1
            #print itot, self.malcoeff[itot], self.malTot2009[itot], self.malTot2008[itot]

        for ir in range(0, len(Tot2008okv)):         # ��� 2007  � 2008 ���� �� �����
            for itot in range(1, len(Tot2009)):
                if Tot2008okv[ir][itot] > 0:
                    coeffokv[ir][itot] = Tot2009okv[ir][itot] / Tot2008okv[ir][itot]
                else:
                    coeffokv[ir][itot] = 1
            for itot in range(1, len(Tot2009)):
                if Tot2888okv[ir][itot] > 0:
                    coeff8okv[ir][itot] = Tot2009okv[ir][itot] / Tot2888okv[ir][itot]
                else:
                    coeff8okv[ir][itot] = 1

#        for i in range(0, len(coeffokv)):
#            print coeffokv[i]
        etime = time() - stime
        ToPrintLog('..����� ������� ������ �� '+ mFile.rsplit('\\',1)[1] + ' : %.2f ���.' % (etime))
        

#class rTerr:
#    def __init__ (self, nO, sT, kO, sO):
#        self.OKATO = "%02i" % nO     # ����� 
#        self.sTerr = sT     # �������� ����������
#        self.nOkrug = "%02i" % kO    # ��� ������
#        self.sOkrug = sO    # �������� ������

class TTotal(cTerr):
    aTerr         = []
    rTerr         = []
    numpokmal     = 12
    numpokmic     = 11
    vMikro        = False
    malpok = [
    [1  , '���������� ����������� �� ������-�������, ���� ',                                                                                'tab_33pred-2007.xls',     'tab_33pred-corr'         ,2 ,'tabl_33_1401(1)_1351000010001-2008.xls',   ''],
    [5  , 'C������ ����������� ���������� - �����,���.',                                                                                    'tab_33(01-09)_list1.xls', 'tab_33(01-09)_list1-corr',4 ,'���_tabl_33_1401(1)_1121100010001.xls',    ''],
    [8  , '������� (�����) �� ���������� �������, ���������, �����, ����� (��� ���, ������� � ����������� ������������ ��������),���.���.', 'tab_33(12)_list1.xls',    'tab_33(12)_list1-corr'   ,13,'���_tabl_33_1402(1)_2113100020001.xls',    ''],
    [11 , '���������� � �������� ������� (� ����� ����� � ������������� �� ������� �������� �������),���.���.',                             'tab_33(13)_list1.xls',    'tab_33(13)_list1-corr'   ,14,'���_tabl_33_1402(1)_1141100020001.xls',    ''],
    [14 , '������ ����������� ,���.���.',                                                                                                   'tab_33(o)_list1.xls',     'tab_33(o)_list1-corr'    ,10,'����_tabl_33_1402(1)_211330002000902.xls', ''],
    [17 , '��������� ������� ������������ ������������, ��������� ����� � ����� ������������ ������ (��� ��� � ������),���.���.',           'tab_33(06)_list1.xls',    'tab_33(06)_list1-corr'   ,11,'���_tabl_33_1402(1)_2131100020021.xls',    ''],
    [20 , '������� ������� �������������� ������������ (��� ��� � �������),���.���.',                                                       'tab_33(07)_list1.xls',    'tab_33(07)_list1-corr'   ,12,'����_tabl_33_1402(1)_2135000020005.xls',   ''],
    [23 , '������� ����������� ���������� ���������� ������� (��� ������� �������������),���.',                                             'tab_33(02-09)_list1.xls', 'tab_33(02-09)_list1-corr',5 ,'����_tabl_33_1401(1)_1121100010003.xls',   ''],
    [26 , '���� ����������� ���������� ����� ������� �������������,���.���.',                                                               'tab_33(03-11)_list1.xls', 'tab_33(03-11)_list1-corr',9 ,'tabl_33_1401(1)_2152320020005.xls',''],
    [29 , '���� ����������� ���������� ����� ���� ����������,���.���.',                                                                     'tab_33(01-11)_list1.xls', 'tab_33(01-11)_list1-corr',6 ,'tabl_33_1401(1)_2152320020022.xls',''],
    [32 , '���� ����������� ���������� ����� ���������� ���������� ������� (��� ������� �������������),���.���.',                           'tab_33(02-11)_list1.xls', 'tab_33(02-11)_list1-corr',8 ,'tabl_33_1401(1)_2152320020003.xls','']
    ]
    micpok = [
    [1  , '���������� ����������� �� ������-�������, ���� ',                                                                                                           'tab_33pred-2007.xls',     'tab_33pred-corr'         ,2 ,'���������������2008.xls',''],
    [5  , '������� ����������� ���������� (������� ����������� ������ �� ��������� ����������-��������� ���������),���.',                                              'tab_33(01-09)_list1.xls', 'tab_33(01-09)_list1-corr',4 ,'��� � �� 2008.xls',      ''],
    [8  , '������� (�����) �� ������� �������, ���������, �����, ����� (��� ���, ������� � ����������� ������������ ��������),���.���.',                               'tab_33(12)_list1.xls',    'tab_33(12)_list1-corr'   ,13,'�������2008.xls',''],
    [11 , '���������� � �������� ������� (� ����� ����� � ������������ �� ������� �������� �������) - �����,���.���.',                                                'tab_33(13)_list1.xls',    'tab_33(13)_list1-corr'   ,14,'����������2008.xls',''],
    [14 , '������ �����������,���.���.',                                                                                                                               'tab_33(o)_list1.xls',     'tab_33(o)_list1-corr'    ,10,'������ � �� 2008.xls',''],
    [17 , '��������� ������� ������������ ������������, ��������� ����� � ����� ������������ ������ (��� ���, ������� � ����������� ������������ ��������),���.���.',  'tab_33(06)_list1.xls',    'tab_33(06)_list1-corr'   ,11,'�������� � �� 2008.xls',''],
    [20 , '������� ������� �������������� ������������ (��� ���, ������� � ����������� ������������ ��������),���.���.',                                               'tab_33(07)_list1.xls',    'tab_33(07)_list1-corr'   ,12,'�������  2008.xls',''],
    [23 , '������� ����������� ���������� ���������� ������� (��� ������� �������������),���.',                                                                        'tab_33(02-09)_list1.xls', 'tab_33(02-09)_list1-corr',5 ,'��� ��� ���� 2008.xls',''],
    [26 , '���� ����������� ���������� ����� ����������,���.���.',                                                                                                     'tab_33(01-11)_list1.xls', 'tab_33(01-11)_list1-corr',6 ,'��� �����2008.xls',''],
    [29 , '���� ����������� ���������� ����� ���������� ���������� ������� � ������� �������������,���.���.',                                                          'tab_33(02-11)_list1.xls', 'tab_33(02-11)_list1-corr',7 ,'��� ���� ����2008.xls','']
    ]

    def __init__ (self):
        self.OkRul = OkvedAss(ToPrintLog, newDir)           # OkRul.rulesAssPlus
        self.ListOkv = ListOkved(ToPrintLog, newDir)        # self.ListOkv.listOkvedmal[][]
    def getTotalPred(self, totalDir, vMicro, mpok, npFile, strtColTerr, colOkv):
        def setmalTot(valu, ggt, rrt, nnp):
            if   ggt == '7':
                rrt.malTot2007[nnp] = valu
            elif ggt == '8':
                if vMicro:
                    rrt.micTot2__8[nnp] = valu
                else:
                    rrt.malTot2__8[nnp] = valu
            
        def setmalTotOkv(valu, ggt, rrt, nnp, kk):
            if   ggt == '7':
                rrt.malTot2007Okv[kk][nnp] = valu
            elif ggt == '8':
                if vMicro:
                    rrt.micTot2088Okv[kk][nnp] = valu
                else:
                    rrt.malTot2088Okv[kk][nnp] = valu
        
        np = 0
        if '7' in totalDir: gt = '7'
        if '8' in totalDir: gt = '8'
        if '9' in totalDir: gt = '9'
#        npFile = 2              # ������ ����� ����� �����
#        strtColTerr = 2         # 
        os.chdir(totalDir)
        for pok in mpok:                 # �� �������� �����������
            np = np + 1                         # ����� ���������� � ������ ������ malTot2007
            if os.access(pok[npFile], os.F_OK):      # ���� �� ���� ����������
                rb = xlrd.open_workbook(totalDir + pok[npFile],formatting_info=True,encoding_override="cp1251")
                sh = rb.sheet_by_index(0)
                ToPrintLog('������ ���� ������ ' + pok[npFile]+' � '+str(sh.ncols)+' ��������.')
                for rt in self.aTerr:                       # ��������� ����������
                    vTerr = True
                    for icol in range(strtColTerr, sh.ncols):       # �������� ���������� � ��������!
                        sName = sh.cell_value(rowx=4, colx=icol).encode('cp1251').strip()
#                        print icol, sName, rt.sTerr, sName.strip() == rt.sTerr.strip()
                        if sName == rt.sTerr.strip() or \
                           rt.OKATO == '79' and '������'    in sName or \
                           rt.OKATO == '82' and '������'    in sName or \
                           rt.OKATO == '92' and '���������' in sName or \
                           rt.OKATO == '97' and '�����'     in sName or \
                           rt.OKATO == '30' and '��������'  in sName or \
                           rt.OKATO == '77' and '�����'     in sName or \
                           rt.OKATO == '99' and '�������'   in sName or \
                           rt.OKATO == '80' and '���������' in sName :       # ���� ����������!
                            vTerr = False
                            value = sh.cell_value(rowx=6, colx=icol)
                            setmalTot(value, gt, rt, np)              # ��������� ���� ���� �������
                            ToPrintLog('�������: ' +' '+rt.OKATO+' '+ sName + ' ' + pok[1][0:17] + ' �����: ' + str(value))
                            k = 0
                            for rokv in rt.malTot2007Okv:   # ����� 2007(2008) �� �����-2007 - ����� ������
                                compp = False
                                for i in range(6, sh.nrows):    # ���� ��� ����� � �����
                                    okvFile = sh.cell_value(rowx=i, colx=colOkv).encode('cp1251').strip()
                                    if rt.malTot2007Okv[k][0] == okvFile or \
                                       rt.malTot2007Okv[k][0] == '00' and okvFile == 'RR' or \
                                       rt.malTot2007Okv[k][0] == '44.00.9' and okvFile == '44.00.09' :   # ���� ������ ����!
                                        compp = True
                                        okvTot = sh.cell_value(rowx=i, colx=icol)
                                        if okvTot == u'-':
                                            setmalTotOkv(0.0, gt, rt, np, k)
                                        else:
#                                            rt.malTot2007Okv[k][np] = float(okvTot)
                                            setmalTotOkv(float(okvTot), gt, rt, np, k)
                                        break
#                                    else:
                                if not compp:
                                    ToPrintLog('    ����������� ����� '+rt.malTot2007Okv[k][0]+' � ����� ������.')
                                k += 1
                            break
#                            for i in range(0, len(rt.malTot2007Okv)):
#                                print rt.malTot2007Okv[i]
                        else:
                            pass
                    if vTerr:
                        ToPrintLog('    ����������� ���������� '+' '+rt.OKATO+' '+rt.sTerr+' � ����� ������.')
            else: 
                ToPrintLog('�� ������!! ���� ������ ' + pok[2])

    def getRatioMic(self):

        def getProd(self, rP, stro, scol, rcol):
            compp = False
            for rownum in range(sh.nrows):
                parsName = sh.cell_value(rowx=rownum, colx=scol).encode('cp1251')
                #print type(parsName), parsName.strip(), parsName.find(stro)
                if parsName.find(stro) >= 0:
                    compp = True
                    valRatio = sh.cell_value(rowx=rownum, colx=rcol)
#                    ToPrintLog('������� ���� :'+stro+':'+parsName+': '+str(valRatio))
                    if type(valRatio) is FloatType : 
                        valRatio = float(valRatio) 
                    else: 
                        valRatio = 0.0
                    return valRatio
            else:
                pass
#                if not compp:
#                    ToPrintLog('    �� ������� ���� �� ����� -'+stro+'-')
            return rP

        os.chdir(ratioDir)
        for rt in self.aTerr:
            if os.access(rt.OKATO + ' ����.xls', os.F_OK):
                rb = xlrd.open_workbook(ratioDir + rt.OKATO + ' ����.xls',formatting_info=True,encoding_override="cp1251")
                sh = rb.sheet_by_index(0)
                rPred  = sh.cell_value(rowx=3, colx=2)
                ToPrintLog('������ ���� ����� ��� ' + ' ' + rt.OKATO + ' ' + rt.sTerr.strip() + ' ' + str(rPred))
                rProd  = getProd(self, rPred, '������� �������', 0, 2)
                rOtgr  = getProd(self, rPred, '��������� �������', 0, 2)
                
                rt.micRatio2007[0 ]  =  0
                rt.micRatio2007[1 ]  =  rPred
                rt.micRatio2007[2 ]  =  rPred
                rt.micRatio2007[3 ]  =  rProd
                rt.micRatio2007[4 ]  =  rProd
                rt.micRatio2007[5 ]  =  rProd
                rt.micRatio2007[6 ]  =  rOtgr
                rt.micRatio2007[7 ]  =  rProd
                rt.micRatio2007[8 ]  =  rPred
                rt.micRatio2007[9 ]  =  rPred
                rt.micRatio2007[10]  =  rPred
                rt.micRatio2007[11]  =  rPred

                # ������ �� �����
                k = 0
#                print '����: ', len(rt.micRatio2007Okv), '����� �� ', len(rt.micRatio2007Okv[0]), ' �����.'
                for rokv in rt.micRatio2007Okv:   # ���� ���������������� � 2007 �� �����-2007 - ����� ������
                    rProk = getProd(self, rPred, rokv[0], 0, 3)
#                    print '���� ��� ', k, rokv[0],
                    rt.micRatio2007Okv[k][1 ] = rProk
                    rt.micRatio2007Okv[k][2 ] = rProk
                    rt.micRatio2007Okv[k][3 ] = rProd
                    rt.micRatio2007Okv[k][4 ] = rProd
                    rt.micRatio2007Okv[k][5 ] = rProd
                    rt.micRatio2007Okv[k][6 ] = rOtgr
                    rt.micRatio2007Okv[k][7 ] = rProd
                    rt.micRatio2007Okv[k][8 ] = rProk
                    rt.micRatio2007Okv[k][9 ] = rProk
                    rt.micRatio2007Okv[k][10] = rProk
                    rt.micRatio2007Okv[k][11] = rProk
                    k += 1
            else:
                ToPrintLog('�� ������!! ���� ����� ��� ' + rt.OKATO + ' ' + rt.sTerr.strip())

            for np in range(0, len(rt.micRatio2007)):              # ����� �����, ��������� ��� ��������� �� ����� ��� 2007
                rt.malRatio2007[np]  = 1 - rt.micRatio2007[np]          # ���� ����� �� �����
                rt.malTos2007[np] = rt.malTot2007[np] * rt.malRatio2007[np] * rt.malcoeff[np]
            for np in range(0, len(rt.micTot2007)):
                rt.micTot2007[np] = rt.malTot2007[np]
                rt.micTos2007[np] = rt.malTot2007[np] * rt.micRatio2007[np] * rt.miccoeff[np]
            rt.micTot2007[9] = rt.malTot2007[11]
            rt.micTos2007[9] = rt.malTot2007[11] * rt.micRatio2007[11] * rt.miccoeff[9]
            
            for np in range(0, len(rt.malTot2__8)):              # ����� �����, ��������� ��� ��������� �� ����� ��� 2008 ����
                rt.malTos2__8[np] = rt.malTot2__8[np] * rt.malcoeff8[np]
            for np in range(0, len(rt.micTot2__8)):
                rt.micTos2__8[np] = rt.micTot2__8[np] * rt.miccoeff8[np]
#            rt.micTot2__8[9] = rt.micTot2__8[11]
#            rt.micTos2__8[9] = rt.micTot2__8[11] * rt.miccoeff8[9]
                
            for i, rat in enumerate(rt.micRatio2007Okv):           # ����� ��������� �� ����� ��� 2007 ����
                for j in range(1, len(rat)) :
                    rt.malRatio2007Okv[i][j] = 1-rt.micRatio2007Okv[i][j]
                    rt.malTos2007Okv[i][j]   = rt.malTot2007Okv[i][j]*rt.malRatio2007Okv[i][j]*rt.malcoeffOkv[i][j]
                for j in range(1, len(rt.micTot2007Okv[0])) :
                    rt.micTot2007Okv[i][j] = rt.malTot2007Okv[i][j]
                    rt.micTos2007Okv[i][j]   = rt.malTot2007Okv[i][j]*rt.micRatio2007Okv[i][j]*rt.miccoeffOkv[i][j]
                rt.micTot2007Okv[i][9] = rt.malTot2007Okv[i][11]
                rt.micTos2007Okv[i][9] = rt.malTot2007Okv[i][11] * rt.micRatio2007Okv[i][11] * rt.miccoeffOkv[i][9]

            for i, rat in enumerate(rt.malTot2088Okv):           # ����� ��������� �� ����� ��� 2008 ����
                for j in range(1, len(rat)):
                    rt.malTos2088Okv[i][j] = rt.malTot2088Okv[i][j] * rt.malcoeff8Okv[i][j]
                for j in range(1, len(rt.micTot2088Okv[0])):
                    rt.micTos2088Okv[i][j] = rt.micTot2088Okv[i][j] * rt.miccoeff8Okv[i][j]
#                print i, rat
#                print i, rt.micTos2088Okv[i]
#                print i, rt.micTot2088Okv[i]
#                print i, rt.miccoeff8Okv[i]
            self.sumTotalRR(rt.malTos2007Okv, 0)
            self.sumTotalRR(rt.micTos2007Okv, 0)
            self.sumTotalRR(rt.malTos2088Okv, 0)
            self.sumTotalRR(rt.micTos2088Okv, 0)

    def sumTotalRR(self, Tos, nt):
        for i in range(1, len(Tos[nt])):                             # �������� �������� ����
            Tos[nt][i] = 0.0
            ssum = sum(rv[i] for rv in Tos)
            Tos[nt][i] = ssum

    def allTerr(self):
        for rt in self.aTerr:
            stime = time()
            ToPrintLog('....'+rt.nOkrug+' '+rt.sOkrug+' '+rt.OKATO+' '+rt.sTerr.strip()+' ����� �����������')
            self.vMikro = False
            rt.evalTotal(self.vMikro, rt.malFileXls, self.ListOkv.listOkvedmal, self.OkRul.rulesAssPlus,
            rt.malTot2088, rt.malTot2008, rt.malTot2009, rt.malcoeff, rt.malcoeff8,
            rt.malTot2888Okv, rt.malTot2008Okv, rt.malTot2009Okv, rt.malcoeffOkv, rt.malcoeff8Okv,
            rt.malTot2007Okv, rt.malTos2007Okv, rt.malRatio2007Okv, rt.malTot2088Okv, rt.malTos2088Okv, rt.numvesmalXls,
            rt.malnCHT, rt.malnVIR, rt.malCHTres, rt.malVIRres, self.malpok)

            ToPrintLog('....'+rt.nOkrug+' '+rt.sOkrug+' '+rt.OKATO+' '+rt.sTerr.strip()+' ����� �����������')
            self.vMikro = True
            rt.evalTotal(self.vMikro, rt.micFileXls, self.ListOkv.listOkvedmic, self.OkRul.rulesAssPlus,
            rt.micTot2088, rt.micTot2008, rt.micTot2009, rt.miccoeff, rt.miccoeff8,
            rt.micTot2888Okv, rt.micTot2008Okv, rt.micTot2009Okv, rt.miccoeffOkv, rt.miccoeff8Okv,
            rt.micTot2007Okv, rt.micTos2007Okv, rt.micRatio2007Okv, rt.micTot2088Okv, rt.micTos2088Okv, rt.numvesmicXls,
            rt.micnCHT, rt.micnVIR, rt.micCHTres, rt.micVIRres, self.micpok)

            ToPrintLog('        |����|       |            | �����������  | ����������� ')
            ToPrintLog('��08 �����           |{0:12}|{1:13.0f} |{2:13.0f}'.format(rt.malTot2088[0],rt.malTot2088[1],rt.malTot2088[2]))
            ToPrintLog('��07 ���' + '|{0:4}|{1:7d}|{2:12}|{3:13.0f} |{4:13.0f}'.format(rt.OKATO,rt.malnnab,rt.malTot2008[0],rt.malTot2008[1], rt.malTot2008[2]))
            ToPrintLog('��09 �����           |{0:12}|{1:13.0f} |{2:13.0f}'.format(rt.malTot2009[0],rt.malTot2009[1],rt.malTot2009[2]))
            ToPrintLog('   ����� 09/07 ����� |{0:12}|{1:13.4f} |{2:13.4f}'.format(rt.malcoeff[0],rt.malcoeff[1],rt.malcoeff[2]))
            ToPrintLog('   ����� 09/08 ����� |{0:12}|{1:13.4f} |{2:13.4f}'.format(rt.malcoeff8[0],rt.malcoeff8[1],rt.malcoeff8[2]))

            ToPrintLog('��08 �����           |{0:12}|{1:13.0f} |{2:13.0f}'.format(rt.micTot2088[0],rt.micTot2088[1],rt.micTot2088[2]))
            ToPrintLog('��07 ���' + '|{0:4}|{1:7d}|{2:12}|{3:13.0f} |{4:13.0f}'.format(rt.OKATO,rt.malnnab,rt.micTot2008[0],rt.micTot2008[1], rt.micTot2008[2]))
            ToPrintLog('��09 �����           |{0:12}|{1:13.0f} |{2:13.0f}'.format(rt.micTot2009[0],rt.micTot2009[1],rt.micTot2009[2]))
            ToPrintLog('   ����� 09/07 ����� |{0:12}|{1:13.4f} |{2:13.4f}'.format(rt.miccoeff[0],rt.miccoeff[1],rt.miccoeff[2]))
            ToPrintLog('   ����� 09/08 ����� |{0:12}|{1:13.4f} |{2:13.4f}'.format(rt.miccoeff8[0],rt.miccoeff8[1],rt.miccoeff8[2]))

            etime = time() - stime
            ToPrintLog('....����� ������� ������ �� '+rt.OKATO+ ' : %.2f ���.' % (etime))
        
        self.getTotalPred(totalDir2007, False, self.malpok, 2, 3, 2)             # ����� ����� 2007 ����
        self.getTotalPred(totalDir2008, False, self.malpok, 5, 2, 0)             # ����� ����� 2008 ���� �����
        self.getTotalPred(totalDir2008, True,  self.micpok, 5, 2, 0)             # ����� ����� 2008 ���� �����
        self.getRatioMic()              # ����� ���� ����������������
        
    def allTerrTotalToXLS(self):
        ToPrintLog('������� ������ �� ����������� � Excel')
        dirTotal = newDir + totalDir + '\\'
        sFont = 'Calibri'
        ezxf = easyxf
        heading_xf = ezxf('font: name '+sFont+', bold on, height 230; \
        align: wrap yes, vert centre, horiz center') 
        headcol_xf = ezxf('font: name '+sFont+', bold off; \
        align: wrap yes, vert centre, horiz center; \
        borders: left 1, right 1, top 1, bottom 1') 
        headrow_xf = ezxf('font: name '+sFont+', bold off; \
        align: wrap yes, vert centre, horiz left; \
        borders:left 1, right 1, top 1, bottom 1') 
        headrowc_xf = ezxf('font: name '+sFont+', bold off; \
        align: wrap yes, vert centre, horiz centre; \
        borders:left 1, right 1, top 1, bottom 1') 
        valnrow_xf = ezxf('font: name '+sFont+', bold off; \
        align: wrap no, vert centre, horiz right; \
        borders:left 1, right 1, top 1, bottom 1', \
        num_format_str="### ### ### ###") 
        valdrow_xf = ezxf('font: name '+sFont+', bold off; \
        align: wrap no, vert centre, horiz right; \
        borders:left 1, right 1, top 1, bottom 1', \
        num_format_str="### ### ### ###.0") 
        valfrow_xf = ezxf('font: name '+sFont+', bold off; \
        align: wrap no, vert centre, horiz right; \
        borders:left 1, right 1, top 1, bottom 1', \
        num_format_str="0.0000") 

        def allTerrTotalCoeffmal(self, mic, ws, namePok, nTot, startRow):
            it = 2
            ws.write(startRow,0,nTot,headrow_xf)
            #ws.write(startRow + 1,0,"",headrow_xf)
            if mic:
                nameRatio = "    ���� ����������������"
            else:
                nameRatio = "    ���� ����� �����������"
            ws.write(startRow + 1,1,"    ���� 2008 ���� - ������� 2007 ����",headrow_xf)
            ws.write(startRow + 2,1,"    ���� 2008 ���� - ������� 2009 ����",headrow_xf)
            ws.write(startRow + 3,1,"    ����������� ��������� ����",headrow_xf)
            ws.write(startRow + 4,1,"    ����� 2007 ���� ��������",headrow_xf)
            ws.write(startRow + 5,1,nameRatio, headrow_xf)
            ws.write(startRow + 6,1,"    ����� 2007 ���� �����������������",headrow_xf)
            for rt in self.aTerr:
                if mic:
                    if nTot in (1, 2, 8):
                        ws.write(startRow + 1, it, rt.micTot2008[nTot],     valnrow_xf)  
                        ws.write(startRow + 2, it, rt.micTot2009[nTot],     valnrow_xf)  
                        ws.write(startRow + 3, it, rt.miccoeff[nTot],       valfrow_xf) 
                        ws.write(startRow + 4, it, rt.micTot2007[nTot],     valnrow_xf)  
                        ws.write(startRow + 5, it, rt.micRatio2007[nTot],   valfrow_xf)
                        ws.write(startRow + 6, it, rt.micTos2007[nTot],     valnrow_xf)  
                    else:
                        ws.write(startRow + 1, it, rt.micTot2008[nTot],     valdrow_xf)  
                        ws.write(startRow + 2, it, rt.micTot2009[nTot],     valdrow_xf)  
                        ws.write(startRow + 3, it, rt.miccoeff[nTot],       valfrow_xf) 
                        ws.write(startRow + 4, it, rt.micTot2007[nTot],     valdrow_xf)  
                        ws.write(startRow + 5, it, rt.micRatio2007[nTot],   valfrow_xf)
                        ws.write(startRow + 6, it, rt.micTos2007[nTot],     valdrow_xf)  
                else:
                    if nTot in (1, 2, 8):
                        ws.write(startRow + 1, it, rt.malTot2008[nTot],     valnrow_xf)  
                        ws.write(startRow + 2, it, rt.malTot2009[nTot],     valnrow_xf)  
                        ws.write(startRow + 3, it, rt.malcoeff[nTot],       valfrow_xf) 
                        ws.write(startRow + 4, it, rt.malTot2007[nTot],     valnrow_xf)  
                        ws.write(startRow + 5, it, rt.malRatio2007[nTot],   valfrow_xf)
                        ws.write(startRow + 6, it, rt.malTos2007[nTot],     valnrow_xf)  
                    else:
                        ws.write(startRow + 1, it, rt.malTot2008[nTot],     valdrow_xf)  
                        ws.write(startRow + 2, it, rt.malTot2009[nTot],     valdrow_xf)  
                        ws.write(startRow + 3, it, rt.malcoeff[nTot],       valfrow_xf) 
                        ws.write(startRow + 4, it, rt.malTot2007[nTot],     valdrow_xf)  
                        ws.write(startRow + 5, it, rt.malRatio2007[nTot],   valfrow_xf)
                        ws.write(startRow + 6, it, rt.malTos2007[nTot],     valdrow_xf)  
                    
                it = it + 1
            ws.write_merge(startRow, startRow, 1, it-1, namePok,headrow_xf)
            #ws.col(1).width = 0x0d00 * 3
            return startRow + 7

        def allTerrTotalTitle(self, ws, titleTab):
            ws.write_merge(0,0, 2,8, titleTab, heading_xf)
            ws.col(2).width = 0x0d00
            #ws.write_merge(lineStart, lineFinish, columnStart, columnFinish, 'text', style)
            
            ws.write(2,0,"",headcol_xf)
            ws.row(0).height = 0x0d00 + 50
            ws.write(2,1,"��� �������� ����������",headcol_xf)
            ws.write(3,0,"A",headcol_xf) 
            ws.write(3,1,"B",headcol_xf)
    
            it = 2
            for rt in self.aTerr:
#                ToPrintLog(rt.nOkrug+ '|{0:4}|'.format(rt.OKATO))
                ws.write(2,it,rt.sTerr,headcol_xf)                                # �������� ����������
                ws.write(3,it,it-1,headcol_xf)
                ws.col(it).width = 0x0d00 + 50                                    # ������ ������� � �������
                it = it + 1
            
        def allOkvTerrTotalTitle(self, ws, titleTab):
            ws.write_merge(0,2, 2,8, titleTab, heading_xf)
            #ws.write_merge(lineStart, lineFinish, columnStart, columnFinish, 'text', style)
            
#            ws.write(2,0,"",headcol_xf)
            ws.col(0).width = 0x0d00 * 3
            ws.write(3,0,"���� ������������",headcol_xf)
            ws.write(3,1,"��� ����� (������)",headcol_xf)
            ws.write(4,0,"A",headrowc_xf) 
            ws.write(4,1,"�",headrowc_xf)
    
            it = 2
            for rt in self.aTerr:
#                ToPrintLog(rt.nOkrug+ '|{0:4}|'.format(rt.OKATO))
                ws.write(3,it,rt.sTerr,headcol_xf)                                # �������� ����������
                ws.write(4,it,it-1,headrowc_xf)
                ws.col(it).width = 0x0d00 + 50                                    # ������ ������� � �������
                it = it + 1
            for i, okv in enumerate(self.ListOkv.listOkvedmal):
                ws.write(i+5, 0, okv[1], headrow_xf)
                ws.write(i+5, 1, okv[0], headrowc_xf)

            ws.col(0).width = 0x0d00 * 3
            ws.row(0).height = 0x0d00 + 50

        def allTerrCoeffmal (self, ws, npok):
            nextRow = allTerrTotalCoeffmal(self, self.vMikro, ws, "���������� �����������", 1, 4)
            for i in range(2, len(npok)):
                nextRow  = allTerrTotalCoeffmal(self, self.vMikro, ws, npok[i-1][1], i, nextRow)
    
            ws.col(1).width = 0x0d00 * 3                            # ������ ������� � ��������� ����������
        
        def allOkvTerrCorr(self, vM, sGod):
            
            def allOkvTerrStr(self, ws, vM, np):
                cl = 2
                for rt in self.aTerr:
                    if vM:
                        if    sGod == '2007':
                            Tos = rt.micTos2007Okv
                        elif sGod == '2008':
                            Tos = rt.micTos2088Okv
                    else:
                        if    sGod == '2007':
                            Tos = rt.malTos2007Okv
                        elif sGod == '2008':
                            Tos = rt.malTos2088Okv
                    for stroka, ts in enumerate(Tos):
                        i = stroka+5
                        tval = ts[np]
                        
                        if np == 0:
                            pass
                        elif np in (1, 2, 8):
                            tval = round(tval)
                            if tval <= 0.0001:
                                tval = '-'
                            ws.write(i, cl, tval, valnrow_xf)
                        else:
                            tval = round(tval,1)
                            if tval <= 0.0001:
                                tval = '-'
                            ws.write(i, cl, tval, valdrow_xf)
                    cl += 1
            #----
            if vM:
                nameSheet      = '����� �� �����������������'
                nameSuffTitle  = ' (����������������� �������� '+sGod+' ���� �� �����������������)'
                nameSuffFile   = '-�����-'+sGod
                npok = self.micpok
            else:
                nameSheet      = '����� �� �����'
                nameSuffTitle  = ' (����������������� �������� '+sGod+' ���� �� ����� ������������)'
                nameSuffFile   = '-��-'+sGod
                npok = self.malpok
            k = 0
            for np in npok:
                wbl = Workbook(encoding='cp1251')
                wsl = wbl.add_sheet(nameSheet)
                allOkvTerrTotalTitle(self, wsl, np[1]+nameSuffTitle)
                
                allOkvTerrStr(self, wsl, vM, k+1)
                
                wbl.save(dirTotal + np[3] + nameSuffFile +'.xls')                             #---  ������ �������  �����  ---
                del wbl
                k += 1
#---
#-------   2007 ���   --------
#        wbl = Workbook(encoding='cp1251')
#        wsl = wbl.add_sheet('����� �� �����')
#        self.vMikro = False
#        allTerrTotalTitle(self, wsl, \
#            """���� ������ ������� ����������� �� ������� �� ������ - ������� 2008 ����, ������������ ��������� ��������� ����� � ����������������� �������� ������ �� 2007 ��� (��� ����������������)"""
#            )
#        allTerrCoeffmal(self, wsl, self.malpok)
#
#        wbl.save(dirTotal + "����������������� �������� ������ �� 2007 ��� ��� ����� �����������" + '.xls')                             #---  ������ �������  �����  ---
#        del wbl
#        
#        wbi = Workbook(encoding='cp1251')
#        wsi = wbi.add_sheet('����� �� �����������������')
#        self.vMikro = True
#        allTerrTotalTitle(self, wsi, \
#            """���� ������ ������� ����������� �� ������� �� ������ - ������� 2008 ����, ������������ ��������� ��������� ����� � ����������������� �������� ������ �� 2007 ��� ��� ����������������"""
#            )
#        allTerrCoeffmal(self, wsi, self.micpok)
#
#        wbi.save(dirTotal + "����������������� �������� ������ �� 2007 ��� ��� ����������������" + '.xls')                             #---  ������ �������  �����  ---
#        del wbi
#-------   2008 ���   --------
#        wbl = Workbook(encoding='cp1251')
#        wsl = wbl.add_sheet('����� �� �����')
#        self.vMikro = False
#        allTerrTotalTitle(self, wsl, \
#            """���� ������ ������� ����������� �� ������� �� ������ - ������� 2008 ����, ������������ ��������� ��������� ����� � ����������������� �������� ������ �� 2008 ��� (��� ����������������)"""
#            )
#        allTerrCoeffmal(self, wsl, self.malpok)
#
#        wbl.save(dirTotal + "����������������� �������� ������ �� 2008 ��� ��� ����� �����������" + '.xls')                             #---  ������ �������  �����  ---
#        del wbl

        
        #---   ����������������� ����� � ������� �� �����   ---
        allOkvTerrCorr(self, False, '2007')                             # ����� 2007
        allOkvTerrCorr(self, True,  '2007')                             # ���������������� 2007
        allOkvTerrCorr(self, False, '2008')                             # ����� 2008
        allOkvTerrCorr(self, True,  '2008')                             # ���������������� 2008
    #---  ������ ������ Excel - ����� ������  ---

#--- Global functions  ---

def missVal(d):
    if type(d) is str:
        d = d.strip()
    elif type(d) is NoneType: 
        d = 0.0 
    return d

def missValAll(d):
    return map(missVal, d)

def missFloat(d):
    if type(d) is FloatType : 
        d = float(d) 
    else: 
        d = 0.0
    return d

def ToPrintLog (sMess):
    print str(datetime.now().strftime("%d.%m.%Y %H:%M:%S ")) + str(sMess)

def findFilesTerrAllXls():
    TT = TTotal()
    dataDir = dataDirXls
    os.chdir(dataDir)
    # �����, ����������, ���������, ����������������, �����
    #   0         1          2            3             4  
    j = 0
    k = 0
    if os.access(newDir +ListTOGS+ '.xls', os.F_OK):
        rb = xlrd.open_workbook(newDir +ListTOGS+ '.xls',formatting_info=True,encoding_override="cp1251")
        sh = rb.sheet_by_index(0)
        for i in range(1, sh.nrows):
#            print i, sh.cell_value(rowx=i, colx=1).encode('cp1251').strip(), sh.cell_value(rowx=i, colx=2), \
#                     sh.cell_value(rowx=i, colx=3).encode('cp1251').strip(), sh.cell_value(rowx=i, colx=4)
            nTerr  = sh.cell_value(rowx=i, colx=1).encode('cp1251').strip()
            kOkrug = sh.cell_value(rowx=i, colx=2)
            nOkrug = sh.cell_value(rowx=i, colx=3).encode('cp1251').strip()
            kOKATO = sh.cell_value(rowx=i, colx=4)
#            if kOKATO not in [14, 15, 17]:  continue                  # ����������� ��� �������
            if kOKATO <= 99:
                sokato = "%02i" % kOKATO
                TT.rTerr.append(cTerr(dataDir, mal, mic, kOKATO, nTerr, kOkrug, nOkrug))
                ToPrintLog(TT.rTerr[k].OKATO + ' ' + TT.rTerr[k].sTerr + ' ' + TT.rTerr[k].nOkrug + ' ' + TT.rTerr[k].sOkrug)
                nfoMal2008 = sokato+mal+'.xls'    # ��� ����� ������ ����� �����������.
                nfoMic2008 = sokato+mic+'.xls'    # ��� ����� ������ ����������������.
                if os.access(nfoMal2008, os.F_OK):
                    sMesFile = sMesFileYes
                    TT.aTerr.append(cTerr(dataDir,  mal, mic, kOKATO, nTerr, kOkrug, nOkrug))
                    ToPrintLog(sMesFile+' '+str(kOKATO)+' '+TT.aTerr[j].malFileXls)
                    j = j+1
                else:
                    sMesFile = sMesFileNo
                print sMesFile, i+1, kOKATO, nTerr,
                k = k + 1
                    
        ToPrintLog('���������� ������ �� '+str(j)+' ����������.')

    
    os.chdir(newDir)
    return TT
############################################################

if __name__ == '__main__':
    __name__ = 'RowsRatio'

os.chdir(newDir)
#print '������� ���������� (Python): ', os.curdir, os.getcwd()
dt = datetime.now().strftime("%d.%m.%Y %H:%M:%S  ")

ToPrintLog("---::  ����� ��������� ������.  ::---")

TT = findFilesTerrAllXls()

t0 = time()
ToPrintLog("����� (���������� ������): %s" % ctime(t0))

TT.allTerr()            # ���������� ������

t01 = time()
t1 = time() - t0
ToPrintLog("������ �� ������ (���������� ������): %.2f ���." % (t1))

TT.allTerrTotalToXLS()  # ������ ������ Excel

t2 = time() - t0
t21 = time() - t01
ToPrintLog("������ �� ������ (������ ������ Excel: %.2f ���.): %.2f ���." % (t21, t2))
