BEGIN PROGRAM PYTHON.
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
#         - �������� �� ������ 2008 ���� ������� 2007 ���� - ����������� ��
#           �����������, � ������ �����
#         + ��������� � ���������� ����� ���������������� � 2007 ���� �� ������� �����
#         
############################################
#       �������� �.�.   11.09.2010 20:12:20
############################################
from __future__ import with_statement
import os
from datetime import datetime, date, time
from time import *
from types import *
import SpssClient
import spss 
from xlwt import *
import xlrd

newDir         = 'e:\\Tmp\\Spss\\rowsratio'
dataDir        = newDir + '\ObData'
totalDir       = '\\Total'
totalDir2007   = newDir + '\\����� �� 2007\\'
ratioDir       = newDir + '\Ratio\\'
ListTOGS       = '\\ListTOGS'
nameListTOGS   = 'TOGS'
sav            = '.sav'
mal            = '_��'
mic            = '_��(�����)'
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
        self.sTerr = sT                 # �������� ����������
        self.nOkrug = "%02i" % kO       # ��� ������
        self.sOkrug = sO                # �������� ������
        self.malFile = dataDir + '\\' + self.OKATO + mal + '.sav'    # ��� ����� ������ ����� �����������.
        self.micFile = dataDir + '\\' + self.OKATO + mic + '.sav'    # ��� ����� ������ ����������������.
        self.ismalFile = False
        self.ismicFile = False
        self.nameDataSet = 'D' + self.OKATO                       # ��� DATASET

        self.malTot2008 = [0,0,0,0,0,0,0,0,0,0,0,0.0]                 # ����� �� 2008 ����
        self.malTot2009 = [0,0,0,0,0,0,0,0,0,0,0,0.0]                 # ����� �� 2009 ����
        self.malTot2007 = [0,0,0,0,0,0,0,0,0,0,0,0.0]                 # ����� �� 2007 ����
        self.malTos2007 = [0,0,0,0,0,0,0,0,0,0,0,0.0]                 # ����� �� 2007 ���� ����������������� 
        self.malcoeff   = [1,1,1,1,1,1,1,1,1,1,1,1.0]                 # ������������ ���������

        self.malTot2008Okv = []                 # ����� �� 2008 ���� ������ �� �����
        self.malTot2009Okv = []                 # ����� �� 2009 ���� ������ �� �����
        self.malTot2007Okv = []                 # ����� �� 2007 ���� ������ �� �����
        self.malTos2007Okv = []                 # ����� �� 2007 ���� ����������������� ������ �� ����� 
        self.malcoeffOkv   = []                 # ������������ ��������� ������ �� �����

        self.malnnab     = 0                                      # ���������� ����������
        self.malknab     = 0                                      # ���������� ���������� ����������
        self.malrnab     = 0                                      # ���������� ��������� ����������
        self.numvesmal   = 33           # ������ ���� � ����� = ����� ���������� - 1
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

        self.micTot2008 = [0,0,0,0,0,0,0,0,0,0,0.0]                 # ����� �� 2008 ����
        self.micTot2009 = [0,0,0,0,0,0,0,0,0,0,0.0]                 # ����� �� 2009 ����
        self.micTot2007 = [0,0,0,0,0,0,0,0,0,0,0.0]                 # ����� �� 2007 ����
        self.micTos2007 = [0,0,0,0,0,0,0,0,0,0,0.0]                 # ����� �� 2007 ���� �����������������
        self.miccoeff   = [1,1,1,1,1,1,1,1,1,1,1.0]                 # ������������ ���������

        self.micTot2008Okv = []                 # ����� �� 2008 ���� ������ �� �����
        self.micTot2009Okv = []                 # ����� �� 2009 ���� ������ �� �����
        self.micTot2007Okv = []                 # ����� �� 2007 ���� ������ �� �����
        self.micTos2007Okv = []                 # ����� �� 2007 ���� ����������������� ������ �� ����� 
        self.miccoeffOkv   = []                 # ������������ ��������� ������ �� �����

        self.micnnab     = 0                                      # ���������� ����������
        self.micknab     = 0                                      # ���������� ���������� ����������
        self.micrnab     = 0                                      # ���������� ��������� ����������
        self.numvesmic   = 30          # ������ ���� � ����� = ����� ���������� - 1
        self.micCHT      = 5            # ������ ������� ����������� � �������� ����� SPSS
        self.micVIR      = 8            # ������ ������� � �������� ����� SPSS
        self.micnCHT      = 2            # ������ ������� ����������� � �����. ������� ��������� ������
        self.micnVIR      = 3            # ������ ������� � �����. ������� ��������� ������
        self.micCHTres   = 15           # ����������� ������� ����������� ����������������
        self.micVIRres   = 60000000     # ����������� ������� ����������������

    def evalTotal(self, vM, mFile, listOkved, rulOk, \ 
                        Tot2008, Tot2009, coeff, Tot2008okv, Tot2009okv, coeffokv,\ 
                        Tot2007okv, Tos2007okv, Ratio2007okv, \ 
                        numves, kCHT, kVIR, CHTres, VIRres, npok):

        spss.Submit(r"""
        GET FILE = '%s'. 
        DATASET NAME %s.
        """ %(mFile, self.nameDataSet))

        stime = time()
#        ToPrintLog('������ Cursor ����������� ������ �� '+ mFile)

        nump=[num[0] for num in npok]
        nump.insert(1,numves)
        dataCursor=spss.Cursor(nump)
        data2008=map(missValAll, dataCursor.fetchall())
        dataCursor.close()

#        for k in range(0,len(data2008)):
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
        ToPrintLog('����� � ������� 2008 ����                     :'+str(len(data2008))+' �����������')
        ToPrintLog('�������� � ������� �� �������� 2007 ����      :'+str(len(data))+' �����������')
        ToPrintLog('����� ���������    �� �������� 2007 ����      :'+str(len(data2008)-len(data))+' �����������')
        ToPrintLog('    �� ��� � �������� ������                  :'+str(len([rw for rw in data2008 if rw[1] <= 0])))
#        for rw in data:
#            print rw
        data9 = [rw for rw in data2008 if (rw[1] > 0 \ 
                and rw[kCHT] <= CHTres and rw[kVIR] <= VIRres)]          # �� �������� 2009 ����

        ToPrintLog('�������� � ������� �� �������� 2009 ����      :'+str(len(data9))+' �����������')
        ToPrintLog('����� ���������    �� �������� 2009 ����      :'+str(len(data2008)-len(data9))+' �����������')
        ToPrintLog('    �� ��� � �������� ������                  :'+str(len([rw for rw in data2008 if rw[1] <= 0])))

        i = len([rv[0] for rv in data])                                  # ���������� �������� ����������
        i9 = len([rv[0] for rv in data9])                                # ���������� �������� ����������
        Tot2008[0] = '�����'                                             # ��� �����
        Tot2009[0] = '�����'                                             # ��� �����
        Tot2008[1] = sum(rv[1] for rv in data)                           # ������ ���������� ����������� �� ����
        Tot2009[1] = sum(rv[1] for rv in data9)                          # ������ ���������� ����������� �� ����
        
        ToPrintLog('������ �� �������� 2007 ����                  :'+str(Tot2008[1])+' �����������')
        ToPrintLog('������ �� �������� 2009 ����                  :'+str(Tot2009[1])+' �����������')
        
        for i in range(2, len(Tot2008)):                                 # �� �����������
            Tot2008[i]  = sum(rv[1] * rv[i] for rv in data)
            Tot2009[i]  = sum(rv[1] * rv[i] for rv in data9)

        k = 0
        for rokv in listOkved:                                           # ���������� ������ �� ����� (�� 2007 - ����� ������)
#            print rokv[0], rokv[1], rokv[2]
            Tot2008okv.append([0.0 for rw in Tot2008])
            Tot2009okv.append([0.0 for rw in Tot2009])
            coeffokv.append([0.0 for rw in coeff])
            
            Tot2007okv.append([0.0 for rw in Tot2008])
            Tos2007okv.append([0.0 for rw in Tot2008])
            Ratio2007okv.append([0.0 for rw in coeff])
            
            Tot2008okv[k][0] = rokv[0]
            Tot2009okv[k][0] = rokv[0]
            coeffokv[k][0]   = rokv[0]

            Tot2007okv[k][0] = rokv[0]
            Tos2007okv[k][0] = rokv[0]
            Ratio2007okv[k][0] = rokv[0]

#            print [rv[0] for rv in data if rv[0] in rokv[2]]
            Tot2008okv[k][1] = sum(rv[1] for rv in data if rv[0] in rokv[2])
            Tot2009okv[k][1] = sum(rv[1] for rv in data9 if rv[0] in rokv[2])
            for i in range(2, len(Tot2008)):                                      # �� �����������
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

        
        for itot in range(1, len(Tot2009)):
            if Tot2008[itot] > 0:
                coeff[itot] = Tot2009[itot] / Tot2008[itot]
            else:
                coeff[itot] = 1
            #print itot, self.malcoeff[itot], self.malTot2009[itot], self.malTot2008[itot]

        for ir in range(0, len(Tot2008okv)):
            for itot in range(1, len(Tot2009)):
                if Tot2008okv[ir][itot] > 0:
                    coeffokv[ir][itot] = Tot2009okv[ir][itot] / Tot2008okv[ir][itot]
                else:
                    coeffokv[ir][itot] = 1

#        for i in range(0, len(coeffokv)):
#            print coeffokv[i]

        etime = time() - stime
        ToPrintLog('����� ������� ������ �� '+ mFile.rsplit('\\',1)[1] + ' : %.2f ���.' % (etime))
        
        spss.Submit(r"""
        DATASET CLOSE *. 
        """ )

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
    [1  , '�����', 'tab_33pred-2007.xls', 'tab_33pred-2007-corr'],
    [5  , 'C������ ����������� ���������� - �����,���.', 'tab_33(01-09)_list1.xls', 'tab_33(01-09)_list1-corr'],
    [8  , '������� (�����) �� ���������� �������, ���������, �����, ����� (��� ���, ������� � ����������� ������������ ��������),���.���.', 'tab_33(12)_list1.xls', 'tab_33(12)_list1-corr'],
    [11 , '���������� � �������� ������� (� ����� ����� � ������������� �� ������� �������� �������),���.���.', 'tab_33(13)_list1.xls', 'tab_33(13)_list1-corr'],
    [14 , '������ ����������� �� ����� ������������,���.���.', 'tab_33(o)_list1.xls', 'tab_33(o)_list1-corr'],
    [17 , '��������� ������� ������������ ������������, ��������� ����� � ����� ������������ ������ (��� ��� � ������),���.���.', 'tab_33(06)_list1.xls', 'tab_33(06)_list1-corr'],
    [20 , '������� ������� �������������� ������������ (��� ��� � �������),���.���.', 'tab_33(07)_list1.xls', 'tab_33(07)_list1-corr'],
    [23 , '������� ����������� ���������� ���������� ������� (��� ������� �������������),���.', 'tab_33(02-09)_list1.xls', 'tab_33(02-09)_list1-corr'],
    [26 , '���� ����������� ���������� ����� ������� �������������,���.���.', 'tab_33(03-11)_list1.xls', 'tab_33(03-11)_list1-corr'],
    [29 , '���� ����������� ���������� ����� ���� ����������,���.���.', 'tab_33(01-11)_list1.xls', 'tab_33(01-11)_list1-corr'],
    [32 , '���� ����������� ���������� ����� ���������� ���������� ������� (��� ������� �������������),���.���.', 'tab_33(02-11)_list1.xls', 'tab_33(02-11)_list1-corr']
    ]
    micpok = [
    [1  , '�����'],
    [5  , '������� ����������� ���������� (������� ����������� ������ �� ��������� ����������-��������� ���������),���.'],
    [8  , '������� (�����) �� ������� �������, ���������, �����, ����� (��� ���, ������� � ����������� ������������ ��������),���.���.'],
    [11 , '���������� � �������� ������� (� ����� ����� � ������������ �� ������� �������� �������) - �����,���.���.'],
    [14 , '������ �����������,���.���.'],
    [17 , '��������� ������� ������������ ������������, ��������� ����� � ����� ������������ ������ (��� ���, ������� � ����������� ������������ ��������),���.���.'],
    [20 , '������� ������� �������������� ������������ (��� ���, ������� � ����������� ������������ ��������),���.���.'],
    [23 , '������� ����������� ���������� ���������� ������� (��� ������� �������������),���.'],
    [26 , '���� ����������� ���������� ����� ����������,���.���.'],
    [29 , '���� ����������� ���������� ����� ���������� ���������� ������� � ������� �������������,���.���.']
    ]

    def __init__ (self):
        self.OkRul = OkvedAss(ToPrintLog, newDir)           # OkRul.rulesAssPlus
        self.ListOkv = ListOkved(ToPrintLog, newDir)        # self.ListOkv.listOkvedmal[][]
    def getTotal2007(self):
        np = 0
        os.chdir(totalDir2007)
        for pok in self.malpok:                 # �� �������� �����������
            np = np + 1                         # ����� ���������� � ������ ������ malTot2007
            if os.access(pok[2], os.F_OK):      # ���� �� ���� ����������
                ToPrintLog('������ ���� ������ ' + pok[2])
                rb = xlrd.open_workbook(totalDir2007 + pok[2],formatting_info=True,encoding_override="cp1251")
                sh = rb.sheet_by_index(0)
                for rt in self.aTerr:                       # ��������� ����������
                    for icol in range(3, sh.ncols-1):       # �������� ���������� � ��������!
                        sName = sh.cell_value(rowx=4, colx=icol).encode('cp1251').strip()
                        #print icol, type(sName), type(rt.sTerr), sName, '-', rt.sTerr, sName.strip() == rt.sTerr.strip()
                        if sName == rt.sTerr.strip():       # ���� ����������!
                            rt.malTot2007[np] = sh.cell_value(rowx=6, colx=icol)
                            ToPrintLog('�������: ' + sName + ' ' + pok[1][0:17] + ' �����: ' + str(rt.malTot2007[np]))
                            k = 0
                            for rokv in rt.malTot2007Okv:   # ����� 2007 �� �����-2007 - ����� ������
                                compp = False
                                for i in range(6, sh.nrows):    # ���� ��� ����� � �����
                                    okvFile = sh.cell_value(rowx=i, colx=2).encode('cp1251').strip()
                                    if rt.malTot2007Okv[k][0] == okvFile:   # ���� ������ ����!
                                        compp = True
                                        okvTot = sh.cell_value(rowx=i, colx=icol)
                                        if okvTot == u'-':
                                            rt.malTot2007Okv[k][np] = 0.0
                                        else:
                                            rt.malTot2007Okv[k][np] = float(okvTot)
                                        break
                                else:
                                    if not compp:
                                        ToPrintLog('    ����������� ����� '+rt.malTot2007Okv[k][0]+' � ����� ������.')

                                k += 1
                                
#                            for i in range(0, len(rt.malTot2007Okv)):
#                                print rt.malTot2007Okv[i]
 

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

                for np in range(0, len(rt.micRatio2007)):
                    rt.malRatio2007[np]  = 1 - rt.micRatio2007[np]          # ���� ����� �� �����
                    rt.malTos2007[np] = rt.malTot2007[np] * rt.malRatio2007[np] * rt.malcoeff[np]
                for np in range(0, len(rt.micTot2007)):
                    rt.micTot2007[np] = rt.malTot2007[np]
                    rt.micTos2007[np] = rt.malTot2007[np] * rt.micRatio2007[np] * rt.miccoeff[np]
                rt.micTot2007[9] = rt.malTot2007[11]
                rt.micTos2007[9] = rt.malTot2007[11] * rt.micRatio2007[11] * rt.miccoeff[9]
                
                # ������ �� �����
                k = 0
                for rokv in rt.micRatio2007Okv:   # ���� ���������������� � 2007 �� �����-2007 - ����� ������
                    rProk = getProd(self, rPred, rokv[0], 0, 3)
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
#                for i, r in enumerate(rt.micRatio2007Okv):
#                    print r
#                rt.malRatio2007Okv[0][1 ] = rPred
#                rt.malRatio2007Okv[0][2 ] = rPred
#                rt.malRatio2007Okv[0][3 ] = rProd
#                rt.malRatio2007Okv[0][4 ] = rProd
#                rt.malRatio2007Okv[0][5 ] = rProd
#                rt.malRatio2007Okv[0][6 ] = rOtgr
#                rt.malRatio2007Okv[0][7 ] = rProd
#                rt.malRatio2007Okv[0][8 ] = rPred
#                rt.malRatio2007Okv[0][9 ] = rPred
#                rt.malRatio2007Okv[0][10] = rPred
#                rt.malRatio2007Okv[0][11] = rPred
            else:
                ToPrintLog('�� ������!! ���� ����� ��� ' + rt.OKATO + ' ' + rt.sTerr.strip())
            for i, rat in enumerate(rt.micRatio2007Okv):
                for j in range(1, len(rat)) :
                    rt.malRatio2007Okv[i][j] = 1 - rt.micRatio2007Okv[i][j]
                    rt.malTos2007Okv[i][j] = rt.malTot2007Okv[i][j] * rt.malRatio2007Okv[i][j] * rt.malcoeffOkv[i][j]

    def allTerr(self):
        for rt in self.aTerr:
            ToPrintLog('....'+rt.nOkrug+' '+rt.sOkrug+' '+rt.OKATO+' '+rt.sTerr.strip()+' ����� �����������')
            self.vMikro = False
            rt.evalTotal(self.vMikro, rt.malFile, self.ListOkv.listOkvedmal, self.OkRul.rulesAssPlus,
            rt.malTot2008, rt.malTot2009, rt.malcoeff, rt.malTot2008Okv, rt.malTot2009Okv, rt.malcoeffOkv, 
            rt.malTot2007Okv, rt.malTos2007Okv, rt.malRatio2007Okv, rt.numvesmal,
            rt.malnCHT, rt.malnVIR, rt.malCHTres, rt.malVIRres, self.malpok)

            ToPrintLog('....'+rt.nOkrug+' '+rt.sOkrug+' '+rt.OKATO+' '+rt.sTerr.strip()+' ����� �����������')
            self.vMikro = True
            rt.evalTotal(self.vMikro, rt.micFile, self.ListOkv.listOkvedmic, self.OkRul.rulesAssPlus,
            rt.micTot2008, rt.micTot2009, rt.miccoeff, rt.micTot2008Okv, rt.micTot2009Okv, rt.miccoeffOkv, 
            rt.micTot2007Okv, rt.micTos2007Okv, rt.micRatio2007Okv, rt.numvesmic,
            rt.micnCHT, rt.micnVIR, rt.micCHTres, rt.micVIRres, self.micpok)

            ToPrintLog('        |����|       |            | �����������  | ����������� ')
            ToPrintLog('2008 ���' + '|{0:4}|{1:7d}|{2:12}|{3:13.0f} |{4:13.0f}'.format(rt.OKATO,rt.malnnab,rt.malTot2008[0],rt.malTot2008[1], rt.malTot2008[2]))
            ToPrintLog('2009 �����           |{0:12}|{1:13.0f} |{2:13.0f}'.format(rt.malTot2009[0],rt.malTot2009[1],rt.malTot2009[2]))
            ToPrintLog('   ����������� ����� |{0:12}|{1:13.4f} |{2:13.4f}'.format(rt.malcoeff[0],rt.malcoeff[1],rt.malcoeff[2]))

            ToPrintLog('2008 ���' + '|{0:4}|{1:7d}|{2:12}|{3:13.0f} |{4:13.0f}'.format(rt.OKATO,rt.malnnab,rt.micTot2008[0],rt.micTot2008[1], rt.micTot2008[2]))
            ToPrintLog('2009 �����           |{0:12}|{1:13.0f} |{2:13.0f}'.format(rt.micTot2009[0],rt.micTot2009[1],rt.micTot2009[2]))
            ToPrintLog('   ����������� ����� |{0:12}|{1:13.4f} |{2:13.4f}'.format(rt.miccoeff[0],rt.miccoeff[1],rt.miccoeff[2]))
        
        closeSPSSfiles()
        
        self.getTotal2007()             # ����� ����� 2007 ����
        self.getRatioMic()              # ����� ���� ����������������
        
    def allTerrTotalToXLS(self):
        ToPrintLog('������� ������ �� ����������� � Excel')
        dirTotal = newDir + totalDir + '\\'
        ezxf = easyxf
        heading_xf = ezxf('font: name Calibri, bold on, height 230; \
        align: wrap yes, vert centre, horiz center') 
        headcol_xf = ezxf('font: name Calibri, bold on; \
        align: wrap yes, vert centre, horiz center; \
        borders: left 1, right 1, top 1, bottom 1') 
        headrow_xf = ezxf('font: name Calibri; \
        align: wrap yes, vert centre, horiz left; \
        borders:left 1, right 1, top 1, bottom 1') 
        valnrow_xf = ezxf('font: name Calibri, bold off; \
        align: wrap no, vert centre, horiz right; \
        borders:left 1, right 1, top 1, bottom 1', \
        num_format_str="### ### ### ###") 
        valdrow_xf = ezxf('font: name Calibri, bold off; \
        align: wrap no, vert centre, horiz right; \
        borders:left 1, right 1, top 1, bottom 1', \
        num_format_str="### ### ### ###.0") 
        valfrow_xf = ezxf('font: name Calibri, bold off; \
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
                ToPrintLog(rt.nOkrug+ '|{0:4}|'.format(rt.OKATO))
                ws.write(2,it,rt.sTerr,headcol_xf)                                # �������� ����������
                ws.write(3,it,it-1,headcol_xf)
                ws.col(it).width = 0x0d00 + 50                                    # ������ ������� � �������
                it = it + 1
            
        def allTerrCoeffmal (self, ws, npok):
            nextRow = allTerrTotalCoeffmal(self, self.vMikro, ws, "���������� �����������", 1, 4)
            for i in range(2, len(npok)):
                nextRow  = allTerrTotalCoeffmal(self, self.vMikro, ws, npok[i-1][1], i, nextRow)
    
            ws.col(1).width = 0x0d00 * 3                            # ������ ������� � ��������� ����������
#---

        wbl = Workbook(encoding='cp1251')
        wsl = wbl.add_sheet('����� �� �����')
        self.vMikro = False
        allTerrTotalTitle(self, wsl, \
            """���� ������ ������� ����������� �� ������� �� ������ - ������� 2008 ����, ������������ ��������� ��������� ����� � ����������������� �������� ������ �� 2007 ��� (��� ����������������)"""
            )
        allTerrCoeffmal(self, wsl, self.malpok)

        wbl.save(dirTotal + "����������������� �������� ������ �� 2007 ��� ��� ����� �����������" + '.xls')                             #---  ������ �������  �����  ---
        del wbl
        
        wbi = Workbook(encoding='cp1251')
        wsi = wbi.add_sheet('����� �� �����������������')
        self.vMikro = True
        allTerrTotalTitle(self, wsi, \
            """���� ������ ������� ����������� �� ������� �� ������ - ������� 2008 ����, ������������ ��������� ��������� ����� � ����������������� �������� ������ �� 2007 ��� ��� ����������������"""
            )
        allTerrCoeffmal(self, wsi, self.micpok)

        wbi.save(dirTotal + "����������������� �������� ������ �� 2007 ��� ��� ����������������" + '.xls')                             #---  ������ �������  �����  ---
        del wbi
        
        #---   ����������������� ����� � ������� �� �����   ---
        for np in malpok:
            xlsfile = np[3]+'.xls'
            
        
        
        
    #---  ������ ������ Excel - ����� ������  ---

#--- Global functions  ---

def closeSPSSfiles():
    spss.Submit(r"""
    DATASET CLOSE ALL.
    NEW FILE.
    """)

def missVal(d):
    if type(d) is str:
        d = d.strip()
    elif type(d) is NoneType: 
        d = 0.0 
    return d

def missValAll(d):
    return map(missVal, d)

def ToPrintLog (sMess):
    print str(datetime.now().strftime("%d.%m.%Y %H:%M:%S ")) + str(sMess)

def ModifyTypeVesAll(TT):
    """ �������������� ���� ���������� ��� �� ���������� � ��������   """
    ToPrintLog('�������������� ���� ���������� ��� �� ���������� � ��������')
    for rTerr in TT.aTerr:
        #ToPrintLog(str(rTerr.OKATO)+str(rTerr.sTerr)+str(rTerr.nOkrug)+str(rTerr.sOkrug)+str(rTerr.malFile)+str(rTerr.micFile))
        #ToPrintLog(str(rTerr.OKATO)+str(rTerr.sTerr))
        #ToPrintLog("---  :: ����� � ����� ����������� ::  ---")
        spss.Submit(r"""
        GET FILE = '%s'. 
        ALTER TYPE      ���(F5.2).
        VARIABLE LEVEL  ���(SCALE).
        FORMATS         ���(F5.2). 
        EXECUTE.
        SAVE /OUTFILE = '%s'. 
        DATASET CLOSE *.
        GET FILE = '%s'. 
        ALTER TYPE      ���(F5.2).
        VARIABLE LEVEL  ���(SCALE).
        FORMATS         ���(F5.2). 
        EXECUTE.
        SAVE /OUTFILE = '%s'. 
        DATASET CLOSE *.
        """ %(rTerr.malFile, rTerr.malFile, rTerr.micFile, rTerr.micFile))

def FindFilesTerrAll():
    """ �������� �� ������� ������ ������ ������� �� ����� � ����� � ������������ ������� �������� cTerr """
    spss.Submit(r"""
    DATASET CLOSE ALL.
    GET FILE = '%s%s%s'. 
    DATASET NAME %s.
    """ %(newDir, ListTOGS, sav, nameListTOGS))
    TT = TTotal()
    # �����, ����������, ���������, ����������������, �����, �������������, ��������, 
    #   0         1          2            3             4           5          6

    with spss.DataStep():
        # ������� �������� ����
        # SpssClient.LogToViewer('������ ������� �������� ����.')
        datasetListTOGS = spss.Dataset(name=nameListTOGS)
        dt = datetime.now().strftime("%d.%m.%Y %H:%M:%S ")
        #print dt, '���������� ��� ������.'
        sokato = '00'
        os.chdir(dataDir)
        j = 0
        k = 0
        for i in range(len(datasetListTOGS.cases)):
           if datasetListTOGS.cases[i,4][0] <= 99: 
              sokato = "%02i" % datasetListTOGS.cases[i,4][0]
              TT.rTerr.append(cTerr(dataDir, mal, mic, datasetListTOGS.cases[i,4][0]
                                     ,datasetListTOGS.cases[i,1][0]
                                     ,datasetListTOGS.cases[i,2][0]
                                     ,datasetListTOGS.cases[i,3][0]
                                      )
                               )
              ToPrintLog(TT.rTerr[k].OKATO + ' ' + TT.rTerr[k].sTerr + ' ' + TT.rTerr[k].nOkrug + ' ' + TT.rTerr[k].sOkrug)
              k = k + 1
              nfoMal2008 = sokato + mal + sav
              nfoMic2008 = sokato + mic + sav
              if os.access(nfoMal2008, os.F_OK):
                  sMesFile = sMesFileYes
                  TT.aTerr.append(cTerr(dataDir,  mal, mic, datasetListTOGS.cases[i,4][0]
                                     ,datasetListTOGS.cases[i,1][0]
                                     ,datasetListTOGS.cases[i,2][0]
                                     ,datasetListTOGS.cases[i,3][0]
                                      )
                               )
                  ToPrintLog(sMesFile+' '+str(datasetListTOGS.cases[i,4][0])+' '+TT.aTerr[j].malFile)
                  j = j+1

              else:
                  sMesFile = sMesFileNo
                  #print sMesFile, i+1, datasetListTOGS.cases[i,4][0], nfoMal2008
              datasetListTOGS.cases[i,6] = sMesFile +  nfoMal2008
              datasetListTOGS.cases[i,5] = dt
              if j >= 2:                                          # ����������� ��� �������
                  break 
        ToPrintLog('���������� ������ �� '+str(j)+' ����������.')
    spss.Submit(r"""
    DATASET ACTIVATE %s.
    SAVE OUTFILE = '%s%s%s'.
    DATASET CLOSE *. 
    """ %(nameListTOGS, newDir, ListTOGS, sav))
    os.chdir(newDir)
    return TT
        
############################################################

if __name__ == '__main__':
    __name__ = 'RatioRows'

SpssClient.StartClient()
SpssClient.SetCurrentDirectory(newDir)
os.chdir(newDir)
print '������� ���������� (Python): ', os.curdir, os.getcwd()
print '������� ���������� (Spss  ): ', SpssClient.GetCurrentDirectory() 
dt = datetime.now().strftime("%d.%m.%Y %H:%M:%S  ")
print (str(dt) + '����� ��������� ������ TestStr.')

#NewOutputDoc = SpssClient.NewOutputDoc()
#NewOutputDoc.SetAsDesignatedOutputDoc()
NewOutputDoc = SpssClient.GetDesignatedOutputDoc()
SpssOutputUI = NewOutputDoc.GetOutputUI() 
SpssOutputUI.SetVisible(True) 

spss.Submit(r"""
DATASET CLOSE ALL.
GET FILE = '%s%s%s'. 
DATASET NAME %s.
""" %(newDir, ListTOGS, sav, nameListTOGS))

TT = FindFilesTerrAll()

#ModifyTypeVesAll(TT)

t0 = time()
ToPrintLog("Start (���������� ������): %s" % ctime(t0))

TT.allTerr()            # ���������� ������

t01 = time()
t1 = time() - t0
ToPrintLog("������ �� ������ (���������� ������): %.2f ���." % (t1))

TT.allTerrTotalToXLS()  # ������ ������ Excel

t2 = time() - t0
t21 = time() - t01
ToPrintLog("������ �� ������ (������ ������ Excel: %.2f ���.): %.2f ���." % (t21, t2))


# ---
closeSPSSfiles()

SpssClient.StopClient()

END PROGRAM.