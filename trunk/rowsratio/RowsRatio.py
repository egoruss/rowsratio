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
#         
########################################
#       11.08.2010 23:17:47
########################################
from __future__ import with_statement
import os
from datetime import datetime, date, time
from time import *
from types import *
import SpssClient
import spss 
from xlwt import *
import xlrd

newDir         = 'e:\Tmp\Spss\Rows'
dataDir        = newDir + '\ObData'
totalDir       = '\Total'
totalDir2007   = newDir + '\����� �� 2007\\'
ratioDir       = newDir + '\Ratio\\'
ListTOGS       = '\ListTOGS'
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
        self.malnnab     = 0                                      # ���������� ����������
        self.malknab     = 0                                      # ���������� ���������� ����������
        self.malrnab     = 0                                      # ���������� ��������� ����������
        self.numvesmal   = 33           # ������ ���� � ����� = ����� ���������� - 1
        self.malCHT      = 5            # ������ ������� �����������
        self.malVIR      = 8            # ������ �������
        self.malCHTres   = 100          # ����������� ������� ����������� �����
        self.malVIRres   = 400000000    # ����������� ������� �����

        self.micRatio2007 = [0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0]             # ���� ���������������� �� 2007 ����
        self.malRatio2007 = [0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0]             # ���� ����� ����������� �� 2007 ����

        self.micTot2008 = [0,0,0,0,0,0,0,0,0,0,0.0]                 # ����� �� 2008 ����
        self.micTot2009 = [0,0,0,0,0,0,0,0,0,0,0.0]                 # ����� �� 2009 ����
        self.micTot2007 = [0,0,0,0,0,0,0,0,0,0,0.0]                 # ����� �� 2007 ����
        self.micTos2007 = [0,0,0,0,0,0,0,0,0,0,0.0]                 # ����� �� 2007 ���� �����������������
        self.miccoeff   = [1,1,1,1,1,1,1,1,1,1,1.0]                 # ������������ ���������
        self.micnnab     = 0                                      # ���������� ����������
        self.micknab     = 0                                      # ���������� ���������� ����������
        self.micrnab     = 0                                      # ���������� ��������� ����������
        self.numvesmic   = 30          # ������ ���� � ����� = ����� ���������� - 1
        self.micCHT      = 5            # ������ ������� �����������
        self.micVIR      = 8            # ������ �������
        self.micCHTres   = 15           # ����������� ������� ����������� ����������������
        self.micVIRres   = 60000000     # ����������� ������� ����������������

    def evalTotal(self, vM, mFile, numves, nnab, Tot2008, Tot2009, coeff, nCHT, nVIR, CHTres, VIRres, npok):

        def totalmal(self, rules, row, ves, total):
            if rules:
                total[0]  = total[0]  + ves                           # ���������� �����������
                for i in range(1, len(total)-1):
                    total[i]  = total[i]  + missVal(dmal.cases[row, npok[i-1][0]][0]) * ves
                
                total[len(total)-1] = total[len(total)-1] + 1                        # ���������� ���������� ����������

#        def totalmal(self, rules, row, ves, total):
#            if rules:
#                total[0]  = total[0]  + ves                           # ���������� �����������
#                for i in range(1, len(total)-1):
#                    total[i]  = total[i]  + missVal(dmal.cases[row, npok[i-1][0]][0]) * ves
#                
#                total[len(total)-1] = total[len(total)-1] + 1                        # ���������� ���������� ����������

        def rules2008(self, row):
            """ ������� 2008 ����  """
            return True
        
        def rules2009(self, row, nCHT, nVIR, CHTres, VIRres):
            """ ������� 2009 ����  """
            if  missVal(dmal.cases[row,nCHT] [0]) <= CHTres and \
                missVal(dmal.cases[row,nVIR] [0]) <= VIRres :
                return True
            return False

#        def rules2009(self, row, nCHT, nVIR, CHTres, VIRres):
#            """ ������� 2009 ����  """
#            if  missVal(dmal.cases[row,nCHT] [0]) <= CHTres and \
#                missVal(dmal.cases[row,nVIR] [0]) <= VIRres :
#                return True
#            return False

        def rules2007(self, row):
            """ ������� 2007 ����  """
            return True
            
        spss.Submit(r"""
        GET FILE = '%s'. 
        DATASET NAME %s.
        """ %(mFile, self.nameDataSet))
        with spss.DataStep():
            # �� ����������� �������
            dmal = spss.Dataset(name=self.nameDataSet)
            self.malknab = 0
            rowval = []
            stime = time()
            ToPrintLog('������ ����������� ������ �� '+ mFile)
            for i in range(len(dmal.cases)):
                if missVal(dmal.cases[i,0][0]) <> 0 and missVal(dmal.cases[i,numves][0]) <> 0:         # �� ��������� ���� � ��������� �����
                    print '�����,����, ��� ', i, missVal(dmal.cases[i,0][0]), missVal(dmal.cases[i,numves][0])
                    rowval[i][0].append(missVal(dmal.cases[i,numves][0]))                # ��� �����������
                    for ip in range(1, len(Tot2008)-1):
                        rowval[i][ip].append(missVal(dmal.cases[i, npok[ip-1][0]][0]))
            etime = time() - stime
            ToPrintLog('����� ������ ����������� ������ �� '+ mFile + ' : %.2f ���.' % (etime))

            stime = time()
            ToPrintLog('���������� ������ ����� �� SPSS-����� '+ mFile)
            for i in range(len(dmal.cases)):
               ves = missVal(dmal.cases[i,numves][0])                # ��� �����������
               if missVal(dmal.cases[i,0][0]) <> 0 and ves <> 0:         # �� ��������� ���� � ��������� �����
                   totalmal(self, True,    i, ves, Tot2008)              # �� �������� 2008 ����
                   totalmal(self, rules2009(self, i, nCHT, nVIR, CHTres, VIRres), i, ves, Tot2009)              # �� �������� 2009 ����
                
            etime = time() - stime
            ToPrintLog('����� ���������� ������ ����� �� SPSS-����� '+ mFile + ' : %.2f ���.' % (etime))

        nnab = i                # ���������� ����������
        
        for itot in range(len(Tot2009)):
            if Tot2008[itot] > 0:
                coeff[itot] = Tot2009[itot] / Tot2008[itot]
            else:
                coeff[itot] = 1
            #print itot, self.malcoeff[itot], self.malTot2009[itot], self.malTot2008[itot]
            
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
    [5  , 'C������ ����������� ���������� - �����,���.', 'tab_33(01-09)_list1.xls'],
    [8  , '������� (�����) �� ���������� �������, ���������, �����, ����� (��� ���, ������� � ����������� ������������ ��������),���.���.', 'tab_33(12)_list1.xls'],
    [11 , '���������� � �������� ������� (� ����� ����� � ������������� �� ������� �������� �������),���.���.', 'tab_33(13)_list1.xls'],
    [14 , '������ ����������� �� ����� ������������,���.���.', 'tab_33(o)_list1.xls'],
    [17 , '��������� ������� ������������ ������������, ��������� ����� � ����� ������������ ������ (��� ��� � ������),���.���.', 'tab_33(06)_list1.xls'],
    [20 , '������� ������� �������������� ������������ (��� ��� � �������),���.���.', 'tab_33(07)_list1.xls'],
    [23 , '������� ����������� ���������� ���������� ������� (��� ������� �������������),���.', 'tab_33(02-09)_list1.xls'],
    [26 , '���� ����������� ���������� ����� ������� �������������,���.���.', 'tab_33(03-11)_list1.xls'],
    [29 , '���� ����������� ���������� ����� ���� ����������,���.���.', 'tab_33(01-11)_list1.xls'],
    [32 , '���� ����������� ���������� ����� ���������� ���������� ������� (��� ������� �������������),���.���.', 'tab_33(02-11)_list1.xls']
    ]
    micpok = [
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
        self.OkRul = OkvedAss(ToPrintLog, newDir)
        self.ListOkv = ListOkved(ToPrintLog, newDir)
    def getTotal2007(self):
        np = 0
        for pok in self.malpok:
            os.chdir(totalDir2007)
            np = np + 1                 # ����� ���������� � ������ ������ malTot2007
            if os.access(pok[2], os.F_OK):
                ToPrintLog('������ ���� ������ ' + pok[2])
                rb = xlrd.open_workbook(totalDir2007 + pok[2],formatting_info=True,encoding_override="cp1251")
                sh = rb.sheet_by_index(0)
                for rt in self.aTerr:
                    for icol in range(3, sh.ncols-1):
                        sName = sh.cell_value(rowx=4, colx=icol).encode('cp1251').strip()
                        #print icol, type(sName), type(rt.sTerr), sName, '-', rt.sTerr, sName.strip() == rt.sTerr.strip()
                        if sName == rt.sTerr.strip():
                            rt.malTot2007[np] = sh.cell_value(rowx=6, colx=icol)
                            ToPrintLog('�������: ' + sName + ' ' + str(rt.malTot2007[np]))
            else: 
                ToPrintLog('�� ������!! ���� ������ ' + pok[2])

    def getRatioMic(self):
        def getProd(self, rP, stro, scol, rcol):
            for rownum in range(sh.nrows):
                parsName = sh.cell_value(rowx=rownum, colx=scol).encode('cp1251')
                #print type(parsName), parsName.strip(), parsName.find(stro)
                if parsName.find(stro) >= 0:
                    ToPrintLog('������� ���� -' + parsName + '- ' + str(sh.cell_value(rowx=rownum, colx=rcol)))
                    return sh.cell_value(rowx=rownum, colx=rcol)
                    
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
                
                rt.micRatio2007[0]  =  rPred
                rt.micRatio2007[1]  =  rPred
                rt.micRatio2007[2]  =  rProd
                rt.micRatio2007[3]  =  rProd
                rt.micRatio2007[4]  =  rProd
                rt.micRatio2007[5]  =  rOtgr
                rt.micRatio2007[6]  =  rProd
                rt.micRatio2007[7]  =  rPred
                rt.micRatio2007[8]  =  rPred
                rt.micRatio2007[9]  =  rPred
                rt.micRatio2007[10] =  rPred
                rt.micRatio2007[11] =  rPred

                for np in range(0, len(rt.micRatio2007)):
                    rt.malRatio2007[np]  = 1 - rt.micRatio2007[np]          # ���� ����� �� �����
                    rt.malTos2007[np] = rt.malTot2007[np] * rt.malRatio2007[np] * rt.malcoeff[np]
                for np in range(0, len(rt.micTot2007)):
                    rt.micTot2007[np] = rt.malTot2007[np]
                    rt.micTos2007[np] = rt.malTot2007[np] * rt.micRatio2007[np] * rt.miccoeff[np]
                rt.micTot2007[8] = rt.malTot2007[10]
                rt.micTos2007[8] = rt.malTot2007[10] * rt.micRatio2007[10] * rt.miccoeff[8]
            else:
                ToPrintLog('�� ������!! ���� ����� ��� ' + rt.OKATO + ' ' + rt.sTerr.strip())
                        

    def allTerr(self):
        ToPrintLog('    ����� : ���������� ������������� �� �����������')
        ToPrintLog('���.    |����|�������|���� �������| �����������  | ����������� ')
        for rt in self.aTerr:
            self.vMikro = False
            rt.evalTotal(self.vMikro, rt.malFile, rt.numvesmal, rt.malnnab, 
            rt.malTot2008, rt.malTot2009, rt.malcoeff, 
            rt.malCHT, rt.malVIR, rt.malCHTres, rt.malVIRres, self.malpok)

            self.vMikro = True
            rt.evalTotal(self.vMikro, rt.micFile, rt.numvesmic, rt.micnnab, 
            rt.micTot2008, rt.micTot2009, rt.miccoeff, 
            rt.micCHT, rt.micVIR, rt.micCHTres, rt.micVIRres, self.micpok)

            ToPrintLog(rt.nOkrug+'2008��' + '|{0:4}|{1:7d}|{2:12.0f}|{3:13.0f} |{4:13.0f}'.format(rt.OKATO,rt.malnnab,rt.malTot2008[11],rt.malTot2008[0], rt.malTot2008[1]))
            ToPrintLog('   2009 �����        |{0:12.0f}|{1:13.0f} |{2:13.0f}'.format(rt.malTot2009[11],rt.malTot2009[0],rt.malTot2009[1]))
            ToPrintLog('   ����������� ����� |{0:12.4f}|{1:13.4f} |{2:13.4f}'.format(rt.malcoeff[11],rt.malcoeff[0],rt.malcoeff[1]))

            ToPrintLog(rt.nOkrug+'2008��' + '|{0:4}|{1:7d}|{2:12.0f}|{3:13.0f} |{4:13.0f}'.format(rt.OKATO,rt.malnnab,rt.micTot2008[10],rt.micTot2008[0], rt.micTot2008[1]))
            ToPrintLog('   2009 �����        |{0:12.0f}|{1:13.0f} |{2:13.0f}'.format(rt.micTot2009[10],rt.micTot2009[0],rt.micTot2009[1]))
            ToPrintLog('   ����������� ����� |{0:12.4f}|{1:13.4f} |{2:13.4f}'.format(rt.miccoeff[10],rt.miccoeff[0],rt.miccoeff[1]))
        
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
            ws.write(startRow + 1,1,"    ���� 2008 ���� -�������",headrow_xf)
            ws.write(startRow + 2,1,"    ���� 2008 ���� - ��� ��������",headrow_xf)
            ws.write(startRow + 3,1,"    ����������� ��������� ����",headrow_xf)
            ws.write(startRow + 4,1,"    ����� 2007 ���� ��������",headrow_xf)
            ws.write(startRow + 5,1,nameRatio, headrow_xf)
            ws.write(startRow + 6,1,"    ����� 2007 ���� �����������������",headrow_xf)
            for rt in self.aTerr:
                if mic:
                    if nTot in (0, 1, 7):
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
                    if nTot in (0, 1, 7):
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
            nextRow = allTerrTotalCoeffmal(self, self.vMikro, ws, "���������� �����������", 0, 4)
            for i in range(1, len(npok)+1):
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
    #---  ������ ������ Excel - ����� ������  ---

#--- Global functions  ---

def closeSPSSfiles():
    spss.Submit(r"""
    DATASET CLOSE ALL.
    NEW FILE.
    """)

def missVal(d):
    if type(d) is NoneType: d = 0
    return d

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
                  print sMesFile, i+1, datasetListTOGS.cases[i,4][0], nfoMal2008, TT.aTerr[j].malFile
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