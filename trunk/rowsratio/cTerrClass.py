# !python.exe
# coding: cp1251
#
import SpssClient
import spss 

class cTerr(object):

    def __init__ (self, dataDir, mal, mic, nO, sT, kO, sO):
        #print globals()
        self.OKATO = "%02i" % nO        # ����� 
        self.sTerr = sT                 # �������� ����������
        self.nOkrug = "%02i" % kO       # ��� ������
        self.sOkrug = sO                # �������� ������
        self.malFile = dataDir + '\\' + self.OKATO + mal + '.sav'    # ��� ����� ������ ����� �����������.
        self.micFile = dataDir + '\\' + self.OKATO + mic + '.sav'    # ��� ����� ������ ����������������.
        self.nameDataSet = 'D' + self.OKATO                       # ��� DATASET

        self.malTot2008 = [0,0,0,0,0,0,0,0,0,0,0,0.0]                 # ����� �� 2008 ����
        self.malTot2009 = [0,0,0,0,0,0,0,0,0,0,0,0.0]                 # ����� �� 2009 ����
        self.malTot2007 = [0,0,0,0,0,0,0,0,0,0,0,0.0]                 # ����� �� 2007 ����
        self.malTos2007 = [0,0,0,0,0,0,0,0,0,0,0,0.0]                 # ����� �� 2007 ���� ����������������� 
        self.malcoeff   = [1,1,1,1,1,1,1,1,1,1,1,1.0]                 # ������������ ��������
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
        self.miccoeff   = [1,1,1,1,1,1,1,1,1,1,1.0]                 # ������������ ��������
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

        def rules2008(self, row):
            """ ������� 2008 ����  """
            return True
        
        def rules2009(self, row, nCHT, nVIR, CHTres, VIRres):
            """ ������� 2009 ����  """
            if  missVal(dmal.cases[row,nCHT] [0]) <= CHTres and \
                missVal(dmal.cases[row,nVIR] [0]) <= VIRres :
                return True
            return False

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
            for i in range(len(dmal.cases)):
               ves = missVal(dmal.cases[i,numves][0])                # ��� �����������
               if missVal(dmal.cases[i,0][0]) <> 0 and ves <> 0:         # �� ��������� ���� � ��������� �����
                   totalmal(self, True,    i, ves, Tot2008)              # �� �������� 2008 ����
                   totalmal(self, rules2009(self, i, nCHT, nVIR, CHTres, VIRres), i, ves, Tot2009)              # �� �������� 2009 ����
                   #totalmal(self, rules2007(self, i), i, ves, self.malTot2007)              # �� �������� 2007 ����
                
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
