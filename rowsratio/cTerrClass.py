# !python.exe
# coding: cp1251
#
import SpssClient
import spss 

class cTerr(object):

    def __init__ (self, dataDir, mal, mic, nO, sT, kO, sO):
        #print globals()
        self.OKATO = "%02i" % nO        # ОКАТО 
        self.sTerr = sT                 # Название территории
        self.nOkrug = "%02i" % kO       # Код округа
        self.sOkrug = sO                # Название округа
        self.malFile = dataDir + '\\' + self.OKATO + mal + '.sav'    # Имя файла данных малых предприятий.
        self.micFile = dataDir + '\\' + self.OKATO + mic + '.sav'    # Имя файла данных микропредприятий.
        self.nameDataSet = 'D' + self.OKATO                       # Имя DATASET

        self.malTot2008 = [0,0,0,0,0,0,0,0,0,0,0,0.0]                 # Итоги по 2008 году
        self.malTot2009 = [0,0,0,0,0,0,0,0,0,0,0,0.0]                 # Итоги по 2009 году
        self.malTot2007 = [0,0,0,0,0,0,0,0,0,0,0,0.0]                 # Итоги по 2007 году
        self.malTos2007 = [0,0,0,0,0,0,0,0,0,0,0,0.0]                 # Итоги по 2007 году скорректированные 
        self.malcoeff   = [1,1,1,1,1,1,1,1,1,1,1,1.0]                 # Коэффициенты корреции
        self.malnnab     = 0                                      # Количество наблюдений
        self.malknab     = 0                                      # Количество корректных наблюдений
        self.malrnab     = 0                                      # Количество ошибочных наблюдений
        self.numvesmal   = 33           # Индекс веса у малых = номер переменной - 1
        self.malCHT      = 5            # Индекс средней численности
        self.malVIR      = 8            # Индекс выручки
        self.malCHTres   = 100          # Ограничение средней численности малых
        self.malVIRres   = 400000000    # Ограничение выручки малых

        self.micRatio2007 = [0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0]             # Доля микропредприятий по 2007 году
        self.malRatio2007 = [0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0]             # Доля малых предприятий по 2007 году

        self.micTot2008 = [0,0,0,0,0,0,0,0,0,0,0.0]                 # Итоги по 2008 году
        self.micTot2009 = [0,0,0,0,0,0,0,0,0,0,0.0]                 # Итоги по 2009 году
        self.micTot2007 = [0,0,0,0,0,0,0,0,0,0,0.0]                 # Итоги по 2007 году
        self.micTos2007 = [0,0,0,0,0,0,0,0,0,0,0.0]                 # Итоги по 2007 году скорректированные
        self.miccoeff   = [1,1,1,1,1,1,1,1,1,1,1.0]                 # Коэффициенты корреции
        self.micnnab     = 0                                      # Количество наблюдений
        self.micknab     = 0                                      # Количество корректных наблюдений
        self.micrnab     = 0                                      # Количество ошибочных наблюдений
        self.numvesmic   = 30          # Индекс веса у микро = номер переменной - 1
        self.micCHT      = 5            # Индекс средней численности
        self.micVIR      = 8            # Индекс выручки
        self.micCHTres   = 15           # Ограничение средней численности микропредприятий
        self.micVIRres   = 60000000     # Ограничение выручки микропредприятий

    def evalTotal(self, vM, mFile, numves, nnab, Tot2008, Tot2009, coeff, nCHT, nVIR, CHTres, VIRres, npok):
        def totalmal(self, rules, row, ves, total):
            if rules:
                total[0]  = total[0]  + ves                           # Количество предприятий
                for i in range(1, len(total)-1):
                    total[i]  = total[i]  + missVal(dmal.cases[row, npok[i-1][0]][0]) * ves
                
                total[len(total)-1] = total[len(total)-1] + 1                        # Количество корректных наблюдений

        def rules2008(self, row):
            """ Правила 2008 года  """
            return True
        
        def rules2009(self, row, nCHT, nVIR, CHTres, VIRres):
            """ Правила 2009 года  """
            if  missVal(dmal.cases[row,nCHT] [0]) <= CHTres and \
                missVal(dmal.cases[row,nVIR] [0]) <= VIRres :
                return True
            return False

        def rules2007(self, row):
            """ Правила 2007 года  """
            return True
            
        spss.Submit(r"""
        GET FILE = '%s'. 
        DATASET NAME %s.
        """ %(mFile, self.nameDataSet))
        with spss.DataStep():
            # По наблюдениям выборки
            dmal = spss.Dataset(name=self.nameDataSet)
            self.malknab = 0
            for i in range(len(dmal.cases)):
               ves = missVal(dmal.cases[i,numves][0])                # Вес предприятий
               if missVal(dmal.cases[i,0][0]) <> 0 and ves <> 0:         # По ненулевым ОКПО и ненулевым весам
                   totalmal(self, True,    i, ves, Tot2008)              # по правилам 2008 года
                   totalmal(self, rules2009(self, i, nCHT, nVIR, CHTres, VIRres), i, ves, Tot2009)              # по правилам 2009 года
                   #totalmal(self, rules2007(self, i), i, ves, self.malTot2007)              # по правилам 2007 года
                
        nnab = i                # Количество наблюдений
        
        for itot in range(len(Tot2009)):
            if Tot2008[itot] > 0:
                coeff[itot] = Tot2009[itot] / Tot2008[itot]
            else:
                coeff[itot] = 1
            #print itot, self.malcoeff[itot], self.malTot2009[itot], self.malTot2008[itot]
            
        spss.Submit(r"""
        DATASET CLOSE *. 
        """ )
