GET FILE='E:\Tmp\Spss\Rows\03(микро).sav'. 
DATASET NAME Micro WINDOW=FRONT. 
* Задать свойства переменных. 
*ВЕС. 
ALTER TYPE      ВЕС(F5.2).
VARIABLE LEVEL  ВЕС(SCALE).
FORMATS         ВЕС(F5.2). 
EXECUTE.
SAVE /OUTFILE = 'E:\Tmp\Spss\Rows\03(микро).sav'. 
DATASET CLOSE ALL.