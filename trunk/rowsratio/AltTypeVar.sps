GET FILE='E:\Tmp\Spss\Rows\03(�����).sav'. 
DATASET NAME Micro WINDOW=FRONT. 
* ������ �������� ����������. 
*���. 
ALTER TYPE      ���(F5.2).
VARIABLE LEVEL  ���(SCALE).
FORMATS         ���(F5.2). 
EXECUTE.
SAVE /OUTFILE = 'E:\Tmp\Spss\Rows\03(�����).sav'. 
DATASET CLOSE ALL.