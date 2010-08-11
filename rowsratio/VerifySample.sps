***************************************************************************
*   VerifySample.sps
***************************************************************************
* Проверка файлов выборок на полноту показателей, весов и данных
*  *  *  *
*  
*
*
***************************************************************************
* Последняя модификация: 25 мая 2010, Сергей Степанов
***************************************************************************

DEFINE @WORKDIR () 		'e:\Tmp\Spss\Rows' 	!ENDDEFINE.
DEFINE @SAMPLEMAL () 	 	'mal'		!ENDDEFINE.
DEFINE @SAMPLEMICRO ()  	micro		!ENDDEFINE.
DEFINE @Y2007 ()  		2007		!ENDDEFINE.
DEFINE @Y2008 ()  		2008		!ENDDEFINE.
DEFINE @Y2009 ()  		2009		!ENDDEFINE.

DEFINE @SAV ()  		'.sav'		!ENDDEFINE.
DEFINE @ListTOGS ()  		'\ListTOGS'	!ENDDEFINE.

DEFINE @NAMEDATE ()  		ДатаИзменения		!ENDDEFINE.

DEFINE @NOM ()  		номер		!ENDDEFINE.
DEFINE @VARSTRAT () 		ОКВЭД		!ENDDEFINE.

set MXLOOPS = 1200.
set MITERATE = 20000.
* set WORKSPACE = 1000000.
set RNG = MT MTINDEX = 362.
set decimal dot.
preserve.
CD @WORKDIR.
SHOW LOCALE.

/* OUTPUT SAVE OUTFILE = @WORKDIR + 'VerifySample.spv' . DOT
/*  Переменные и подготовка. 
/* INSERT FILE = @WORKDIR + '.sps' SYNTAX=BATCH ERROR=STOP.  

HOST COMMAND=['chcp 1251' 'chcp'
		'dir *.sav /b /l > listfile.txt'].

GET DATA  /TYPE = TXT
  /FILE = @WORKDIR + '\listfile.txt'
  /DELIMITERS = "\t\\ ,;"
  /VARIABLES = filetogs A20.
DATASET NAME filesTOGS.

get file = @WORKDIR + @ListTOGS + @SAV. 
DATASET NAME TOGS.
matrix. 
print /title='***  Модуль  VerifySample.sps - начало.  ***'. 
end matrix.

string #File (A20).
string #MessFile (A64).
string nfoMal2008 (a20). 
compute ДатаИзменения = $time. 

DO IF $CASENUM EQ 1.
  print "-----------------------". 
END IF. 

DO IF ОКАТО <= 99 . 
  compute nfoMal2008 = CONCAT(STRING(ОКАТО, N2), @SAMPLEMAL, "2008", @SAV).
  compute #File = nfoMal2008. 
  compute #MessFile = CONCAT("Отсутствует файл : ", #File). 
  compute #MessFile = CONCAT("Начало обработки файла : ", #File). 

  compute ПоДанным = #MessFile. 
END IF. 
save outfile = @WORKDIR + @ListTOGS + @SAV. 

/*WRITE OUTFILE = @WORKDIR + @ListTOGS + ".txt" /ALL.
EXECUTE.
DATASET CLOSE ALL.

MATRIX.

get ListTOGS
  /FILE = @WORKDIR + @ListTOGS + @SAV 
  /variables = Номер, ОКАТО, Территория, Кодокруга, Федеральныйокруг, @NAMEDATE, ПоДанным
  /names=VARNAMES. 

*print {ListTOGS(:,2)}
   / title = "Список территорий:"
   / clabels = "Территория"
   / format "A20". 

***     Размерность данных  и вспомогательные массивы.

compute nomTOGS = nrow(ListTOGS).	/* Количество записей территорий.  
compute FileMal8 = "mal". 		/* Имя файла выборки малых.    
compute FileMic8 = 'micro'. 		/* Имя файла выборки микропредприятий.    
compute strTOGS = '51'.
compute kolTOGS = 0.

LOOP i = 1 to nomTOGS.    
  DO IF (ListTOGS(i,2) <= 99). 
     compute kolTOGS = kolTOGS + 1.
*    print {i, ListTOGS(i,2)}    
	/clabels = "Номер", "Терр-ия" 
	/format "F5". 
  END IF. 
END LOOP. 
print kolTOGS /title = "Количество территорий". 
*SAVE ListTOGS
  /OUTFILE = @WORKDIR + @ListTOGS + @SAV
  /variables = Номер, ОКАТО, Территория, Кодокруга, Федеральныйокруг, @NAMEDATE, ПоДанным
  /strings = Территория, Федеральныйокруг, ПоДанным. 

END MATRIX.
restore.