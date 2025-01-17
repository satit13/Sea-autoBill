
FUNCTION getmaxno

** return current max docno from prefix parameter

	lpara	keyprefix   && prefixword + yy+mm+'-'
  
	lccommand = "select CAST(max(RIGHT(RTRIM(docno),4)) as integer) as maxdocno from bcpaybill where docno like '%"+keyprefix+"%'"
	
	result = SQLEXEC(dbconn,lccommand , 'current_docno')
	SELECT current_docno
	IF ISNULL(maxdocno) 
		lno  = 0
	ELSE 
		lno = maxdocno 
	ENDIF 
	RETURN lno
ENDFUNC 


PROCEDURE setDefaultWorkFolder
 LOCAL lcsys16,lcprogram,lcpath,lcolddir
	lcsys16=SYS(16,1)
	lcolddir=(SYS(5)+CURDIR())
	lcprogram = SUBSTR(lcsys16,AT(":",lcsys16)-1)
	CD LEFT(lcprogram,RAT("\",lcprogram))
	 SET DEFAULT TO (LEFT(lcprogram,RAT("\",lcprogram)))
	 DEFAULT_DIR = LEFT(lcprogram,RAT("\",lcprogram))
	 
	 
ENDPROC 