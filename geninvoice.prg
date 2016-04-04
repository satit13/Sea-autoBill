**** Gen invoice from tmp_bcarinvoice 
**** Create bcarinvoice from template by condition from app



SET PROCEDURE TO lib.prg

DO setEnv 
DO setDefaultWorkFolder
IF  !getLogin() 
	MESSAGEBOX('Program  exiting...') 
	RETURN 
ENDIF 
DO getarlist
DO genDocdateList   && Create docdate list by period count
DO genInvoice   && Create invoice 






PROCEDURE genDocdateList 
	lnPeriodCount = tmpinvoice.period 
	
	CREATE CURSOR Docdatelist (period int,docdate date,duedate date,docprefix char(20),curno int)

	startdocdate =  tmpinvoice.firstPerioddate
	
	FOR i = 1 TO lnPeriodCount && Loop for period 
	
		INSERT INTO Docdatelist VALUES  (i,startdocdate ,DATE(),'',0)
		IF tmpinvoice.docdateby =1  && 1 = Lastday , 2 = every ... of month
			startdocdate = getDateByPeriod(startdocdate,'01')  && must be fix for get -1 date for gen lastdate
		ELSE 
			lcday = ALLTRIM(STR(tmpinvoice.everydate))
			lcday = IIF(LEN(lcday)=1 ,'0'+lcday,lcday)
			startdocdate = getDateByPeriod(startdocdate,lcday )
		ENDIF 
		
		&& get max by period 
		
		
		
	ENDFOR 
	IF tmpinvoice.docdateby =1  && replace for gen lastdate 
		REPLACE ALL docdate WITH docdate-1 FOR period > 1 
	ENDIF 



	&& FINDING  DOCNO BY DOCDATELIST.DOCDATE
	SELECT docdatelist 
	GO TOP 
	DO WHILE !EOF()
		lcprefix = ALLTRIM(config.ivprefix)
		lcyear = RIGHT(IIF(YEAR(docdate)<2500,ALLTRIM(STR(YEAR(docdate)+543)),ALLTRIM(STR(YEAR(docdate)))),2)
		lcmonth = IIF(MONTH(docdate)<10,'0'+ALLTRIM(STR(MONTH(docdate))),ALLTRIM(STR(MONTH(docdate))))
		lcprefix=lcprefix+lcyear+lcmonth+'-'
		
		
		&& PREPARE TO GEN INVOICE DOCNO 
		lnDBCurNo = ALLTRIM(STR(getmaxno(lcprefix)))
		REPLACE docdatelist.docprefix WITH lcprefix
		REPLACE docdatelist.curno WITH getmaxno(lcprefix)
		
			
		
		lcCurdocno = lcPrefix+lnDBCurNo
		MESSAGEBOX(lcCurdocno )
		SELECT docdatelist 
		SKIP 
	ENDDO 
	
ENDPROC 



FUNCTION getDateByPeriod 
	PARAMETERS curdate ,pday
	SET DATE DMY 
		lcday = pday
		lcNextMonth = ALLTRIM(STR(IIF((MONTH(curdate))<=11 , MONTH(curdate)+1,1)))
		lcyear = allt(STR(IIF(lcNextMonth=='1',YEAR(curdate)+1,YEAR(curdate))))
		lcNextDocdate = lcday+'/'+lcNextMonth+'/'+lcYear
		*MESSAGEBOX(CTOD(lcNextDocdate )) 
	RETURN CTOD(lcNextDocdate )
ENDFUNC 


PROCEDURE setEnv
	IF !USED('tmpinvoice')
		use tmpinvoice in 0 share 
	ENDIF 
	IF !USED('bcarinvoicex')
		USE bcarinvoicex IN 0 SHARED 
	ENDIF 
	IF !USED('bcarinvoicesubx')
		USE bcarinvoicesubx IN 0 SHARED 
	ENDIF 
	
	IF !USED('bcarinvoice')
		USE bcarinvoice IN 0 EXCLUSIVE
	ENDIF 
	
	IF !USED('bcarinvoicesub')
		USE bcarinvoicesub IN 0 EXCLUSIVE 
	ENDIF 
	
	IF !USED('config') 
		USE config IN 0 SHARED 
	ENDIF 
	
	IF USED('arlist')
		SELECT arlist 
		USE 
	ENDIF 
	
	IF USED('Docdatelist')
		SELECT Docdatelist  
		USE 
	ENDIF 
	
	
	SELECT bcarinvoice 
	ZAP 
	SELECT bcarinvoicesub 
	ZAP 
	
ENDPROC 


PROCEDURE getArList 
		arcode1 = ALLTRIM(tmpinvoice.arcode1)
		arcode2 = ALLTRIM(tmpinvoice.arcode2)
		lccommand = "select code,name1,billaddress,billcredit from bcar where code between '"+arcode1+"' and '"+arcode2+"'"
		*MESSAGEBOX(lccommand)
		
		result = SQLEXEC(dbconn,lccommand,'arlist')
		
ENDPROC 


FUNCTION  getLogin 
	PUBLIC dbconn 
	dbconn = -1
	lcservername = ALLTRIM(config.servername)
	lcuser = ALLTRIM(config.username)
	lcpass = ALLTRIM(config.password)
	lcdatabase = ALLTRIM(config.dadabasename)
	SqlSetProp( 0 , 'DispLogin' , 3 ) 
	strconn = "driver=SQL Server;server="+lcservername+";uid="+lcuser+";"+;
						"pwd="+lcpass+";database="+lcdatabase
						*MESSAGEBOX(strconn)
						
	dbconn = SQLSTRINGCONNECT(strconn)
	IF dbconn < 0
		MESSAGEBOX('Error to connect database ',16,'Cannot connect db ')
		RETURN .f.
	ELSE 
		RETURN .t.
	ENDIF 
	
ENDPROC 



PROCEDURE genInvoice 

	&&lcPrefix = ALLTRIM(config.ivprefix)
	
	lpara	keyprefix   && prefixword + yy+mm+'-'
	

	SELECT docdatelist &&  Peroid by docdate list
	GO TOP 
	DO WHILE !EOF()
		
		SELECT arlist   &&& Loop gen by arcode list 
		GO TOP 
		DO WHILE !EOF()
			
			WAIT WINDOW RECNO() NOWAIT  
			** GEN INVOICE 
			SELECT bcarinvoicex
			 	SCATTER MEMVAR 
			SELECT bcarinvoice 
				APPEND BLANK
				GATHER MEMVAR
				
				 
				*REPLACE docno WITH 
				REPLACE arcode WITH ALLTRIM(arlist.code)
				replace arname WITH ALLTRIM(arlist.name1)
				replace araddress WITH ALLTRIM(arlist.billaddress)
				REPLACE docdate WITH docdatelist.docdate 
				IF tmpinvoice.creditdayby = 1 && fix credit  
					REPLACE duedate WITH docdatelist.docdate + tmpinvoice.creditday
				ELSE && by customer credit
					REPLACE  duedate WITH docdatelist.docdate+ arlist.billcredit
				ENDIF  
				
			
			** GEN INVOICESUB 
			SELECT bcarinvoicesubx
			 	SCATTER MEMVAR 
			SELECT bcarinvoicesub 
				APPEND BLANK
				GATHER MEMVAR
				
				 
				*REPLACE docno WITH 
				REPLACE arcode WITH ALLTRIM(arlist.code)
				*replace arname WITH ALLTRIM(arlist.name1)
				*replace araddress WITH ALLTRIM(arlist.billaddress)
				REPLACE docdate WITH docdatelist.docdate 
				  
				
			SELECT arlist 
			SKIP 
		ENDDO 
		
		SELECT docdateList 
		SKIP 1 
	ENDDO 
ENDPROC 