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

DO postInvoiceToBC && send to bcAccount


*** 

PROCEDURE genDocdateList 
	lnPeriodCount = tmpinvoice.period 
	
	CREATE CURSOR Docdatelist (period int,docdate date,duedate date,docprefix char(20),curno int)

	startdocdate =  tmpinvoice.firstPerioddate
	
	IF tmpinvoice.docdateby =1  && 1 = Lastday , 2 = every ... of month
 
		startdocdate = getDateByPeriod(startdocdate,'01')
	ENDIF 
	
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
		REPLACE ALL docdate WITH docdate-1   &&FOR period > 1 
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
		*MESSAGEBOX(lcCurdocno )
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
	
	IF USED('arlist')
		SELECT arlist 
		USE 
	ENDIF 
	
	IF USED('Docdatelist')
		SELECT Docdatelist  
		USE 
	ENDIF 

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
	
	IF !USED('bcoutputtax')
		USE bcoutputtax IN 0 EXCLUSIVE 
	ENDIF 
	
	IF !USED('config') 
		USE config IN 0 SHARED 
	ENDIF 

	
	SELECT bcarinvoicesub 
	zap 
	SET ORDER TO TAG Docno  IN Bcarinvoicesub
	
	SELECT bcoutputtax 
	ZAP 
	SET ORDER TO tag docno IN bcoutputtax 
	
	SELECT bcarinvoice 
	ZAP
	SET RELATION TO 
	SET RELATION TO docno INTO Bcarinvoicesub ADDITIVE
	SET RELATION TO docno INTO BcOutputtax ADDITIVE
	
	
	
ENDPROC 

**** Get ARList from ar condition
PROCEDURE getArList 
		arcode1 = ALLTRIM(tmpinvoice.arcode1)
		arcode2 = ALLTRIM(tmpinvoice.arcode2)
		lccommand = "select code,ISNULL(name1,'') as name1,ISNULL(billaddress,'') as billaddress,ISNULL(billcredit,0) as billcredit from bcar where code between '"+arcode1+"' and '"+arcode2+"'"
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
			 	lnBeforetaxamount = bcarinvoicex.beforetaxamount 
			 	lnTaxamount = bcarinvoicex.taxamount 
			 	lnExcepttaxamount = bcarinvoicex.excepttaxamount

			SELECT bcarinvoice 
				APPEND BLANK
				GATHER MEMVAR
				
				 && Gen next number 	
					  ln_nextno = docdatelist.curno+1
					  REPLACE docdatelist.curno WITH ln_nextno
					   DO CASE  
					  	CASE LEN(ALLTRIM(STR(ln_nextno)))=1 
					  		zero_prefix ='000'
					  	CASE LEN(ALLTRIM(STR(ln_nextno)))=2 
					  		zero_prefix ='00'	
					  	CASE LEN(ALLTRIM(STR(ln_nextno)))=3
					  		zero_prefix ='0'	
					  	OTHERWISE 
					  		zero_prefix =''
					  ENDCASE 
					  
					  lc_nextno = ALLTRIM(docdatelist.docprefix)+zero_prefix+ALLTRIM(STR(ln_nextno))
  				
				
				 
				REPLACE docno WITH lc_nextno 
				REPLACE taxno WITH lc_nextno 
				
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
			GO TOP 
				DO WHILE !EOF()
					SELECT bcarinvoicesubx
			 		SCATTER MEMVAR 
					SELECT bcarinvoicesub 
					APPEND BLANK
					GATHER MEMVAR
					
					REPLACE docno WITH lc_nextno 
					REPLACE taxno WITH lc_nextno 
					
					*REPLACE docno WITH 
					REPLACE arcode WITH ALLTRIM(arlist.code)
					*replace arname WITH ALLTRIM(arlist.name1)
					*replace araddress WITH ALLTRIM(arlist.billaddress)
					REPLACE docdate WITH docdatelist.docdate 
					SELECT bcarinvoicesubx
					SKIP 
				ENDDO 		  
		
		
			** Gen Outputtax 
				SELECT bcoutputtax 
				APPEND BLANK 
				REPLACE docno WITH lc_nextno 
				REPLACE taxno WITH lc_nextno 
				REPLACE docdate WITH docdatelist.docdate 
				REPLACE taxdate WITH docdatelist.docdate 
				REPLACE beforetaxamount WITH lnBeforetaxamount 
				REPLACE taxamount WITH lnTaxamount 
				REPLACE excepttaxamount WITH lnExcepttaxamount 
				REPLACE arcode WITH ALLTRIM(arlist.code)
							

			SELECT arlist 
			SKIP 
		ENDDO 
		
		SELECT docdateList 
		SKIP 1 
	ENDDO 
ENDPROC 




FUNCTION getmaxno

** return current max docno from prefix parameter

	lpara	keyprefix   && prefixword + yy+mm+'-'
  
	lccommand = "select CAST(max(RIGHT(RTRIM(docno),4)) as integer) as maxdocno from bcarinvoice where docno like '%"+keyprefix+"%'"
	
	result = SQLEXEC(dbconn,lccommand , 'current_docno')
	SELECT current_docno
	IF ISNULL(maxdocno) 
		lno  = 0
	ELSE 
		lno = maxdocno 
	ENDIF 
	RETURN lno
ENDFUNC 






**********************************
 
 PROCEDURE PostInvoiceToBc
 && Posting bcarinvoice , sub to BC Account by connect "dbconn"
 
 	SELECT bcarinvoice
 	
	SET PROCEDURE TO insert.prg
    lccommand = "set dateformat dmy "
	 =SQLEXEC(dbconn,lccommand)
	 && begin send data 
	 SELECT bcarinvoice 
	 GO TOP 
	 DO WHILE !EOF()
	 	WAIT WINDOW 'send data to bc : '+ALLTRIM(bcarinvoice.docno) nowa
	 	lcdocno = ALLTRIM(bcarinvoice.docno) 
	 	lccommand = "delete bcarinvoice where docno = '"+lcdocno+"'"
	 	&&MESSAGEBOX(lccommand)
	 	result=SQLEXEC(dbconn,lccommand)

	 	lccommand = "delete bcarinvoicesub where docno = '"+lcdocno+"'"
	 	&&MESSAGEBOX(lccommand)
	 	result=SQLEXEC(dbconn,lccommand)

	 	lccommand = "delete bcoutputtax where docno = '"+lcdocno+"'"
	 	result=SQLEXEC(dbconn,lccommand)
	 	&&MESSAGEBOX(lccommand)
	 	
	 	
	 	
	 	result = insertdata()   && insert bcpaybill 
	 	IF !result 
	 			&&MESSAGEBOX('insert bcarinvoice: '+lcdocno)
	 		DO errhand
	 	ENDIF 
	 	
	 	
	 		**  insert bcarinvoicesub 
		 	SELECT bcarinvoicesub  && Insert Detail table 
		 	SEEK lcdocno 
		 	DO WHILE ALLTRIM(bcarinvoicesub.docno) == ALLTRIM(lcdocno) AND !EOF()
		 		SELECT bcarinvoicesub
		 		result = insertdata()   && insert bcpaybill 
		 		
			 	IF !result 
			 		&&MESSAGEBOX('insert bcarinvoicesub: '+lcdocno)
			 		DO errhand
			 	ELSE 
			 		&& update invoice status 
			 		*lccommand = "update bcarinvoice set paybillstatus =1 , paybillamount = 0 where docno = '"+ALLTRIM(bcpaybillsub.invoiceno)+"'"
			 		*ivUpdate=SQLEXEC(dbconn,lccommand)
			 	ENDIF 
		 		SKIP 

		 	ENDDO 



	 		**  insert bcoutputtax 
		 	SELECT bcoutputtax && Insert Detail table 
		 	SEEK lcdocno 
		 	DO WHILE ALLTRIM(bcoutputtax .docno) == ALLTRIM(lcdocno) AND !EOF()
		 		SELECT bcoutputtax 
		 		result = insertdata()   && insert bcpaybill 
		 		
			 	IF !result 
			 		&&MESSAGEBOX('insert bcoutputtax : '+lcdocno)
			 		DO errhand
			 	ELSE 
			 		&& update invoice status 
			 		*lccommand = "update bcarinvoice set paybillstatus =1 , paybillamount = 0 where docno = '"+ALLTRIM(bcpaybillsub.invoiceno)+"'"
			 		*ivUpdate=SQLEXEC(dbconn,lccommand)
			 	ENDIF 
		 		SKIP 

		 	ENDDO 



		 	
	 	SELECT bcarinvoice
	 	SKIP 
	 	
	 ENDDO 
 	MESSAGEBOX('Send to BC Account Competed',64,'Completed')
 ENDPROC 