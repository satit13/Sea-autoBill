
result = initial()
result = replaceInvoice_Paybilln()

GenHeader = GenData('paybill')
GenDetail = GenData('paybillsub')
DO insertLocalPaybill 
*DO PostDataToBc

PROCEDURE insertLocalPayBill 
	SELECT bcpaybill 
	ZAP 
	SELECT bcpaybillsub 
	ZAP 
	INSERT INTO bcpaybill SELECT * FROM bcpaybill_temp
	INSERT INTO bcpaybillsub SELECT * FROM bcpaybillsub_temp

ENDPROC 


FUNCTION GenData
LPARAMETERS tablename 
	wait wind 'Generate table '+tablename TIME(0.2)
	DO CASE  
		CASE ALLTRIM(UPPER(tablename))='PAYBILL'
			SELECT paybilldocno as docno , DATE() as docdate,arcode,'autobill' as creatorcode,DATETIME() as createdatetime,SUM(billbalance) as sumofinvoice,;
			0 as sumofdebitnote,0 as sumofcreditnote,0 as beforetaxamount,0 as taxamount,SUM(billbalance) as totalamount,'auto - bill ' as mydescription,;
			creditday,0 as billstatus,DATE()+creditday as duedate,1 as iscompletesave,SUM(billbalance) as billtemporary ;
			FROM invoice WHERE selected = 1 ;
			GROUP BY paybilldocno ,arcode,creditday ;
			INTO CURSOR bcpaybill_temp
		CASE ALLTRIM(UPPER(tablename))='PAYBILLSUB'
			SELECT paybilldocno as docno, DATE() as docdate,arcode,docdate as invoicedate,docno as invoiceno , billbalance as invbalance ,;
				billbalance as payamount,billbalance as  paybalance,billbalance as homeamount, 'auto - bill detail' as mydescription,0 as linenumber,;
				0 as billtype ;
				FROM invoice WHERE selected = 1 ;
				INTO CURSOR bcpaybillsub_temp
				
		otherwise 
			
	endcase 
	
ENDFUNC 






FUNC gennexno
  LPARAMETERS docprefix,no
  
  && check environment table 
  result =  initial()  
	IF !result
		MESSAGEBOX('fail initial phase',16,'please try again later')
		RETURN 
	ELSE 
		*MESSAGEBOX('Initial environment success',64,'')
	ENDIF 
	
	
	
  lc_nextno = ''

  && generate pattern prefix document 	
  lcmonth = ALLTRIM(STR(month(DATE())))
  lcmonth = IIF(LEN(lcmonth)=1,'0'+lcmonth,lcmonth)
  lcyear  = IIF(YEAR(DATE()) < 2500, right(ALLTRIM(STR(YEAR(DATE())+543)),2),right(ALLTRIM(STR(YEAR(DATE()))),2))
  lcDocKey = docprefix+lcyear+lcmonth+'-'
  ln_currno = no
	
  && Gen next number 	
  ln_nextno = ln_currno+1
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
  
  lc_nextno = lcDocKey+zero_prefix+ALLTRIM(STR(ln_nextno))
  
 && MESSAGEBOX(lc_nextno )f
 
  RETURN lc_nextno 
ENDFUNC
	


FUNCTION getmaxno

** return current max docno from prefix parameter

	lpara	keyprefix   && prefixword + yy+mm+'-'
	
	 && generate pattern prefix document 	
	  lcmonth = ALLTRIM(STR(month(DATE())))
	  lcmonth = IIF(LEN(lcmonth)=1,'0'+lcmonth,lcmonth)
	  lcyear  = IIF(YEAR(DATE()) < 2500, right(ALLTRIM(STR(YEAR(DATE())+543)),2),right(ALLTRIM(STR(YEAR(DATE()))),2))
	  lcDocKey = ALLTRIM(keyprefix)+lcyear+lcmonth+'-'
	  
	lccommand = "select CAST(max(RIGHT(RTRIM(docno),4)) as integer) as maxdocno from bcpaybill where docno like '%"+lcDocKey+"%'"
	
	result = SQLEXEC(dbconn,lccommand , 'current_docno')
	SELECT current_docno
	IF ISNULL(maxdocno) 
		lno  = 0
	ELSE 
		lno = maxdocno 
	ENDIF 
	RETURN lno
ENDFUNC 



* init table for gen paybill 
* required invoice table from main program 
FUNCTION  initial
	IF !USED('config') 
		USE config IN 0 SHARED
		
	ENDIF 

	IF !USED('invoice') 
		MESSAGEBOX('Cannot work because no Invoice not found')
		RETURN  .f. 
	ENDIF 

	IF !USED('bcpaybill')
		USE bcpaybill IN 0 exclusive 
		SELECT bcpaybill 
		SET ORDER TO tag docno 
	
	ENDIF 
	
	IF !USED('bcpaybillsub')
		USE bcpaybillsub IN 0 exclusive 
		SELECT bcpaybillsub 
		SET ORDER TO tag docno 
	ENDIF 
	
	SELECT bcpaybill
	
	*SET RELATION TO docno INTO Bcpaybillsub ADDITIVE
	RETURN .t. 
ENDFUNC


FUNCTION replaceInvoice_Paybilln
	maxno = getmaxno(ALLTRIM(config.pbprefix))
	lcdocno =  gennexno(ALLTRIM(config.pbprefix),maxno )
	CurAr = ALLTRIM(invoice.arcode)

	SELECT invoice 
	SET FILTER TO selected = 1
	GO TOP 
	DO WHILE !EOF()
	
			IF ALLTRIM(arcode)<> CurAr  && chance number paybill for arcode change
				CurAr = ALLTRIM(invoice.arcode)

					maxno = maxno+1
					lcdocno =  gennexno(ALLTRIM(config.pbprefix),maxno )

			ENDIF 
	*		IF invoice.selected = 1 
				replace invoice.paybilldocno  WITH lcdocno 
	*		ENDIF 
		
	SELECT invoice 
	SKIP 
	ENDDO 
	SELECT invoice 
		SET FILTER TO 
	RETURN .t.
 ENDFUNC 
 
 
 
 
 PROCEDURE PostDataToBc
 && Posting bcpaybill , sub to BC Account by connect "dbconn"
 
 
	SET PROCEDURE TO insert.prg
	 &&& Loop insert 
	 IF !USED('bcpaybill')
	 	MESSAGEBOX('Not found data : bcpaybill table',16,'Error not found data!!')
	 	RETURN .f. 
	 ENDIF 
	  IF !USED('bcpaybillsub')
	 	MESSAGEBOX('Not found data : bcpaybillsub table',16,'Error not found data!!')
	 	RETURN .f. 
	 ENDIF 
	 connChk = SQLTABLES(dbconn, "'VIEW', 'SYSTEM TABLE'", "mydbresult")
	 IF connchk <0 
	 	DBConn = -1
		SqlSetProp( 0 , 'DispLogin' , 3 ) 
		strconn = "driver=SQL Server;server="+ALLTRIM(config.servername)+";uid="+allt(config.username)+";"+;
							"pwd="+ALLTRIM(config.password)+";database="+ALLTRIM(config.databasename)
							MESSAGEBOX(strconn)
						
		DBConn = SQLSTRINGCONNECT(strconn)
	 ENDIF 
	 lccommand = "set dateformat dmy "
	 =SQLEXEC(dbconn,lccommand)
	 && begin send data 
	 SELECT bcpaybill 
	 GO TOP 
	 DO WHILE !EOF()
	 	WAIT WINDOW 'send data to bc : '+ALLTRIM(bcpaybill.docno) nowa
	 	lcdocno = ALLTRIM(bcpaybill.docno) 
	 	lccommand = "delete bcpaybill where docno = '"+lcdocno+"'"
	 	result=SQLEXEC(dbconn,lccommand)
	 	lccommand = "delete bcpaybillsub where docno = '"+lcdocno+"'"
	 	result=SQLEXEC(dbconn,lccommand)
	 	
	 	
	 	
	 	result = insertdata()   && insert bcpaybill 
	 	IF !result 
	 		DO errhand
	 	ENDIF 
	 	
	 	
	 		**  insert paybillsub 
		 	SELECT bcpaybillsub   && Insert Detail table 
		 	SEEK lcdocno 
		 	DO WHILE ALLTRIM(bcpaybillsub.docno) == ALLTRIM(lcdocno) AND !EOF()
		 		SELECT bcpaybillsub 
		 		result = insertdata()   && insert bcpaybill 
		 		
			 	IF !result 
			 		DO errhand
			 	ELSE 
			 		&& update invoice status 
			 		lccommand = "update bcarinvoice set paybillstatus =1 , paybillamount = 0 where docno = '"+ALLTRIM(bcpaybillsub.invoiceno)+"'"
			 		ivUpdate=SQLEXEC(dbconn,lccommand)
			 	ENDIF 
		 		SKIP 

		 	ENDDO 

		 	
	 	SELECT bcpaybill 
	 	SKIP 
	 	
	 ENDDO 
 	MESSAGEBOX('Send to BC Account Competed',64,'Completed')
 ENDPROC 