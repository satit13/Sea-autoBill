 CLOSE ALL
 CLEAR ALL 
 CLEAR ALL memory 
 
 PUBLIC DEFAULT_DIR ,dbconn1,dbconn2,gsShortDate
 
 
 LOCAL lcsys16,lcprogram,lcpath,lcolddir
	lcsys16=SYS(16,1)
	lcolddir=(SYS(5)+CURDIR())
	lcprogram = SUBSTR(lcsys16,AT(":",lcsys16)-1)
	CD LEFT(lcprogram,RAT("\",lcprogram))
	 SET DEFAULT TO (LEFT(lcprogram,RAT("\",lcprogram)))
	 DEFAULT_DIR = LEFT(lcprogram,RAT("\",lcprogram))
	 
	 
 SET CENTURY ON
 SET NEAR ON
 SET EXCLUSIVE OFF
 SET SAFETY OFF
 SET CPDIALOG OFF
 SET DELETED ON
 SET MULTILOCKS ON
  ON ERROR 
 SET CONSOLE ON
 SET EXACT ON 
 ON SHUTDOWN QUIT
   LCLASTSETTALK = SET('TALK')
 SET TALK OFF
 _screen.Icon=SYS(2003)+"\icons\Extensions.ico"
 _screen.Caption="GL-AUTOBILLING(Build 01-2016)"
 _screen.WindowState=2
 CSTARTPATH =FULLPATH(CURDIR())
 LCTXTFILE = 'INIT.TXT'






 STORE 0 TO dbconn1,dbconn2
 CD (CSTARTPATH)
 SET DEFAULT TO (CSTARTPATH)
 DEFAULT_DIR = CSTARTPATH


  SET SYSMENU TO 
  DO main.mpr 
  DO FORM "00-login"
  
  READ EVENTS 
 

 SET SYSMENU TO DEFAULT
 CLOSE DATABASES 
  set talk &lcLastSetTalk
  set path to &lcLastSetPath
*----------------------------------------------------------------*
PROCEDURE cGet
 PARAMETER CFILE , NNUM
 LOCAL _PATH
 _PATH = ''
 CD (CSTARTPATH)
 IF FILE(CFILE)
    SETUPFILE = FOPEN(CFILE,0)
    IF SETUPFILE < 0
        MESSAGEBOX('!äÁèÊÒÁÒÃ¶à»Ô´á¿éÁ  ' + FULLPATH(CURDIR()) + 'setup.txt file...',16,'Unable to Open file ' + CFILE)
    ELSE 
       STORE FSEEK(SETUPFILE,0,2) TO GNEND
       STORE FSEEK(SETUPFILE,0) TO GNTOP
       NCOUNT = 1
       DO WHILE NCOUNT <= NNUM
          A = FGETS(SETUPFILE)
          NCOUNT = NCOUNT + 1
       ENDDO 
       _PATH = SUBSTR(A,AT('=',A) + 1,LEN(ALLTRIM(A)) - AT('=',A))
    ENDIF 
    = FCLOSE(SETUPFILE)
    RETURN _PATH
 ELSE 
     MESSAGEBOX('ËÒá¿éÁ ' +CFILE+ CHR(13) + CHR(13) +  ;
   'Program being Cancel...!',16,'Error File Does not Exist .....')
    CANCEL 
    RETURN .F.
 ENDIF 
ENDPROC
*------*

PROCEDURE Getconnect
IF USED('init')
ELSE
	USE init IN 0 SHARED 
ENDIF 
SqlSetProp( 0 , 'DispLogin' , 3 ) 

		=SQLDISCONNECT(DBConn)
		
		WAIT WINDOW 'Reconnecting Source Server : '+ALLTRIM(init.servername) nowa
		DBConn1= SQLSTRINGCONNECT("driver=SQL Server;server="+ALLTRIM(config.sourceserver)+";uid="+allt(config.sourceusername)+";"+;
					"pwd="+ALLTRIM(config.sourcepassword)+";database="+ALLTRIM(config.sourcedatabase))
		IF Dbconn1 < 0
			MESSAGEBOX('Cannot connect database : '+ALLTRIM(init.database),16,'Error connecting database..')
		ENDIF 	
RETURN 

PROCEDURE errhand
   = AERROR(aErrorArray)  && Data from most recent error
*!*	   FOR n = 1 TO 7  && Display all elements of the array
*!*	      ? aErrorArray(n)
*!*	   ENDFOR
  MESSAGEBOX(aErrorArray(2))
RETURN 


*!* Let's get the CPU ID
FUNCTION GETCPUID 
	LOCAL lcComputerName, loWMI, lowmiWin32Objects, lowmiWin32Object
	lcComputerName = GETWORDNUM(SYS(0),1)
	loWMI = GETOBJECT("WinMgmts://" + lcComputerName)
	lowmiWin32Objects = loWMI.InstancesOf("Win32_Processor")
	FOR EACH lowmiWin32Object IN lowmiWin32Objects
		WITH lowmiWin32Object
		  *? "ProcessorId: " + TRANSFORM(.ProcessorId)
		  	 *?TRANSFORM(.ProcessorId)
		  	lccpuid = TRANSFORM(.ProcessorId)
		ENDWITH
	ENDFOR
	RETURN lccpuid 
ENDFUNC 
	





FUNCTION ENCRYPTx && Call this to encrypt password
PARAMETERS cPassword
cEncrypted_password = ''
FOR i = 1 TO LEN(cPassword)
cLetter = SUBSTR(cPassword, i, 1)
cEncrypted_password = cEncrypted_password + ;
CHR(MOD(ASC(cLetter)+10+i,255)) && Encryption formula
NEXT i
RETURN cEncrypted_password
ENDFUNC

FUNCTION DECRYPTx && Call this to Decrypt password
PARAMETERS cPassword
cUnencrypted_password = ''
FOR i = 1 TO LEN(cPassword)
cLetter = SUBSTR(cPassword, i, 1)
cUnencrypted_password = cUnencrypted_password + ;
CHR(MOD(ASC(cLetter)-10-i,255)) && Reverse of encryption formula
NEXT i
RETURN cUnencrypted_password
ENDFUNC