

******************************
FUNCTION   Insertdata 
	*Wait Window Alias()+'.'+Alltrim(Str(Recno())) Nowa
	nfield=Fcount()
	lccommand=" insert  into "+Alltrim(Alias())+" ("
	lFiled =""
	For initEnv = 1 To nfield
		IF UPPER(ALLTRIM(Fields(initEnv))) != 'ROWORDER'
			lFiled=lFiled+Iif(Len(Alltrim(lFiled))=0,'',',')+Fields(initEnv)
		ENDIF
	Endfor
	lccommand=lccommand+lFiled+" )"
	lccommand = lccommand+" values ("
	lFiled=''
	NextChar=''
	For initEnv = 1 To nfield
		Cflage=0
		Tflage=0
		NFlage=0
		MFlage=0
		IF UPPER(ALLTRIM(Fields(initEnv))) != 'ROWORDER'
			Do Case
			Case Type(Fields(initEnv)) = 'N'Or Type(Fields(initEnv)) = 'Y'   && Numeric or money
				lcx="Nextchar=ALLTRIM(STR("+Fields(initEnv)+",15,2))"
				Cflage=0
				Tflage=0
				NFlage=1
			Case Type(Fields(initEnv)) = 'C' Or Type(Fields(initEnv)) = 'M'    && Char or Memo field
				
				lcx="Nextchar=ALLTRIM("+Fields(initEnv)+")"
				Cflage=1
				Tflage=0
				NFlage=0
			Case Type(Fields(initEnv)) = 'T'      && Datetime field
				lcx="Nextchar=DTOC(TTOD("+Fields(initEnv)+"))"
				Cflage=0
				Tflage=1
				NFlage=0
			Case Type(Fields(initEnv)) = 'D'      && Datetime field
				lcx="Nextchar=DTOC("+Fields(initEnv)+")"
				Cflage=0
				Tflage=1
				NFlage=0
		
		Endcase
			&lcx
			&& กรอง field ที่เป็น char และมีอักษร '  อยู่ใน field ทำให้ insert date Error ออกไปก่อน
			IF !ISNULL(nextchar)  
				
					IF ("'"$nextchar)
						xnextchar=''
						FOR i = 1 TO LEN(ALLTRIM(nextchar))
							IF SUBSTR(nextchar,i,1)="'"
								xnextchar=xnextchar+' ' 
							ELSE 
								xnextchar=xnextchar+SUBSTR(nextchar,i,1)
							ENDIF 
						ENDFOR 
						*MESSAGEBOX( 'old : '+nextchar+CHR(13)+'New : '+xnextchar)

						nextchar=xnextchar
					ENDIF 
		
			ENDIF 
			
			Do Case
			Case Cflage=1
				If Isnull(NextChar)
					NextChar=''
				Endif
				lFiled=lFiled+"'"+NextChar+"'"
			Case NFlage=1
				If Isnull(NextChar)
					NextChar='0'
				Endif
				lFiled=lFiled+NextChar
			Case Tflage=1
				If Isnull(NextChar)
					NextChar='01/01/1900'
				Endif
				lFiled=lFiled+"'"+NextChar+"'"
			Endcase
			If initEnv < nfield
				lFiled=lFiled+","
			ENDIF
		ENDIF 
	Endfor
	lccommand=lccommand+lFiled+" )"
	MESSAGEBOX(lccommand)
	result=SQLEXEC(dbconn,lccommand)
		*IF ALLTRIM(UPPER(alias())) = 'VATOUT'
			*MESSAGEBOX(lccommand)
		*ENDIF 
	IF result < 0
		DO Errhand
		RETURN .f. 
	ELSE 
		RETURN .t.	
	ENDIF 
	
ENDFUNC 
*******
Procedure errhand
= Aerror(aErrorArray)  && Data from most recent error
MESSAGEBOX(aErrorArray(2))
Return aErrorArray(2)
Endproc
***************************
Procedure initEnv
	=SQLEXEC(b,"set dateformat dmy")
	=SQLEXEC(b,"SET LANGUAGE Thai")
	
Return

