CLEAR
DO GetTableList

*	DO Delete_Exists_docno
Do initEnv
*	DO trf_BCAR
*	DO Trf_BCAP
*	Do Trf_BCItem
*	DO trf_bcarinvoice
*	DO trf_bcarinvoicesub
*	DO trf_bcardeposit
*	DO trf_bcardeposituse
*	DO Trf_bccreditnote
*	DO trf_bcCreditnotesub
*	DO trf_bcInvCreditnote
*	DO trf_bcDebitnote1
*	DO Trf_bcdebitnotesub1
*	DO Trf_bcinvdebitnote1
*	DO trf_bcchqin
*	DO trf_bcRecmoney
*	DO trf_bcCreditcard
	DO trf_bcOutputtax
**********************************
Procedure trf_bcarinvoice
=SQLEXEC(a,"select * FROM trf.dbo.bcarinvoice ",'bcarinvoice')

Select bcarinvoice
result=SQLEXEC(b,"SET IDENTITY_INSERT "+Alias()+" OFF ")
Go Top
Do While !Eof()
	Wait Window Alltrim(Str(Recno())) Nowa
	Do Insertdata
	Skip
Enddo

Endproc
*********************************
Procedure trf_bcarinvoicesub
* run fix itemname data for ' in name of transaction Replace with ''
lccommand=	"Update TRF.dbo.bcarinvoicesub set itemname =  REPLACE(itemname,'''','') where  itemname like '%''%'"
result=SQLEXEC(a,lccommand)
If result<0
	Do errhand
Endif
=SQLEXEC(a,"select * FROM trf.dbo.bcarinvoicesub ",'bcarinvoicesub')

Select bcarinvoicesub
result=SQLEXEC(b,"SET IDENTITY_INSERT "+Alias()+" OFF ")
Go Top
Do While !Eof()
	Wait Window Alltrim(Str(Recno())) Nowa
	Do Insertdata
	Skip
Enddo


Endproc

*********************************
Procedure trf_bcreceipt1
=SQLEXEC(a,"select * FROM trf.dbo.bcreceipt1",'bcreceipt1')

Select bcreceipt1
result=SQLEXEC(b,"SET IDENTITY_INSERT "+Alias()+" OFF ")
Go Top
Do While !Eof()
	Wait Window Alltrim(Str(Recno())) Nowa
	Do Insertdata
	Skip
Enddo

Endproc
*********************************
Procedure trf_bcreceiptsub1
=SQLEXEC(a,"select * FROM trf.dbo.bcreceiptsub1",'bcreceiptsub1')

Select bcreceiptsub1
result=SQLEXEC(b,"SET IDENTITY_INSERT "+Alias()+" OFF ")
Go Top
Do While !Eof()
	Wait Window Alltrim(Str(Recno())) Nowa
	Do Insertdata
	Skip
Enddo

Endproc
***********************************

Procedure trf_bcardeposit
=SQLEXEC(a,"select * FROM trf.dbo.bcardeposit",'bcardeposit')

Select bcardeposit
result=SQLEXEC(b,"SET IDENTITY_INSERT "+Alias()+" OFF ")
Go Top
Do While !Eof()
	Wait Window Alltrim(Str(Recno())) Nowa
	Do Insertdata
	Skip
Enddo

Endproc
*********************************
Procedure trf_bcardeposituse
=SQLEXEC(a,"select * FROM trf.dbo.bcardeposituse",'bcardeposituse')

Select bcardeposituse
result=SQLEXEC(b,"SET IDENTITY_INSERT "+Alias()+" OFF ")
Go Top
Do While !Eof()
	Wait Window Alltrim(Str(Recno())) Nowa
	Do Insertdata
	Skip
Enddo


Endproc
*********************************
Procedure trf_bcCreditnote
=SQLEXEC(a,"select * FROM trf.dbo.bcCreditnote",'bcCreditnote')

Select bcCreditnote
result=SQLEXEC(b,"SET IDENTITY_INSERT "+Alias()+" OFF ")
Go Top
Do While !Eof()
	Wait Window Alltrim(Str(Recno())) Nowa
	Do Insertdata
	Skip
Enddo

Endproc
*********************************
Procedure trf_bcCreditnotesub
=SQLEXEC(a,"select * FROM trf.dbo.bcCreditnotesub",'bcCreditnotesub')

Select bcCreditnotesub
result=SQLEXEC(b,"SET IDENTITY_INSERT "+Alias()+" OFF ")
Go Top
Do While !Eof()
	Wait Window Alltrim(Str(Recno())) Nowa
	Do Insertdata
	Skip
Enddo

Endproc
*********************************
Procedure trf_bcinvCreditnote
=SQLEXEC(a,"select * FROM trf.dbo.bcinvCreditnote",'bcinvCreditnote')

Select bcinvCreditnote
result=SQLEXEC(b,"SET IDENTITY_INSERT "+Alias()+" OFF ")
Go Top
Do While !Eof()
	Wait Window Alltrim(Str(Recno())) Nowa
	Do Insertdata
	Skip
Enddo

Endproc
*********************************
Procedure trf_bcDebitnote1
=SQLEXEC(a,"select * FROM trf.dbo.bcDebitnote1",'bcDebitnote1')

Select bcDebitnote1
result=SQLEXEC(b,"SET IDENTITY_INSERT "+Alias()+" OFF ")
Go Top
Do While !Eof()
	Wait Window Alltrim(Str(Recno())) Nowa
	Do Insertdata
	Skip
Enddo

Endproc
*********************************
Procedure trf_bcDebitnotesub1
=SQLEXEC(a,"select * FROM trf.dbo.bcDebitnotesub1 ",'bcDebitnotesub1')

Select bcDebitnotesub1
result=SQLEXEC(b,"SET IDENTITY_INSERT "+Alias()+" OFF ")
Go Top
Do While !Eof()
	Wait Window Alltrim(Str(Recno())) Nowa
	Do Insertdata
	Skip
Enddo

Endproc
*********************************
Procedure trf_bcinvDebitnote1
=SQLEXEC(a,"select * FROM trf.dbo.bcinvDebitnote1 ",'bcinvDebitnote1')

Select bcinvDebitnote1
result=SQLEXEC(b,"SET IDENTITY_INSERT "+Alias()+" OFF ")
Go Top
Do While !Eof()
	Wait Window Alltrim(Str(Recno())) Nowa
	Do Insertdata
	Skip
Enddo

Endproc
*********************************
Procedure trf_bcChqin
=SQLEXEC(a,"select * FROM trf.dbo.bcChqin",'bcChqin')

Select bcChqin
result=SQLEXEC(b,"SET IDENTITY_INSERT "+Alias()+" OFF ")
Go Top
Do While !Eof()
	Wait Window Alltrim(Str(Recno())) Nowa
	Do Insertdata
	Skip
Enddo

Endproc
*********************************
Procedure trf_bcCreditcard
=SQLEXEC(a,"select * FROM trf.dbo.bcCreditcard",'bcCreditcard')

Select bcCreditcard
result=SQLEXEC(b,"SET IDENTITY_INSERT "+Alias()+" OFF ")
Go Top
Do While !Eof()
	Wait Window Alltrim(Str(Recno())) Nowa
	Do Insertdata
	Skip
Enddo

Endproc
*********************************
Procedure trf_BCOutputtax
=SQLEXEC(a,"select * FROM trf.dbo.BCOutputtax",'BCOutputtax')

Select BCOutputtax
result=SQLEXEC(b,"SET IDENTITY_INSERT "+Alias()+" OFF ")
Go Top
Do While !Eof()
	Wait Window Alltrim(Str(Recno())) Nowa
	Do Insertdata
	Skip
Enddo

Endproc
*********************************
Procedure trf_BCRecmoney
=SQLEXEC(a,"select * FROM trf.dbo.BCRecmoney",'BCRecmoney')

Select BCRecmoney
result=SQLEXEC(b,"SET IDENTITY_INSERT "+Alias()+" OFF ")
Go Top
Do While !Eof()
	Wait Window Alltrim(Str(Recno())) Nowa
	Do Insertdata
	Skip
Enddo

Endproc
*********************************
Procedure trf_bcardeposit
=SQLEXEC(a,"select * FROM trf.dbo.bcardeposit",'bcardeposit')

Select bcardeposit
result=SQLEXEC(b,"SET IDENTITY_INSERT "+Alias()+" OFF ")
Go Top
Do While !Eof()
	Wait Window Alltrim(Str(Recno())) Nowa
	Do Insertdata
	Skip
Enddo

Endproc
************************************
Procedure Trf_BCItem
lccommand=	"Update bcitem set name1=REPLACE(name1,'''','-'), name2=REPLACE(name2,'''','') where  name1 like '%''%' or name2 like '%''%'"
result=SQLEXEC(a,lccommand)
If result<0
	Do errhand
Endif
=SQLEXEC(a,"select * from bcitem where code in (select code from trf.dbo.NP_Transfer_Newitem)",'bcitem')

* Generate Roworder สำหรับ โอนไปปลายทาง
lccommand = "select  ISNULL(MAX(roworder) ,0)as nextroworder from  "+Alltrim(targetdb)+".dbo.bcitem"
=SQLEXEC(b,lccommand,'NewRow')
lnNext=NewRow.nextroworder+10
Select bcitem
Go Top
Do While !Eof()
	Replace roworder With lnNext
	lnNext=lnNext+1
	Select bcitem
	Skip
Enddo
Select  NewRow
Use

* เรียก Procedure เพื่อ insert ข้อมูลปลายทาง
Select bcitem
result=SQLEXEC(b,"SET IDENTITY_INSERT "+Alias()+" OFF ")
Go Top
Do While !Eof()
	Wait Window ' ทะเบียนสินค้าใหม่ :'+Alltrim(Str(Recno())) Nowa
	Do Insertdata
	Skip
Enddo



** Insert New Stkpacking
=SQLEXEC(a,"select * from bcstkpacking where itemcode in (select code from trf.dbo.NP_Transfer_Newitem)",'bcstkpacking')


* Generate Roworder สำหรับ โอนไปปลายทาง
lccommand = "select  ISNULL(MAX(roworder) ,0)as nextroworder from  "+Alltrim(targetdb)+".dbo.bcstkpacking "
=SQLEXEC(b,lccommand,'NewRow')
lnNext=NewRow.nextroworder+10
Select bcstkpacking
Go Top
Do While !Eof()
	Replace roworder With lnNext
	lnNext=lnNext+1
	Select bcstkpacking
	Skip
Enddo
Select  NewRow
Use

* เรียก Procedure เพื่อ insert ข้อมูลปลายทาง
Select bcstkpacking
result=SQLEXEC(b,"SET IDENTITY_INSERT "+Alias()+" OFF ")
Go Top
Do While !Eof()
	Wait Window 'ตารางหลายหน่วยนับ : '+Alltrim(Str(Recno())) Nowa
	Do Insertdata
	Skip
Enddo



** Insert New BCITEMWAREHOUSE
=SQLEXEC(a,"select * from bcitemwarehouse where whcode in ('010','011') and itemcode in (select code from trf.dbo.NP_Transfer_Newitem)",'bcitemwarehouse')

* Generate Roworder สำหรับ โอนไปปลายทาง
lccommand = "select  ISNULL(MAX(roworder) ,0)as nextroworder from  "+Alltrim(targetdb)+".dbo.bcitemwarehouse "
=SQLEXEC(b,lccommand,'NewRow')
lnNext=NewRow.nextroworder+10
Select bcitemwarehouse
Go Top
Do While !Eof()
	Replace roworder With lnNext
	lnNext=lnNext+1
	Select bcitemwarehouse
	Skip
Enddo
Select  NewRow
Use

* เรียก Procedure เพื่อ insert ข้อมูลปลายทาง
Select bcitemwarehouse
result=SQLEXEC(b,"SET IDENTITY_INSERT "+Alias()+" OFF ")
Go Top
Do While !Eof()
	Wait Window 'ตารางคลังและที่เก็บสินค้า :' +Alltrim(Str(Recno())) Nowa
	Do Insertdata
	Skip
Enddo




Return .T.
Endproc






*********************************
Procedure trf_BCAP
lccommand=	"Update bcap set name1=REPLACE(name1,'''','-'), name2=REPLACE(name2,'''','-') where  name1 like '%''%' or name2 like '%''%'"
result=SQLEXEC(a,lccommand)
If result<0
	Do errhand
Endif

=SQLEXEC(a,"select * from bcap where code in (select code from trf.dbo.NP_Transfer_NewAP)",'BCAP')
*	=SQLEXEC(B,"delete BCAP   ")
Select BCAP
result=SQLEXEC(b,"SET IDENTITY_INSERT "+Alias()+" OFF ")
Go Top
Do While !Eof()
	Wait Window Alltrim(Str(Recno())) Nowa
	Do Insertdata
	Skip
Enddo

Endproc
*********************************
Procedure trf_BCAR
lccommand=	"Update bcar set name1=REPLACE(name1,'''','-'), name2=REPLACE(name2,'''','-') where  name1 like '%''%' or name2 like '%''%'"
result=SQLEXEC(a,lccommand)
If result<0
	Do errhand
Endif

=SQLEXEC(a,"select * from bcar where code in (select code from trf.dbo.NP_Transfer_NewAR)",'BCAR')
*	=SQLEXEC(B,"delete BCAP   ")
Select BCAR
result=SQLEXEC(b,"SET IDENTITY_INSERT "+Alias()+" OFF ")
Go Top
Do While !Eof()
	Wait Window Alltrim(Str(Recno())) Nowa
	Do Insertdata
	Skip
Enddo

Endproc
********************************

Procedure trf_bcchqindeposit
=SQLEXEC(a,"select * FROM trf.dbo.bcchqindeposit",'bcchqindeposit')

Select bcchqindeposit
result=SQLEXEC(b,"SET IDENTITY_INSERT "+Alias()+" OFF ")
Go Top
Do While !Eof()
	Wait Window Alltrim(Str(Recno())) Nowa
	Do Insertdata
	Skip
Enddo

Endproc

****************************

Procedure trf_bcchqindeposSub
=SQLEXEC(a,"select * FROM trf.dbo.bcchqindeposSub",'bcchqindeposSub')

Select bcchqindeposSub
result=SQLEXEC(b,"SET IDENTITY_INSERT "+Alias()+" OFF ")
Go Top
Do While !Eof()
	Wait Window Alltrim(Str(Recno())) Nowa
	Do Insertdata
	Skip
Enddo

Endproc
****************************

Procedure TRF_BCAPINVOICE
=SQLEXEC(a,"select * FROM trf.dbo.BCAPINVOICE",'BCAPINVOICE')

Select BCAPINVOICE
result=SQLEXEC(b,"SET IDENTITY_INSERT "+Alias()+" OFF ")
Go Top
Do While !Eof()
	Wait Window Alltrim(Str(Recno())) Nowa
	Do Insertdata
	Skip
Enddo

ENDPROC


****************************

Procedure TRF_BCAPINVOICESUB 
=SQLEXEC(a,"select * FROM trf.dbo.BCAPINVOICESUB ",'BCAPINVOICESUB ')

Select BCAPINVOICESUB 
result=SQLEXEC(b,"SET IDENTITY_INSERT "+Alias()+" OFF ")
Go Top
Do While !Eof()
	Wait Window Alltrim(Str(Recno())) Nowa
	Do Insertdata
	Skip
Enddo

Endproc

****************************

Procedure TRF_BCIRSUB 
=SQLEXEC(a,"select * FROM trf.dbo.BCIRSUB ",'BCIRSUB ')

Select BCIRSUB 
result=SQLEXEC(b,"SET IDENTITY_INSERT "+Alias()+" OFF ")
Go Top
Do While !Eof()
	Wait Window Alltrim(Str(Recno())) Nowa
	Do Insertdata
	Skip
Enddo

Endproc
****************************

Procedure TRF_BCINPUTTAX
=SQLEXEC(a,"select * FROM trf.dbo.BCINPUTTAX",'BCINPUTTAX')

Select BCINPUTTAX
result=SQLEXEC(b,"SET IDENTITY_INSERT "+Alias()+" OFF ")
Go Top
Do While !Eof()
	Wait Window Alltrim(Str(Recno())) Nowa
	Do Insertdata
	Skip
Enddo

Endproc
****************************

Procedure TRF_BCAPWTAXLIST
=SQLEXEC(a,"select * FROM trf.dbo.BCAPWTAXLIST",'BCAPWTAXLIST')

Select BCAPWTAXLIST
result=SQLEXEC(b,"SET IDENTITY_INSERT "+Alias()+" OFF ")
Go Top
Do While !Eof()
	Wait Window Alltrim(Str(Recno())) Nowa
	Do Insertdata
	Skip
Enddo

ENDPROC
****************************

Procedure TRF_BCPAYMENT
=SQLEXEC(a,"select * FROM trf.dbo.BCPAYMENT",'BCPAYMENT')

Select BCPAYMENT
result=SQLEXEC(b,"SET IDENTITY_INSERT "+Alias()+" OFF ")
Go Top
Do While !Eof()
	Wait Window Alltrim(Str(Recno())) Nowa
	Do Insertdata
	Skip
Enddo

ENDPROC
****************************

Procedure TRF_BCPAYMENTSUB
=SQLEXEC(a,"select * FROM trf.dbo.BCPAYMENTSUB",'BCPAYMENTSUB')

Select BCPAYMENTSUB
result=SQLEXEC(b,"SET IDENTITY_INSERT "+Alias()+" OFF ")
Go Top
Do While !Eof()
	Wait Window Alltrim(Str(Recno())) Nowa
	Do Insertdata
	Skip
Enddo

ENDPROC

****************************

Procedure TRF_BCPAYMONEY
=SQLEXEC(a,"select * FROM trf.dbo.BCPAYMONEY",'BCPAYMONEY')

Select BCPAYMONEY
result=SQLEXEC(b,"SET IDENTITY_INSERT "+Alias()+" OFF ")
Go Top
Do While !Eof()
	Wait Window Alltrim(Str(Recno())) Nowa
	Do Insertdata
	Skip
Enddo

ENDPROC
****************************

Procedure TRF_BCCHQOUT
=SQLEXEC(a,"select * FROM trf.dbo.BCCHQOUT",'BCCHQOUT')

Select BCCHQOUT
result=SQLEXEC(b,"SET IDENTITY_INSERT "+Alias()+" OFF ")
Go Top
Do While !Eof()
	Wait Window Alltrim(Str(Recno())) Nowa
	Do Insertdata
	Skip
Enddo

ENDPROC
****************************

Procedure TRF_BCAPDEPOSIT
=SQLEXEC(a,"select * FROM trf.dbo.BCAPDEPOSIT",'BCAPDEPOSIT')

Select BCAPDEPOSIT
result=SQLEXEC(b,"SET IDENTITY_INSERT "+Alias()+" OFF ")
Go Top
Do While !Eof()
	Wait Window Alltrim(Str(Recno())) Nowa
	Do Insertdata
	Skip
Enddo

ENDPROC
****************************

Procedure TRF_BCAPDEPOSITUSE
=SQLEXEC(a,"select * FROM trf.dbo.BCAPDEPOSITUSE",'BCAPDEPOSITUSE')

Select BCAPDEPOSITUSE
result=SQLEXEC(b,"SET IDENTITY_INSERT "+Alias()+" OFF ")
Go Top
Do While !Eof()
	Wait Window Alltrim(Str(Recno())) Nowa
	Do Insertdata
	Skip
Enddo

ENDPROC
*******************************

Procedure TRF_BCAPDEPOSITSPECIAL
=SQLEXEC(a,"select * FROM trf.dbo.BCAPDEPOSITSPECIAL",'BCAPDEPOSITSPECIAL')

Select BCAPDEPOSITSPECIAL
result=SQLEXEC(b,"SET IDENTITY_INSERT "+Alias()+" OFF ")
Go Top
Do While !Eof()
	Wait Window Alltrim(Str(Recno())) Nowa
	Do Insertdata
	Skip
Enddo

ENDPROC
*******************************
Procedure TRF_BCChqoutpass
=SQLEXEC(a,"select * FROM trf.dbo.BCChqoutpass",'BCChqoutpass')

Select BCChqoutpass
result=SQLEXEC(b,"SET IDENTITY_INSERT "+Alias()+" OFF ")
Go Top
Do While !Eof()
	Wait Window Alltrim(Str(Recno())) Nowa
	Do Insertdata
	Skip
Enddo

ENDPROC
******************************
Procedure TRF_BCChqoutpasssub
=SQLEXEC(a,"select * FROM trf.dbo.BCChqoutpass",'BCChqoutpasssub')

Select BCChqoutpasssub
result=SQLEXEC(b,"SET IDENTITY_INSERT "+Alias()+" OFF ")
Go Top
Do While !Eof()
	Wait Window Alltrim(Str(Recno())) Nowa
	Do Insertdata
	Skip
Enddo

ENDPROC
******************************
Procedure TRF_BCOtherExpense
=SQLEXEC(a,"select * FROM trf.dbo.BCChqoutpass",'BCOtherExpense')

Select BCOtherExpense
result=SQLEXEC(b,"SET IDENTITY_INSERT "+Alias()+" OFF ")
Go Top
Do While !Eof()
	Wait Window Alltrim(Str(Recno())) Nowa
	Do Insertdata
	Skip
Enddo

ENDPROC
******************************
Procedure Insertdata
	Wait Window Alias()+'.'+Alltrim(Str(Recno())) Nowa
	nfield=Fcount()
	lccommand=" insert "+Alltrim(Alias())+" ("
	lFiled =""
	For initEnv = 1 To nfield
		IF UPPER(ALLTRIM(Fields(initEnv))) != 'ROWORDER'
		lFiled=lFiled+Iif(Len(Alltrim(lFiled))=0,'',',')+Fields(initEnv)
		ENDIF 
	Endfor
	lccommand=lccommand+lFiled+" )"
	*MESSAGEBOX(lccommand)
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
			Case Type(Fields(initEnv)) = 'N'Or Type(Fields(initEnv)) = 'Y'
				lcx="Nextchar=ALLTRIM(STR("+Fields(initEnv)+",15,2))"
				Cflage=0
				Tflage=0
				NFlage=1
			Case Type(Fields(initEnv)) = 'C' Or Type(Fields(initEnv)) = 'M'
				lcx="Nextchar=ALLTRIM("+Fields(initEnv)+")"
				Cflage=1
				Tflage=0
				NFlage=0
			Case Type(Fields(initEnv)) = 'T'
				lcx="Nextchar=DTOC(TTOD("+Fields(initEnv)+"))"
				Cflage=0
				Tflage=1
				NFlage=0

			Endcase
			&lcx
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
	result=SQLEXEC(b,lccommand)
	
	lcGenStr="lnRoworder="+Alias()+".roworder"
	&lcGenStr
	lcTable=Alias()

	Do Case
		Case  Alltrim(Upper(Alias()))=='BCINVDEBITNOTE1'
			lcdocno=Alltrim(debitnoteno)+'-'+Alltrim(invoiceno)
		Case Alltrim(Upper(Alias()))== 'BCINVCREDITNOTE'
			lcdocno=Alltrim(creditnoteno)+'-'+Alltrim(invoiceno)
		Case 	Alltrim(Upper(Alias()))=='BCAR' Or  Alltrim(Upper(Alias()))=='BCAP'  Or  Alltrim(Upper(Alias()))=='BCITEM'
			lcdocno = Alltrim(Code)
		Case Alltrim(Upper(Alias()))=='BCSTKPACKING'
			lcdocno=Alltrim(itemcode)+'-'+Alltrim(unitcode)
		Case  Alltrim(Upper(Alias()))=='BCSTKITEMWAREHOUSE'
			lcdocno=Alltrim(itemcode)+'-'+Alltrim(whcode)
		Case  Alltrim(Upper(Alias()))=='BCITEMWAREHOUSE'
			lcdocno=Alltrim(itemcode)+'-'+Alltrim(whcode)
		Otherwise
			lcdocno=docno
	Endcase

	If result<0   && Error Sql
		lcError =  errhand ()
		Insert Into trflog (trfno,roworder,Table,Complete,Errdesc,docno,trfdate) Values (mtrfno,lnRoworder,lcTable,.F.,lcError,lcdocno,Datetime())
		Return .F.
	Else
		Insert Into trflog (trfno,roworder,Table,Complete,Errdesc,docno,trfdate) Values (mtrfno,lnRoworder,lcTable,.T.,'Complete',lcdocno,Datetime())
	Endif
	Return .T.
Endproc
*******
Procedure errhand
= Aerror(aErrorArray)  && Data from most recent error
MESSAGEBOX(aErrorArray(2),16,'Error insert') 
Return aErrorArray(2)
Endproc
***************************
Procedure initEnv
	=SQLEXEC(b,"set dateformat dmy")
	=SQLEXEC(b,"SET LANGUAGE Thai")
	DO SETIdentAlltableOff
	If Dbused('Transfer')
	Else
		Open Database Transfer Shared
	Endif
	If Used('trflog')
	Else
		Use Transfer!trflog In 0
	Endif
Return
********************************
Procedure Delete_Exists_docno && ลบรายการที่มีข้อมูลอยู่แล้ว
	IF !USED('Module')
		MESSAGEBOX('ไม่สามารถติดต่อฐานข้อมูล Module  Masterได้ !',16,'มีปัญหาเรื่องการรับข้อมูล')
		RETURN 
	ENDIF 
	SELECT module 
	GO TOP 
	DO WHILE !EOF()
		DO CASE
		CASE  Moduleid = 0 AND selected = 1  &&Master 
			=SQLEXEC(b,"delete bcstkpacking where itemcode in (select code from trf.dbo.NP_Transfer_Newitem)")
			=SQLEXEC(b,"delete bcitemwarehouse where itemcode in (select code from trf.dbo.NP_Transfer_Newitem)")
		CASE  Moduleid = 1 AND selected = 1  && ซื้อ
			
			=SQLEXEC(b,"delete BCApinvoice where docno in (select docno from tempdb.dbo.dup_BCAPINVOICE )")
			=SQLEXEC(b,"delete BCApinvoiceSUB where docno in (select docno from tempdb.dbo.dup_BCAPINVOICE )")
			=SQLEXEC(b,"delete BCIRSUB where docno in (select docno from tempdb.dbo.dup_BCAPINVOICE )")	
			=SQLEXEC(b,"delete BCInputtax where docno in (select docno from tempdb.dbo.dup_BCAPINVOICE )")	
			=SQLEXEC(b,"delete BCAPWtaxlist where docno in (select docno from tempdb.dbo.dup_BCAPINVOICE )")	
			=SQLEXEC(b,"delete bctrans where docno in (select docno from tempdb.dbo.dup_BCAPINVOICE )")	
			=SQLEXEC(b,"delete bctranssub where docno in (select docno from tempdb.dbo.dup_BCAPINVOICE )")	
			=SQLEXEC(b,"delete BCChqOut where docno in (select docno from tempdb.dbo.dup_BCAPINVOICE )")	
			=SQLEXEC(b,"delete BCPaymoney where docno in (select docno from tempdb.dbo.dup_BCAPINVOICE )")
			=SQLEXEC(b,"delete bcchqout where docno in (select docno from tempdb.dbo.dup_BCAPINVOICE )")	
			=SQLEXEC(b,"delete BCPaymoney where docno in (select docno from tempdb.dbo.dup_BCPayment  )")
			=SQLEXEC(b,"delete BCpayment where docno in (select docno from tempdb.dbo.dup_BCPayment  )")
			=SQLEXEC(b,"delete BCpaymentsub where docno in (select docno from tempdb.dbo.dup_BCPayment  )")
			=SQLEXEC(b,"delete BCChqOut where docno in (select docno from tempdb.dbo.dup_BCPayment  )")
			=SQLEXEC(b,"delete bctrans where docno in (select docno from tempdb.dbo.dup_BCPayment  )")	
			=SQLEXEC(b,"delete bctranssub where docno in (select docno from tempdb.dbo.dup_BCPayment  )")	
			=SQLEXEC(b,"delete BCAPWtaxlist where docno in (select docno from tempdb.dbo.dup_BCPayment  )")	
			=SQLEXEC(b,"delete BCOtherExpense where docno in (select docno from tempdb.dbo.dub_bcotherExpense  )")
			=SQLEXEC(b,"delete bctranssub where docno in (select docno from tempdb.dbo.dub_bcotherExpense  )")	
			=SQLEXEC(b,"delete BCInputtax where docno in (select docno from tempdb.dbo.dub_bcotherExpense  )")	
			=SQLEXEC(b,"delete BCAPWtaxlist where docno in (select docno from tempdb.dbo.dub_bcotherExpense  )")	
			=SQLEXEC(b,"delete BCChqoutpass where docno in (select docno from tempdb.dbo.dub_BCChqoutpass  )")	
			=SQLEXEC(b,"delete BCChqoutpasssub where docno in (select docno from tempdb.dbo.dub_BCChqoutpasssub   )")	
		CASE  Moduleid = 2 AND selected = 1  && ขาย
			=SQLEXEC(b,"delete BCOutputtax where docno in (select docno from tempdb.dbo.dup_bcoutputtax)")
			=SQLEXEC(b,"delete bcarinvoice  where docno in (select docno from tempdb.dbo.dup_bcarinvoice)")
			=SQLEXEC(b,"delete bcarinvoicesub  where docno in (select docno from tempdb.dbo.dup_bcarinvoice)")
			=SQLEXEC(b,"delete bcardeposit where docno in (select docno from tempdb.dbo.dup_bcarDeposit)")
			=SQLEXEC(b,"delete bcinvCreditnote where Creditnoteno in (select docno from tempdb.dbo.dup_bcCreditnote)")
			=SQLEXEC(b,"delete bcCreditnote where docno in (select docno from tempdb.dbo.dup_bcCreditnote)")
			=SQLEXEC(b,"delete bcCreditnotesub where docno in (select docno from tempdb.dbo.dup_bcCreditnote)")
			=SQLEXEC(b,"delete bcArdeposituse where RTRIM(depositno)+'-'+RTRIM(docno) in (select RTRIM(depositno)+'-'+RTRIM(docno) from tempdb.dbo.dup_bcarDeposituse)")
			=SQLEXEC(b,"delete bcDebitnote1  where docno in (select docno from tempdb.dbo.dup_bcDebitnote1)")
			=SQLEXEC(b,"delete bcDebitnotesub1 where docno in (select docno from tempdb.dbo.dup_bcDebitnote1)")
			=SQLEXEC(b,"delete bcInvDebitnote1 where Debitnoteno in (select docno from tempdb.dbo.dup_bcDebitnote1)")	
			=SQLEXEC(b,"delete BCRecmoney where docno in (select docno from tempdb.dbo.dup_bcarinvoice)")	
			=SQLEXEC(b,"delete bcCreditcard where docno in (select docno from tempdb.dbo.dup_bcarinvoice)")	
			=SQLEXEC(b,"delete bcChqin where docno in (select docno from tempdb.dbo.dup_bcarinvoice)")	
			=SQLEXEC(b,"delete bctrans where docno in (select docno from tempdb.dbo.dup_bcarinvoice)")	
			=SQLEXEC(b,"delete bctranssub where docno in (select docno from tempdb.dbo.dup_bcarinvoice)")	
			=SQLEXEC(b,"delete BCReceipt1 where docno in (select docno from tempdb.dbo.dup_BCReceipt1 )")
			=SQLEXEC(b,"delete BCReceiptsub1 where docno in (select docno from tempdb.dbo.dup_BCReceipt1 )")
			=SQLEXEC(b,"delete BCRecmoney where docno in (select docno from tempdb.dbo.dup_BCReceipt1 )")
			=SQLEXEC(b,"delete bcChqin where docno in (select docno from tempdb.dbo.dup_BCReceipt1 )")	
			=SQLEXEC(b,"delete bcCreditcard where docno in (select docno from tempdb.dbo.dup_BCReceipt1 )")	
			=SQLEXEC(b,"delete bctrans where docno in (select docno from tempdb.dbo.dup_BCReceipt1 )")	
			=SQLEXEC(b,"delete bctranssub where docno in (select docno from tempdb.dbo.dup_BCReceipt1 )")	
			=SQLEXEC(b,"delete BCChqindeposit where docno in (select docno from tempdb.dbo.dup_BCChqinDeposit )")
			=SQLEXEC(b,"delete BCChqindeposSub  where docno in (select docno from tempdb.dbo.dup_BCChqinDeposit )")
			=SQLEXEC(b,"delete bctrans where docno in (select docno from tempdb.dbo.dup_BCChqinDeposit )")	
			=SQLEXEC(b,"delete bctranssub where docno in (select docno from tempdb.dbo.dup_BCChqinDeposit )")	

		ENDCASE
		SKIP 
	ENDDO 
Return

******************************
PROCEDURE    SetIdentityOff
PARAMETERS LcTablename
	lccommand=''
	lccommand = "SET IDENTITY_INSERT "+lctablename +" OFF "
	
	result=SQLEXEC(b,"SET IDENTITY_INSERT  "+lctablename +" OFF ")
	If result<0
		MESSAGEBOX(lccommand)
		Do ErrHand
		RETURN .F.
	ELSE 
		RETURN .T.
	ENDIF
ENDPROC 
*******************************
PROCEDURE GetTableList
	lccommand = "select * from  Trf.dbo.np_transfer_tableList"
	result=SQLEXEC(a,lccommand,'TableList')
RETURN 
*****************************
PROCEDURE  SETIdentAlltableOff
	SELECT tablelist
	GO TOP 
	DO WHILE !EOF()
		WAIT WINDOW 'Set Identity off table : '+ALLTRIM(tablelist.name) nowa
		DO  SetidentityOff WITH ALLTRIM(tablelist.tablename)
		SKIP 
	ENDDO 
RETURN 
********************************

PROCEDURE  Transfer 
	IF !USED('Module')
		MESSAGEBOX('ไม่สามารถติดต่อฐานข้อมูล Module  Masterได้ !',16,'มีปัญหาเรื่องการรับข้อมูล')
		RETURN 
	ENDIF 
	SELECT module 
	GO TOP 
	DO WHILE !EOF()
		DO CASE
		CASE  Moduleid = 0 AND selected = 1  &&Master 
			DO 	Trf_BCItem
		CASE  Moduleid = 1 AND selected = 1  && ซื้อ
			DO TRF_BCApinvoice 
			DO TRF_BCApinvoiceSUB 
			DO TRF_BCIRSUB 	 
			DO TRF_BCInputtax  	&&
			DO TRF_BCAPWtaxlist   &&
			*DO TRF_bctrans  		&&
			*DO TRF_bctranssub       &&
			DO TRF_BCChqOut        &&
			DO TRF_BCPaymoney    &&
			DO TRF_BCpayment 
			DO TRF_BCpaymentsub 	
			DO TRF_BCOtherExpense 
			DO TRF_bctranssub 
	
			DO TRF_BCAPWtaxlist 
			DO TRF_BCapdeposit
			DO TRF_BCApdepositUse
			DO Trf_BCAPdepositSpecial
		CASE  Moduleid = 2 AND selected = 1  && ขาย
			DO TRF_BCOutputtax 
			DO TRF_bcarinvoice  
			DO TRF_bcarinvoicesub  
			DO TRF_bcardeposit 
			DO trf_bcinvCreditnote 
			DO trf_bcCreditnote 
			DO TRF_bcCreditnotesub 
			DO TRF_bcArdeposituse 
			DO TRF_bcDebitnote1  
			DO TRF_bcDebitnotesub1
			DO TRF_bcInvDebitnote1 
			DO TRF_BCRecmoney 
			DO TRF_bcCreditcard 
			DO TRF_bcChqin 
			DO TRF_bctrans 
			DO TRF_bctranssub 
			DO TRF_BCReceipt1 
			DO TRF_BCReceiptsub1 
			DO TRF_BCRecmoney 
			DO TRF_bcChqin 
			DO TRF_bcCreditcard 
			DO TRF_BCChqindeposit 
			DO TRF_BCChqindeposSub  		

		ENDCASE 
	ENDDO 
RETURN 
*******************************888
Procedure Insertdata1
	Wait Window Alias()+'.'+Alltrim(Str(Recno())) Nowa
	nfield=Fcount()
	lccommand=" insert "+Alltrim(Alias())+" ("
	lFiled =""
	For initEnv = 1 To nfield
		IF UPPER(ALLTRIM(Fields(initEnv))) != 'ROWORDER'
		*MESSAGEBOX(UPPER(ALLTRIM(Fields(initEnv))) 	)
		lFiled=lFiled+Iif(Len(Alltrim(lFiled))=0,'',',')+Fields(initEnv)
		ENDIF 
	Endfor
	lccommand=lccommand+lFiled+" )"
	*MESSAGEBOX(lccommand)
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
			Case Type(Fields(initEnv)) = 'N'Or Type(Fields(initEnv)) = 'Y'
				lcx="Nextchar=ALLTRIM(STR("+Fields(initEnv)+",15,2))"
				Cflage=0
				Tflage=0
				NFlage=1
			Case Type(Fields(initEnv)) = 'C' Or Type(Fields(initEnv)) = 'M'
				lcx="Nextchar=ALLTRIM("+Fields(initEnv)+")"
				Cflage=1
				Tflage=0
				NFlage=0
			Case Type(Fields(initEnv)) = 'T'
				lcx="Nextchar=DTOC(TTOD("+Fields(initEnv)+"))"
				Cflage=0
				Tflage=1
				NFlage=0

			Endcase
			&lcx
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
	*MESSAGEBOX(lccommand)
	*cancel
	result=SQLEXEC(b,lccommand)
	IF result <0 
		DO Errhand
	ENDIF 	
	lcGenStr="lnRoworder="+Alias()+".roworder"
	&lcGenStr
	lcTable=Alias()

	Do Case
		Case  Alltrim(Upper(Alias()))=='BCINVDEBITNOTE1'
			lcdocno=Alltrim(debitnoteno)+'-'+Alltrim(invoiceno)
		Case Alltrim(Upper(Alias()))== 'BCINVCREDITNOTE'
			lcdocno=Alltrim(creditnoteno)+'-'+Alltrim(invoiceno)
		Case 	Alltrim(Upper(Alias()))=='BCAR' Or  Alltrim(Upper(Alias()))=='BCAP'  Or  Alltrim(Upper(Alias()))=='BCITEM'
			lcdocno = Alltrim(Code)
		Case Alltrim(Upper(Alias()))=='BCSTKPACKING'
			lcdocno=Alltrim(itemcode)+'-'+Alltrim(unitcode)
		Case  Alltrim(Upper(Alias()))=='BCSTKITEMWAREHOUSE'
			lcdocno=Alltrim(itemcode)+'-'+Alltrim(whcode)
		Case  Alltrim(Upper(Alias()))=='BCITEMWAREHOUSE'
			lcdocno=Alltrim(itemcode)+'-'+Alltrim(whcode)
		Otherwise
			lcdocno=docno
	Endcase

	If result<0   && Error Sql
		lcError =  errhand ()
		Insert Into trflog (trfno,roworder,Table,Complete,Errdesc,docno,trfdate) Values (mtrfno,lnRoworder,lcTable,.F.,lcError,lcdocno,Datetime())
		Return .F.
	Else
		Insert Into trflog (trfno,roworder,Table,Complete,Errdesc,docno,trfdate) Values (mtrfno,lnRoworder,lcTable,.T.,'Complete',lcdocno,Datetime())
	Endif
	Return .T.
Endproc
