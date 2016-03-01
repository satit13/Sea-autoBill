
CREATE TABLE bcarinvoice (;
	docno c(30),;
	docdate d DEFAULT DATE(),;
	taxno c(30),;
	taxtype int  DEFAULT 0,;
	arcode c(30) ,;
	arname c(100),;
	araddress c(100),;
	creditday int DEFAULT 0,;
	duedate d default DATE(),;
	salecode c(30),;
	taxrate n(4,2),;
	isconfirm int DEFAULT  0,;
	mydescription c(100),;
	billtype int DEFAULT 1,;
	sumofitemamount n(15,2),;
	discountamount n(15,2),;
	afterdiscount n(15,2),;
	befortaxamount n(15,2),;
	taxamount n(15,2),;
	totalamount n(15,2),;
	excepttaxamount n(15,2),;
	netdebtamount n(15,2),;
	homeamount n(15,2),;
	glformat c(10),;
	iscancel int default 0,;
	iscompletesave int default 0,;
	ispostgl int default 0,;
	paybillamount n(15,2),;
	billtemporary n(15,2),;
	ceatorcode c(30),;
	createdatetime d )


CREATE TABLE  bcarinvoicesub (;
	BehindIndex int default 1 ,;
	MyType int DEFAULT 4,;
	DocNo c(30) ,;
	TaxNo c(30) , ;
	TaxType int default 0, ;
	ItemCode c(30),;
	DocDate d default DATE(),;
	ArCode c(30),;
	ItemName c(100),;
	WHCode c(30),;
	ShelfCode c(30),;
	CNQty n(15,2) ,;
	Qty n(15,2), ;
	price n(15,2) ,;
	Amount n(15,2),;
	NetAmount n(15,2),;
	HomeAmount n(15,2),;
	sumofcost n(15,2),;
	unitcode c(30),;
	LineNumber int DEFAULT 0,; 
	IsCancel int default 0,;
	TaxRate n(4,2) DEFAULT 0.00,;
	PackingRate1 n(10,2) default 1.00,;
	packingrate2 n(10,2) default 1.00)

