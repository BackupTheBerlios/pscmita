Public Class CMitaSQL
	Private mvarSystemTablesDrop As String
	Private mvarVersionTablesDrop As String
	Private mvarCombiTablesDrop As String
	Private mvarstandardTablesDrop As String
	Private mvarpoolTablesDrop As String

	Private mvarSystemTables As String
	Private mvarVersionTables As String
	Private mvarCombiTables As String
	Private mvarstandardTables(14) As String
	Private mvarpoolTables(0) As String

	Private mvarstandardIndex(6) As String

	Private mvarProcedures(1) As String

	Public ReadOnly Property procedures() As String()
		Get
			Return mvarProcedures
		End Get
	End Property
	Public ReadOnly Property systemTables() As String
		Get
			Return mvarSystemTables
		End Get
	End Property
	Public ReadOnly Property versionTables() As String
		Get
			Return mvarVersionTables
		End Get
	End Property
	Public ReadOnly Property combiTables() As String
		Get
			Return mvarCombiTables
		End Get
	End Property
	Public ReadOnly Property standardTables() As String()
		Get
			Return mvarstandardTables
		End Get
	End Property
	Public ReadOnly Property standardIndex() As String()
		Get
			Return mvarstandardIndex
		End Get
	End Property
	Public ReadOnly Property poolTables() As String()
		Get
			Return mvarpoolTables
		End Get
	End Property
	Public ReadOnly Property standardTablesDrop() As String
		Get
			Return mvarstandardTablesDrop
		End Get
	End Property
	Public ReadOnly Property poolTablesDrop() As String
		Get
			Return mvarpoolTablesDrop
		End Get
	End Property
	Public ReadOnly Property systemTablesDrop() As String
		Get
			Return mvarSystemTablesDrop
		End Get
	End Property
	Public ReadOnly Property versionTablesDrop() As String
		Get
			Return mvarVersionTablesDrop
		End Get
	End Property
	Public ReadOnly Property combiTablesDrop() As String
		Get
			Return mvarCombiTablesDrop
		End Get
	End Property
	Public ReadOnly Property standardIndexLength() As Integer
		Get
			Return mvarstandardIndex.Length
		End Get
	End Property

	Public ReadOnly Property poolTablesLength() As Integer
		Get
			Return mvarpoolTables.Length
		End Get
	End Property
	Public ReadOnly Property standardTablesLength() As Integer
		Get
			Return mvarstandardTables.Length
		End Get
	End Property
	Public ReadOnly Property proceduresLength() As Integer
		Get
			Return mvarProcedures.Length
		End Get
	End Property
	Public Sub New()
		mvarSystemTables = "CREATE TABLE #SCHEMA#.pscsapsystems" _
		& vbCrLf & "(	sapsystemid	number				NOT NULL," _
		& vbCrLf & "	sapname		varchar(10)			NOT NULL," _
		& vbCrLf & "	version		number				NOT NULL," _
		& vbCrLf & "	sapversionid	number				NOT NULL," _
		& vbCrLf & "	environment	varchar(10)			NOT NULL," _
		& vbCrLf & "	activ		varchar(1)	DEFAULT 'Y'	NOT NULL," _
		& vbCrLf & "	sapgateway	varchar(20)			NOT NULL," _
		& vbCrLf & "	sapservice	varchar(20)			NOT NULL," _
		& vbCrLf & "	sapid		varchar(20)			NOT NULL," _
		& vbCrLf & "	sapidclt	varchar(20)			NOT NULL," _
		& vbCrLf & "	sapuser		varchar(10)			NOT NULL," _
		& vbCrLf & "	sapowner	varchar(10)			NOT NULL," _
		& vbCrLf & "	sapsystem	varchar(10)			NOT NULL," _
		& vbCrLf & "	sapserver	varchar(10)			NOT NULL," _
		& vbCrLf & "	sapclient	varchar(10)			NOT NULL," _
		& vbCrLf & "	databas		varchar(20)			NOT NULL," _
		& vbCrLf & "	databastype	varchar(20)			NOT NULL," _
		& vbCrLf & "	databasuser	varchar(20)			NOT NULL," _
		& vbCrLf & "	databaspwd	varchar(20)			NOT NULL," _
		& vbCrLf & "	createtime	date		DEFAULT sysdate	NOT NULL," _
		& vbCrLf & "	createuser	varchar(20)			NULL," _
		& vbCrLf & "	createid	number			NULL," _
		& vbCrLf & "	createhost	varchar(20)			NULL," _
		& vbCrLf & "	deleted	varchar(1)	DEFAULT 'N'	NOT NULL" _
		& vbCrLf & ")" _
		& vbCrLf & "TABLESPACE #SPACE#; " _
		& vbCrLf & "COMMIT;"

		mvarVersionTables = "CREATE TABLE #SCHEMA#.pscsapversions" _
		& vbCrLf & "(	sapversionid	number				NOT NULL," _
		& vbCrLf & "	sapversionname	varchar(10)		NOT NULL," _
		& vbCrLf & "	sapsubversion	varchar(10)		NULL," _
		& vbCrLf & "	version		number				NOT NULL," _
		& vbCrLf & "	activ		varchar(1)	DEFAULT 'Y'	NOT NULL," _
		& vbCrLf & "	createtime	date		DEFAULT sysdate	NOT NULL," _
		& vbCrLf & "	createuser	varchar(20)			NULL," _
		& vbCrLf & "	createid	number			NULL," _
		& vbCrLf & "	createhost	varchar(20)			NULL," _
		& vbCrLf & "	replacestruct	varchar(255)			NULL," _
		& vbCrLf & "	replacefunc	varchar(255)			null" _
		& vbCrLf & ")" _
		& vbCrLf & "TABLESPACE #SPACE#; " _
		& vbCrLf & "COMMIT;"

		mvarCombiTables = "CREATE TABLE #SCHEMA#.psccustcombis" _
	 & vbCrLf & "(" _
	 & vbCrLf & "	combiname	varchar(20)			NOT NULL," _
	 & vbCrLf & "	version		number				NOT NULL," _
	 & vbCrLf & "	activ		varchar(1)	DEFAULT 'Y'	NOT NULL," _
	 & vbCrLf & "	items		varchar(100)			NOT NULL," _
	 & vbCrLf & "	createtime	date		DEFAULT sysdate	NOT NULL," _
	 & vbCrLf & "	createuser	varchar(20)			NULL," _
	 & vbCrLf & "	createid	number			NULL," _
	 & vbCrLf & "	createhost	varchar(20)			NULL" _
	 & vbCrLf & ")" _
	 & vbCrLf & "TABLESPACE #SPACE#; " _
	 & vbCrLf & "COMMIT;"

		mvarpoolTables(0) = "CREATE TABLE #SCHEMA#.psc§sappool (" _
		& vbCrLf & "	avm		varchar(10)		NOT NULL, " _
		& vbCrLf & "	adno		number			NOT NULL, " _
		& vbCrLf & "	combono		number	DEFAULT 0	NOT NULL, " _
		& vbCrLf & "	dsnfile		varchar(255)," _
		& vbCrLf & "	txtfile		varchar(255)," _
		& vbCrLf & "	ordervno		number			NOT NULL," _
		& vbCrLf & "	motivno		number, " _
		& vbCrLf & "	posno		number," _
		& vbCrLf & "	xloc		number," _
		& vbCrLf & "	yloc		number, " _
		& vbCrLf & "	clientno	varchar(3)," _
		& vbCrLf & "	boxno		varchar(10), " _
		& vbCrLf & "	adtype		varchar(20)," _
		& vbCrLf & "	sortword	varchar(40), " _
		& vbCrLf & "	ysize		number, " _
		& vbCrLf & "	xsize		number) " _
		& vbCrLf & "TABLESPACE #SPACE#;" _
		& vbCrLf & "COMMIT;"

		mvarstandardTables(0) = "CREATE TABLE #SCHEMA#.psc§online" _
		& vbCrLf & "(	sapsystemid	number				NOT NULL," _
		& vbCrLf & "	logintime	date		DEFAULT sysdate	NOT NULL," _
		& vbCrLf & "	logouttime	date				NULL," _
		& vbCrLf & "	loginhost	varchar(20)			NOT NULL," _
		& vbCrLf & "	loginid		number				NOT NULL," _
		& vbCrLf & "	loginapp	varchar(20)			NOT NULL," _
		& vbCrLf & "	orderno		varchar(10)			NULL," _
		& vbCrLf & "	lastorder	varchar(10)			NULL," _
		& vbCrLf & "	alive		date				NULL," _
		& vbCrLf & "	alarm		varchar(255)			NULL," _
		& vbCrLf & "	processid	number				NULL," _
		& vbCrLf & "	caption		varchar(255)			NULL," _
		& vbCrLf & "	command		varchar(255)			NULL," _
		& vbCrLf & "	hostindex	number				NOT NULL" _
		& vbCrLf & ")" _
		& vbCrLf & "TABLESPACE #SPACE#; " _
		& vbCrLf & "COMMIT;"

		mvarstandardTables(1) = "CREATE TABLE #SCHEMA#.psc§ordercontrol" _
		& vbCrLf & "(" _
		& vbCrLf & "	pscid		number				NOT NULL," _
		& vbCrLf & "	sapversionid	number			NOT NULL," _
		& vbCrLf & "	sapsystemid	number			NOT NULL," _
		& vbCrLf & "	avm		varchar(10)			NOT NULL," _
		& vbCrLf & "	version		number				NOT NULL," _
		& vbCrLf & "	status		varchar(1)	DEFAULT 'N'	NOT NULL," _
		& vbCrLf & "	activ		varchar(1)	DEFAULT 'Y'	NOT NULL," _
		& vbCrLf & "	length		number				NOT NULL," _
		& vbCrLf & "	createtime	date		DEFAULT sysdate	NOT NULL," _
		& vbCrLf & "	createid	number			NULL," _
		& vbCrLf & "	createhost	varchar(20)			NULL," _
		& vbCrLf & "	client	varchar(20)			NULL," _
		& vbCrLf & "	opentime	date				NULL," _
		& vbCrLf & "	openid		varchar(20)			NULL," _
		& vbCrLf & "	openhost	varchar(20)			NULL," _
		& vbCrLf & "	closetime	date				NULL," _
		& vbCrLf & "	hostflag	varchar(20)			NULL," _
		& vbCrLf & "	idflag		varchar(20)			NULL," _
		& vbCrLf & "	processed	number		DEFAULT 0	NOT NULL," _
		& vbCrLf & "	comments	varchar(255)			NULL" _
		& vbCrLf & ")" _
		& vbCrLf & "TABLESPACE #SPACE#; " _
		& vbCrLf & "COMMIT;"

		mvarstandardTables(2) = "CREATE TABLE #SCHEMA#.psc§orderdata" _
		& vbCrLf & "(" _
		& vbCrLf & "	pscid		number				NOT NULL," _
		& vbCrLf & "	structures	blob				NULL," _
		& vbCrLf & "	hashcode	blob				NULL" _
		& vbCrLf & ")" _
		& vbCrLf & "TABLESPACE #SPACE#; " _
		& vbCrLf & "COMMIT;"

		mvarstandardTables(3) = "CREATE TABLE #SCHEMA#.psc§errorcontrol" _
		& vbCrLf & "(" _
		& vbCrLf & "	pscid		number				NOT NULL," _
		& vbCrLf & "	sapversionid	number			NOT NULL," _
		& vbCrLf & "	sapsystemid	number			NOT NULL," _
		& vbCrLf & "	avm		varchar(10)			NOT NULL," _
		& vbCrLf & "	version		number				NOT NULL," _
		& vbCrLf & "	status		varchar(1)	DEFAULT 'N'	NOT NULL," _
		& vbCrLf & "	activ		varchar(1)	DEFAULT 'Y'	NOT NULL," _
		& vbCrLf & "	length		number				NOT NULL," _
		& vbCrLf & "	createtime	date		DEFAULT sysdate	NOT NULL," _
		& vbCrLf & "	createid	number			NULL," _
		& vbCrLf & "	createhost	varchar(20)			NULL," _
		& vbCrLf & "	opentime	date				NULL," _
		& vbCrLf & "	openid		varchar(20)			NULL," _
		& vbCrLf & "	openhost	varchar(20)			NULL," _
		& vbCrLf & "	closetime	date				NULL," _
		& vbCrLf & "	hostflag	varchar(20)			NULL," _
		& vbCrLf & "	idflag		varchar(20)			NULL," _
		& vbCrLf & "	processed	number		DEFAULT 0	NOT NULL," _
		& vbCrLf & "	comments	varchar(255)			NULL" _
		& vbCrLf & ")" _
		& vbCrLf & "TABLESPACE #SPACE#; " _
		& vbCrLf & "COMMIT;"

		mvarstandardTables(4) = "CREATE TABLE #SCHEMA#.psc§errordata" _
		& vbCrLf & "(" _
		& vbCrLf & "	pscid		number				NOT NULL," _
		& vbCrLf & "	structures	blob				NULL," _
		& vbCrLf & "	hashcode	blob				NULL" _
		& vbCrLf & ")" _
		& vbCrLf & "TABLESPACE #SPACE#; " _
		& vbCrLf & "COMMIT;"

		mvarstandardTables(5) = "CREATE TABLE #SCHEMA#.psc§structures" _
		& vbCrLf & "(" _
		& vbCrLf & "	sapstruct	varchar(20)			NOT NULL," _
		& vbCrLf & "	sapversionid	number				NOT NULL," _
		& vbCrLf & "	rfctype		number				NOT NULL," _
		& vbCrLf & "	version		number				NOT NULL," _
		& vbCrLf & "	activ		varchar(1)	DEFAULT 'Y'	NOT NULL," _
		& vbCrLf & "	rfcfunction	varchar(30)			NOT NULL," _
		& vbCrLf & "	datastruct	varchar(30)			NULL," _
		& vbCrLf & "	datatype	char		DEFAULT 'T'	NOT NULL," _
		& vbCrLf & "	slevel		number				NOT NULL," _
		& vbCrLf & "	recno		number				NOT NULL," _
		& vbCrLf & "	length		number				NOT NULL," _
		& vbCrLf & "	createtime	date		DEFAULT sysdate	NOT NULL," _
		& vbCrLf & "	createuser	varchar(20)			NULL," _
		& vbCrLf & "	createid	number			NULL," _
		& vbCrLf & "	createhost	varchar(20)			NULL" _
		& vbCrLf & ")" _
		& vbCrLf & "TABLESPACE #SPACE#; " _
		& vbCrLf & "COMMIT;"

		mvarstandardTables(6) = "CREATE TABLE #SCHEMA#.psc§structfields" _
		& vbCrLf & "(" _
		& vbCrLf & "	sapstruct	varchar(20)			NOT NULL," _
		& vbCrLf & "	sapversionid	number				NOT NULL," _
		& vbCrLf & "	rfctype		number				NOT NULL," _
		& vbCrLf & "	field		varchar(30)			NOT NULL," _
		& vbCrLf & "	version		number				NOT NULL," _
		& vbCrLf & "	activ		varchar(1)	DEFAULT 'Y'	NOT NULL," _
		& vbCrLf & "	first		number				NOT NULL," _
		& vbCrLf & "	recno		number				NOT NULL," _
		& vbCrLf & "	length		number				NOT NULL," _
		& vbCrLf & "	createtime	date		DEFAULT sysdate	NOT NULL," _
		& vbCrLf & "	createuser	varchar(20)			NULL," _
		& vbCrLf & "	createid	number			NULL," _
		& vbCrLf & "	createhost	varchar(20)			NULL" _
		& vbCrLf & ")" _
		& vbCrLf & "TABLESPACE #SPACE#; " _
		& vbCrLf & "COMMIT;"

		mvarstandardTables(7) = "CREATE TABLE #SCHEMA#.psc§custquery" _
		& vbCrLf & "(" _
		& vbCrLf & "	pscname		varchar(50)			NOT NULL," _
		& vbCrLf & "	sapsystemid	number			NOT NULL," _
		& vbCrLf & "	rfctype		number			NOT NULL," _
		& vbCrLf & "	version		number				NOT NULL," _
		& vbCrLf & "	activ		varchar(1)	DEFAULT 'Y'	NOT NULL," _
		& vbCrLf & "	text		varchar(2000)			NOT NULL," _
		& vbCrLf & "	createtime	date		DEFAULT sysdate	NOT NULL," _
		& vbCrLf & "	createuser	varchar(20)			NULL," _
		& vbCrLf & "	createid	number			NULL," _
		& vbCrLf & "	createhost	varchar(20)			NULL," _
		& vbCrLf & "	comments	varchar(255)			NULL" _
		& vbCrLf & ")" _
		& vbCrLf & "TABLESPACE #SPACE#; " _
		& vbCrLf & "COMMIT;"

		mvarstandardTables(8) = "CREATE TABLE #SCHEMA#.psc§eventlog" _
		& vbCrLf & "(" _
		& vbCrLf & "	pscid		number				NOT NULL," _
		& vbCrLf & "	sapsystemid	number			NOT NULL," _
		& vbCrLf & "	rfctype		number			NOT NULL," _
		& vbCrLf & "	avm		varchar(10)			NULL," _
		& vbCrLf & "	avmversion	number			NULL," _
		& vbCrLf & "	motiv		number				NULL," _
		& vbCrLf & "	adno		number				NULL," _
		& vbCrLf & "	adver		number				NULL," _
		& vbCrLf & "	paper		varchar(10)			NULL," _
		& vbCrLf & "	priority	number		DEFAULT 0	NOT NULL," _
		& vbCrLf & "	typ		varchar(2)	DEFAULT 'L'	NOT NULL," _
		& vbCrLf & "	status		number		DEFAULT 0	NOT NULL," _
		& vbCrLf & "	text		varchar(2000)			NOT NULL," _
		& vbCrLf & "	createtime	date		DEFAULT sysdate	NOT NULL," _
		& vbCrLf & "	createuser	varchar(20)			NULL," _
		& vbCrLf & "	createid	number			NULL," _
		& vbCrLf & "	createhost	varchar(20)			NULL," _
		& vbCrLf & "	comments	varchar(1000)			NULL," _
		& vbCrLf & "	cnt		number				NULL," _
		& vbCrLf & "	event		varchar(32)			NULL" _
		& vbCrLf & ")" _
		& vbCrLf & "TABLESPACE #SPACE#; " _
		& vbCrLf & "COMMIT;"

		mvarstandardTables(9) = "CREATE TABLE #SCHEMA#.psc§eventcontrol" _
		& vbCrLf & "(" _
		& vbCrLf & "	event		varchar(32)			NOT NULL," _
		& vbCrLf & "	sapsystemid	number			NOT NULL," _
		& vbCrLf & "	rfctype		number			NOT NULL," _
		& vbCrLf & "	runtype		varchar(10)	DEFAULT 'PROD'	NOT NULL," _
		& vbCrLf & "	version		number				NOT NULL," _
		& vbCrLf & "	masterversion	number				NOT NULL," _
		& vbCrLf & "	recno		number		DEFAULT -1	NOT NULL," _
		& vbCrLf & "	action		varchar(16)			NULL," _
		& vbCrLf & "	activ		varchar(1)	DEFAULT 'Y'	NOT NULL," _
		& vbCrLf & "	parameter	varchar(255)			NULL," _
		& vbCrLf & "	createtime	date		DEFAULT sysdate	NOT NULL," _
		& vbCrLf & "	createuser	varchar(20)			NULL," _
		& vbCrLf & "	createid	number			NULL," _
		& vbCrLf & "	createhost	varchar(20)			NULL," _
		& vbCrLf & "	comments	varchar(255)			NULL" _
		& vbCrLf & ")" _
		& vbCrLf & "TABLESPACE #SPACE#; " _
		& vbCrLf & "COMMIT;"

		mvarstandardTables(10) = "CREATE TABLE #SCHEMA#.psc§reportcontrol" _
		& vbCrLf & "(" _
		& vbCrLf & "	controlname	varchar(32)			NOT NULL," _
		& vbCrLf & "	controltyp      varchar(1)			NOT NULL," _
		& vbCrLf & "	version		number				NOT NULL," _
		& vbCrLf & "	length		number				NOT NULL," _
		& vbCrLf & "	activ		varchar(1)	DEFAULT 'Y'	NOT NULL," _
		& vbCrLf & "	createtime	date		DEFAULT sysdate	NOT NULL," _
		& vbCrLf & "	createuser	varchar(20)			NULL," _
		& vbCrLf & "	createid	number			NULL," _
		& vbCrLf & "	createhost	varchar(20)			NULL," _
		& vbCrLf & "	comments	varchar(255)			NULL," _
		& vbCrLf & "	controldata     blob				NULL" _
		& vbCrLf & ")" _
		& vbCrLf & "TABLESPACE #SPACE#; " _
		& vbCrLf & "COMMIT;"

		mvarstandardTables(11) = "CREATE TABLE #SCHEMA#.psc§custtables" _
		& vbCrLf & "(" _
		& vbCrLf & "	tablename	varchar(20)			NOT NULL," _
		& vbCrLf & "	sapsystemid	number			NOT NULL," _
		& vbCrLf & "	rfctype		number			NOT NULL," _
		& vbCrLf & "	version		number				NOT NULL," _
		& vbCrLf & "	activ		varchar(1)	DEFAULT 'Y'	NOT NULL," _
		& vbCrLf & "	tleft		varchar(50)			NULL," _
		& vbCrLf & "	tright		varchar(50)			NULL," _
		& vbCrLf & "	createtime	date		DEFAULT sysdate	NOT NULL," _
		& vbCrLf & "	createuser	varchar(20)			NULL," _
		& vbCrLf & "	createid	number			NULL," _
		& vbCrLf & "	createhost	varchar(20)			NULL" _
		& vbCrLf & ")" _
		& vbCrLf & "TABLESPACE #SPACE#; " _
		& vbCrLf & "COMMIT;"

		mvarstandardTables(12) = "CREATE TABLE #SCHEMA#.psc§custfields" _
		& vbCrLf & "(" _
		& vbCrLf & "	field		varchar(20)			NOT NULL," _
		& vbCrLf & "	slevel		number		DEFAULT -1	NOT NULL," _
		& vbCrLf & "	sapsystemid	number			NOT NULL," _
		& vbCrLf & "	rfctype		number			NOT NULL," _
		& vbCrLf & "	version		number				NOT NULL," _
		& vbCrLf & "	activ		varchar(1)	DEFAULT 'Y'	NOT NULL," _
		& vbCrLf & "	sapstruct	varchar(20)			NOT NULL," _
		& vbCrLf & "	sapfield	varchar(20)			NOT NULL," _
		& vbCrLf & "	first		number				NOT NULL," _
		& vbCrLf & "	recno		number				NOT NULL," _
		& vbCrLf & "	length		number				NOT NULL," _
		& vbCrLf & "	ftype		varchar(20)			NOT NULL," _
		& vbCrLf & "	createtime	date		DEFAULT sysdate	NOT NULL," _
		& vbCrLf & "	createuser	varchar(20)			NULL," _
		& vbCrLf & "	createid	number			NULL," _
		& vbCrLf & "	createhost	varchar(20)			NULL" _
		& vbCrLf & ")" _
		& vbCrLf & "TABLESPACE #SPACE#; " _
		& vbCrLf & "COMMIT;"

		mvarstandardTables(13) = "CREATE SEQUENCE #SCHEMA#.orderid§ INCREMENT BY 1 START WITH 1 " _
		& vbCrLf & "    MAXVALUE 1.0E28 MINVALUE 1 NOCYCLE " _
		& vbCrLf & "    CACHE 20 NOORDER;" _
		& vbCrLf & "COMMIT;"

		mvarstandardTables(14) = "CREATE SEQUENCE #SCHEMA#.errorid§ INCREMENT BY 1 START WITH 1 " _
		& vbCrLf & "    MAXVALUE 1.0E28 MINVALUE 1 NOCYCLE " _
		& vbCrLf & "    CACHE 20 NOORDER;" _
		& vbCrLf & "COMMIT;"

		mvarstandardIndex(0) = "CREATE INDEX #SCHEMA#.psc§orderidx" _
		& vbCrLf & "    ON #SCHEMA#.psc§ordercontrol  ('aktiv', 'openid', 'status') " _
		& vbCrLf & "TABLESPACE #SPACE# PCTFREE 10 INITRANS 2 MAXTRANS 255 " _
		& vbCrLf & "    STORAGE ( INITIAL 64K NEXT 0K MINEXTENTS 1 MAXEXTENTS " _
		& vbCrLf & "    2147483645 PCTINCREASE 0 FREELISTS 1 FREELIST GROUPS 1) " _
		& vbCrLf & "    LOGGING;" _
		& vbCrLf & "COMMIT;"

		mvarstandardIndex(1) = "CREATE INDEX #SCHEMA#.psc§avmidx" _
		& vbCrLf & "    ON #SCHEMA#.psc§ordercontrol  ('avm', 'version')" _
		& vbCrLf & "TABLESPACE #SPACE# PCTFREE 10 INITRANS 2 MAXTRANS 255 " _
		& vbCrLf & "    STORAGE ( INITIAL 64K NEXT 0K MINEXTENTS 1 MAXEXTENTS " _
		& vbCrLf & "    2147483645 PCTINCREASE 0 FREELISTS 1 FREELIST GROUPS 1) " _
		& vbCrLf & "    LOGGING;" _
		& vbCrLf & "COMMIT;"

		mvarstandardIndex(2) = "CREATE INDEX #SCHEMA#.psc§controlididx" _
		& vbCrLf & "    ON #SCHEMA#.psc§ordercontrol  ('pscid')" _
		& vbCrLf & "TABLESPACE #SPACE# PCTFREE 10 INITRANS 2 MAXTRANS 255 " _
		& vbCrLf & "    STORAGE ( INITIAL 64K NEXT 0K MINEXTENTS 1 MAXEXTENTS " _
		& vbCrLf & "    2147483645 PCTINCREASE 0 FREELISTS 1 FREELIST GROUPS 1) " _
		& vbCrLf & "    LOGGING;" _
		& vbCrLf & "COMMIT;"

		mvarstandardIndex(3) = "CREATE INDEX #SCHEMA#.psc§openididx" _
		& vbCrLf & "    ON #SCHEMA#.psc§ordercontrol  ('openid')" _
		& vbCrLf & "TABLESPACE #SPACE# PCTFREE 10 INITRANS 2 MAXTRANS 255 " _
		& vbCrLf & "    STORAGE ( INITIAL 64K NEXT 0K MINEXTENTS 1 MAXEXTENTS " _
		& vbCrLf & "    2147483645 PCTINCREASE 0 FREELISTS 1 FREELIST GROUPS 1) " _
		& vbCrLf & "    LOGGING;" _
		& vbCrLf & "COMMIT;"

		mvarstandardIndex(4) = "CREATE INDEX #SCHEMA#.psc§dataididx" _
		& vbCrLf & "    ON #SCHEMA#.psc§orderdata  ('pscid')" _
		& vbCrLf & "TABLESPACE #SPACE# PCTFREE 10 INITRANS 2 MAXTRANS 255 " _
		& vbCrLf & "    STORAGE ( INITIAL 64K NEXT 0K MINEXTENTS 1 MAXEXTENTS " _
		& vbCrLf & "    2147483645 PCTINCREASE 0 FREELISTS 1 FREELIST GROUPS 1) " _
		& vbCrLf & "    LOGGING;" _
		& vbCrLf & "COMMIT;"

		mvarstandardIndex(5) = "CREATE INDEX #SCHEMA#.psc§tablesleftidx " _
		& vbCrLf & "    ON #SCHEMA#.psc§custtables  ('tablename', 'left') " _
		& vbCrLf & "TABLESPACE #SPACE# PCTFREE 10 INITRANS 2 MAXTRANS 255 " _
		& vbCrLf & "    STORAGE ( INITIAL 64K NEXT 0K MINEXTENTS 1 MAXEXTENTS " _
		& vbCrLf & "    2147483645 PCTINCREASE 0 FREELISTS 1 FREELIST GROUPS 1) " _
		& vbCrLf & "    LOGGING;" _
		& vbCrLf & "COMMIT;"

		mvarstandardIndex(6) = "CREATE INDEX #SCHEMA#.psc§poolavmidx " _
		& vbCrLf & "    ON #SCHEMA#.psc§sappool  ('avm', 'clientno') " _
		& vbCrLf & "TABLESPACE #SPACE# PCTFREE 10 INITRANS 2 MAXTRANS 255 " _
		& vbCrLf & "    STORAGE ( INITIAL 64K NEXT 0K MINEXTENTS 1 MAXEXTENTS " _
		& vbCrLf & "    2147483645 PCTINCREASE 0 FREELISTS 1 FREELIST GROUPS 1) " _
		& vbCrLf & "    LOGGING;" _
		& vbCrLf & "COMMIT;"

		mvarSystemTablesDrop = "DROP TABLE #SCHEMA#.pscsapsystems; COMMIT;"
		mvarVersionTablesDrop = "DROP TABLE #SCHEMA#.pscsapversions; COMMIT"
		mvarCombiTablesDrop = "DROP TABLE #SCHEMA#.psccustcombis; COMMIT;"

		mvarstandardTablesDrop = "DROP SEQUENCE #SCHEMA#.errorid§;" _
		& vbCrLf & "DROP SEQUENCE #SCHEMA#.orderid§;" _
		& vbCrLf & "DROP TABLE #SCHEMA#.psc§custfields;" _
		& vbCrLf & "DROP TABLE #SCHEMA#.psc§custquery;" _
		& vbCrLf & "DROP TABLE #SCHEMA#.psc§custtables;" _
		& vbCrLf & "DROP TABLE #SCHEMA#.psc§errorcontrol;" _
		& vbCrLf & "DROP TABLE #SCHEMA#.psc§errordata;" _
		& vbCrLf & "DROP TABLE #SCHEMA#.psc§eventcontrol;" _
		& vbCrLf & "DROP TABLE #SCHEMA#.psc§eventlog;" _
		& vbCrLf & "DROP TABLE #SCHEMA#.psc§online;" _
		& vbCrLf & "DROP TABLE #SCHEMA#.psc§ordercontrol;" _
		& vbCrLf & "DROP TABLE #SCHEMA#.psc§orderdata;" _
		& vbCrLf & "DROP TABLE #SCHEMA#.psc§reportcontrol;" _
		& vbCrLf & "DROP TABLE #SCHEMA#.psc§structfields;" _
		& vbCrLf & "DROP TABLE #SCHEMA#.psc§structures;" _
		& vbCrLf & "COMMIT;"

		mvarpoolTablesDrop = "DROP TABLE #SCHEMA#.psc§sappool; COMMIT;"

		mvarProcedures(0) = "CREATE OR REPLACE PROCEDURE advdb.PSC§_NEXTORDER_SP  (" _
		& vbCrLf & " f_myid IN NUMBER," _
		& vbCrLf & " f_sapsystem IN NUMBER," _
		& vbCrLf & " f_sapversion IN NUMBER," _
		& vbCrLf & " f_pscid OUT NUMBER," _
		& vbCrLf & " f_processed OUT NUMBER" _
		& vbCrLf & ")" _
		& vbCrLf & "AS" _
		& vbCrLf & "  t_count NUMBER;" _
		& vbCrLf & "  t_id NUMBER;" _
		& vbCrLf & "  f_avm VARCHAR(10);" _
		& vbCrLf & "  f_version NUMBER;" _
		& vbCrLf & "  CURSOR my_cursor IS" _
		& vbCrLf & "    SELECT pscid, avm, version, processed FROM psc§ordercontrol " _
		& vbCrLf & "       WHERE status = 'N'" _
		& vbCrLf & "        AND (" _
		& vbCrLf & "          (openid IS NULL OR openid = f_myid)" _
		& vbCrLf & "          AND activ = 'Y'" _
		& vbCrLf & "          OR (openid = f_myid AND activ = 'W')" _
		& vbCrLf & "        )" _
		& vbCrLf & "        AND sapversionid = f_sapversion" _
		& vbCrLf & "        AND sapsystemid = f_sapsystem" _
		& vbCrLf & "        ORDER BY createtime DESC" _
		& vbCrLf & "        FOR UPDATE;" _
		& vbCrLf & "        " _
		& vbCrLf & "  BEGIN" _
		& vbCrLf & "    f_pscid := -1;" _
		& vbCrLf & "    f_avm := '0000000000';" _
		& vbCrLf & "    f_version := -1;" _
		& vbCrLf & "     " _
		& vbCrLf & "    OPEN my_cursor;" _
		& vbCrLf & "    LOOP" _
		& vbCrLf & "      FETCH my_cursor into t_id, f_avm, f_version, f_processed;" _
		& vbCrLf & "      EXIT WHEN my_cursor%NOTFOUND OR" _
		& vbCrLf & "        my_cursor%NOTFOUND IS NULL;" _
		& vbCrLf & "      SELECT COUNT(avm) into t_count FROM psc§ordercontrol" _
		& vbCrLf & "         WHERE status = 'N'" _
		& vbCrLf & "          AND openid <> f_myid" _
		& vbCrLf & "          AND activ = 'W'" _
		& vbCrLf & "          AND avm = f_avm" _
		& vbCrLf & "          AND sapversionid = f_sapversion" _
		& vbCrLf & "          AND sapsystemid = f_sapsystem;" _
		& vbCrLf & "          " _
		& vbCrLf & "      IF t_count = 0 THEN" _
		& vbCrLf & "        f_pscid := t_id;     " _
		& vbCrLf & "        UPDATE  psc§ordercontrol SET activ = 'W', openid = f_myid" _
		& vbCrLf & "            WHERE CURRENT OF my_cursor;" _
		& vbCrLf & "        COMMIT;" _
		& vbCrLf & "        CLOSE my_cursor; " _
		& vbCrLf & "         EXIT; " _
		& vbCrLf & "      END IF;" _
		& vbCrLf & "    END LOOP;" _
		& vbCrLf & "  END;"

		mvarProcedures(1) = "CREATE OR REPLACE  PROCEDURE ADVDB.PSC§_SAPPOOL_SP  (" _
		& vbCrLf & "    p_INTAVM IN VARCHAR2," _
		& vbCrLf & "    p_INTMONR IN NUMBER," _
		& vbCrLf & "    p_INTADNO IN NUMBER," _
		& vbCrLf & "    p_INTCOMBONO  IN NUMBER," _
		& vbCrLf & "    p_INTPOSNR IN NUMBER," _
		& vbCrLf & "    p_CLIENTNO IN VARCHAR2," _
		& vbCrLf & "    p_DSNPATH IN VARCHAR2," _
		& vbCrLf & "    p_TXTPATH IN VARCHAR2," _
		& vbCrLf & "    p_ORDERVNO IN NUMBER," _
		& vbCrLf & "    p_INTBOXNO IN VARCHAR2," _
		& vbCrLf & "    p_TOPTEXT IN VARCHAR2," _
		& vbCrLf & "    p_TDEPTH IN NUMBER," _
		& vbCrLf & "    p_TWIDTH IN NUMBER)" _
		& vbCrLf & "AS" _
		& vbCrLf & " p_SAPCOUNT NUMBER; " _
		& vbCrLf & " BEGIN  " _
		& vbCrLf & "      SELECT count(AVM) INTO p_SAPCOUNT FROM psc§sappool " _
		& vbCrLf & "      WHERE AVM= p_INTAVM " _
		& vbCrLf & "      AND COMBONO= p_INTCOMBONO " _
		& vbCrLf & "      AND CLIENTNO= p_CLIENTNO " _
		& vbCrLf & "      AND MOTIVNO= p_INTMONR;" _
		& vbCrLf & "   IF p_SAPCOUNT >0 THEN  " _
		& vbCrLf & "     UPDATE psc§sappool " _
		& vbCrLf & "     SET adno=p_INTADNO,dsnfile=p_DSNPATH,txtfile=p_TXTPATH," _
		& vbCrLf & "      ordervno=p_ORDERVNO,posno=p_INTPOSNR,boxno=p_INTBOXNO," _
		& vbCrLf & "      sortword=p_TOPTEXT,ysize=p_TDEPTH,xsize=p_TWIDTH " _
		& vbCrLf & "      WHERE AVM= p_INTAVM " _
		& vbCrLf & "      AND COMBONO= p_INTCOMBONO " _
		& vbCrLf & "      AND CLIENTNO= p_CLIENTNO " _
		& vbCrLf & "      AND MOTIVNO= p_INTMONR;" _
		& vbCrLf & "   ELSE  " _
		& vbCrLf & "     INSERT INTO psc§sappool " _
		& vbCrLf & "     (avm,adno,combono,dsnfile," _
		& vbCrLf & "     txtfile,ordervno,motivno,posno,clientno," _
		& vbCrLf & "     boxno,sortword,ysize,xsize)" _
		& vbCrLf & "     VALUES (p_INTAVM,p_INTADNO,p_INTCOMBONO,p_DSNPATH," _
		& vbCrLf & "     p_TXTPATH,p_ORDERVNO,p_INTMONR,p_INTPOSNR,p_CLIENTNO," _
		& vbCrLf & "     p_INTBOXNO,p_TOPTEXT,p_TDEPTH,p_TWIDTH);" _
		& vbCrLf & "  END IF;" _
		& vbCrLf & "  COMMIT;" _
		& vbCrLf & " END;"

	End Sub
End Class
