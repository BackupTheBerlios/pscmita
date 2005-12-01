CREATE TABLE advdb.psccustcombis
(
	combiname	varchar(20)			NOT NULL,
	version		number				NOT NULL,
	activ		varchar(1)	DEFAULT 'Y'	NOT NULL,
	items		varchar(100)			NOT NULL,
	createtime	date		DEFAULT sysdate	NOT NULL,
	createuser	varchar(20)			NULL,
	createid	number			NULL,
	createhost	varchar(20)			NULL
)
TABLESPACE TSADVD0; 
COMMIT;

CREATE TABLE advdb.psconline
(	sapsystemid	number				NOT NULL,
	logintime	date		DEFAULT sysdate	NOT NULL,
	logouttime	date				NULL,
	loginhost	varchar(20)			NOT NULL,
	loginid		number				NOT NULL,
	loginapp	varchar(20)			NOT NULL,
	orderno		varchar(10)			NULL,
	lastorder	varchar(10)			NULL,
	alive		date				NULL,
	alarm		varchar(255)			NULL,
	processid	number				NULL,
	caption		varchar(255)			NULL,
	command		varchar(255)			NULL,
	hostindex	number				NOT NULL
)
TABLESPACE TSADVD0; 
COMMIT;

CREATE TABLE advdb.pscordercontrol
(
	pscid		number				NOT NULL,
	sapversionid	number			NOT NULL,
	sapsystemid	number			NOT NULL,
	avm		varchar(10)			NOT NULL,
	version		number				NOT NULL,
	status		varchar(1)	DEFAULT 'N'	NOT NULL,
	activ		varchar(1)	DEFAULT 'Y'	NOT NULL,
	length		number				NOT NULL,
	createtime	date		DEFAULT sysdate	NOT NULL,
	createid	number			NULL,
	createhost	varchar(20)			NULL,
	createclient	varchar(20)			NULL,
	opentime	date				NULL,
	openid		varchar(20)			NULL,
	openhost	varchar(20)			NULL,
	closetime	date				NULL,
	hostflag	varchar(20)			NULL,
	idflag		varchar(20)			NULL,
	processed	number		DEFAULT 0	NOT NULL,
	comments	varchar(255)			NULL
)
TABLESPACE TSADVD0; 
COMMIT;

CREATE TABLE advdb.pscorderdata
(
	pscid		number				NOT NULL,
	structures	blob				NULL,
	hashcode	blob				NULL
)
TABLESPACE TSADVD0; 
COMMIT;

CREATE TABLE advdb.pscerrorcontrol
(
	pscid		number				NOT NULL,
	sapversionid	number			NOT NULL,
	sapsystemid	number			NOT NULL,
	avm		varchar(10)			NOT NULL,
	version		number				NOT NULL,
	status		varchar(1)	DEFAULT 'N'	NOT NULL,
	activ		varchar(1)	DEFAULT 'Y'	NOT NULL,
	length		number				NOT NULL,
	createtime	date		DEFAULT sysdate	NOT NULL,
	createid	number			NULL,
	createhost	varchar(20)			NULL,
	opentime	date				NULL,
	openid		varchar(20)			NULL,
	openhost	varchar(20)			NULL,
	closetime	date				NULL,
	hostflag	varchar(20)			NULL,
	idflag		varchar(20)			NULL,
	processed	number		DEFAULT 0	NOT NULL,
	comments	varchar(255)			NULL
)
TABLESPACE TSADVD0; 
COMMIT;

CREATE TABLE advdb.pscerrordata
(
	pscid		number				NOT NULL,
	structures	blob				NULL,
	hashcode	blob				NULL
)
TABLESPACE TSADVD0; 
COMMIT;

CREATE TABLE advdb.pscstructures
(
	sapstruct	varchar(20)			NOT NULL,
	sapversionid	number				NOT NULL,
	rfctype		number				NOT NULL,
	version		number				NOT NULL,
	activ		varchar(1)	DEFAULT 'Y'	NOT NULL,
	rfcfunction	varchar(30)			NOT NULL,
	datastruct	varchar(30)			NULL,
	datatype	char		DEFAULT 'T'	NOT NULL,
	slevel		number				NOT NULL,
	recno		number				NOT NULL,
	length		number				NOT NULL,
	createtime	date		DEFAULT sysdate	NOT NULL,
	createuser	varchar(20)			NULL,
	createid	number			NULL,
	createhost	varchar(20)			NULL
)
TABLESPACE TSADVD0; 
COMMIT;

CREATE TABLE advdb.pscstructfields
(
	sapstruct	varchar(20)			NOT NULL,
	sapversionid	number				NOT NULL,
	rfctype		number				NOT NULL,
	field		varchar(30)			NOT NULL,
	version		number				NOT NULL,
	activ		varchar(1)	DEFAULT 'Y'	NOT NULL,
	first		number				NOT NULL,
	recno		number				NOT NULL,
	length		number				NOT NULL,
	createtime	date		DEFAULT sysdate	NOT NULL,
	createuser	varchar(20)			NULL,
	createid	number			NULL,
	createhost	varchar(20)			NULL
)
TABLESPACE TSADVD0; 
COMMIT;

CREATE TABLE advdb.psccustquery
(
	pscname		varchar(50)			NOT NULL,
	sapsystemid	number			NOT NULL,
	rfctype		number			NOT NULL,
	version		number				NOT NULL,
	activ		varchar(1)	DEFAULT 'Y'	NOT NULL,
	text		varchar(2000)			NOT NULL,
	createtime	date		DEFAULT sysdate	NOT NULL,
	createuser	varchar(20)			NULL,
	createid	number			NULL,
	createhost	varchar(20)			NULL,
	comments	varchar(255)			NULL
)
TABLESPACE TSADVD0; 
COMMIT;

CREATE TABLE advdb.psceventlog
(
	pscid		number				NOT NULL,
	sapsystemid	number			NOT NULL,
	rfctype		number			NOT NULL,
	avm		varchar(10)			NULL,
	avmversion	number			NULL,
	motiv		number				NULL,
	adno		number				NULL,
	adver		number				NULL,
	paper		varchar(10)			NULL,
	priority	number		DEFAULT 0	NOT NULL,
	typ		varchar(2)	DEFAULT 'L'	NOT NULL,
	status		number		DEFAULT 0	NOT NULL,
	text		varchar(2000)			NOT NULL,
	createtime	date		DEFAULT sysdate	NOT NULL,
	createuser	varchar(20)			NULL,
	createid	number			NULL,
	createhost	varchar(20)			NULL,
	comments	varchar(1000)			NULL,
	cnt		number				NULL,
	event		varchar(32)			NULL
)
TABLESPACE TSADVD0; 
COMMIT;

CREATE TABLE advdb.psceventcontrol
(
	event		varchar(32)			NOT NULL,
	sapsystemid	number			NOT NULL,
	rfctype		number			NOT NULL,
	runtype		varchar(10)	DEFAULT 'PROD'	NOT NULL,
	version		number				NOT NULL,
	masterversion	number				NOT NULL,
	recno		number		DEFAULT -1	NOT NULL,
	action		varchar(16)			NULL,
	activ		varchar(1)	DEFAULT 'Y'	NOT NULL,
	parameter	varchar(255)			NULL,
	createtime	date		DEFAULT sysdate	NOT NULL,
	createuser	varchar(20)			NULL,
	createid	number			NULL,
	createhost	varchar(20)			NULL,
	comments	varchar(255)			NULL
)
TABLESPACE TSADVD0; 
COMMIT;

CREATE TABLE advdb.pscreportcontrol
(
	controlname	varchar(32)			NOT NULL,
	controltyp      varchar(1)			NOT NULL,
	version		number				NOT NULL,
	length		number				NOT NULL,
	activ		varchar(1)	DEFAULT 'Y'	NOT NULL,
	createtime	date		DEFAULT sysdate	NOT NULL,
	createuser	varchar(20)			NULL,
	createid	number			NULL,
	createhost	varchar(20)			NULL,
	comments	varchar(255)			NULL,
	controldata     blob				NULL
)
TABLESPACE TSADVD0; 
COMMIT;

CREATE TABLE advdb.psccusttables
(
	tablename	varchar(20)			NOT NULL,
	sapsystemid	number			NOT NULL,
	rfctype		number			NOT NULL,
	version		number				NOT NULL,
	activ		varchar(1)	DEFAULT 'Y'	NOT NULL,
	tleft		varchar(50)			NULL,
	tright		varchar(50)			NULL,
	createtime	date		DEFAULT sysdate	NOT NULL,
	createuser	varchar(20)			NULL,
	createid	number			NULL,
	createhost	varchar(20)			NULL
)
TABLESPACE TSADVD0; 
COMMIT;

CREATE TABLE advdb.psccustfields
(
	field		varchar(20)			NOT NULL,
	slevel		number		DEFAULT -1	NOT NULL,
	sapsystemid	number			NOT NULL,
	rfctype		number			NOT NULL,
	version		number				NOT NULL,
	activ		varchar(1)	DEFAULT 'Y'	NOT NULL,
	sapstruct	varchar(20)			NOT NULL,
	sapfield	varchar(20)			NOT NULL,
	first		number				NOT NULL,
	recno		number				NOT NULL,
	length		number				NOT NULL,
	ftype		varchar(20)			NOT NULL,
	createtime	date		DEFAULT sysdate	NOT NULL,
	createuser	varchar(20)			NULL,
	createid	number			NULL,
	createhost	varchar(20)			NULL
)
TABLESPACE TSADVD0; 
COMMIT;

CREATE SEQUENCE advdb.orderid INCREMENT BY 1 START WITH 1 
    MAXVALUE 1.0E28 MINVALUE 1 NOCYCLE 
    CACHE 20 NOORDER;
    
COMMIT;
   
CREATE SEQUENCE advdb.errorid INCREMENT BY 1 START WITH 1 
    MAXVALUE 1.0E28 MINVALUE 1 NOCYCLE 
    CACHE 20 NOORDER;
    
COMMIT;
   
    
CREATE INDEX advdb.pscorderidx
    ON advdb.pscordercontrol  ('aktiv', 'openid', 'status') 
TABLESPACE TSADVD0 PCTFREE 10 INITRANS 2 MAXTRANS 255 
    STORAGE ( INITIAL 64K NEXT 0K MINEXTENTS 1 MAXEXTENTS 
    2147483645 PCTINCREASE 0 FREELISTS 1 FREELIST GROUPS 1) 
    LOGGING;

COMMIT;

CREATE INDEX advdb.pscavmidx
    ON advdb.pscordercontrol  ('avm', 'version')
TABLESPACE TSADVD0 PCTFREE 10 INITRANS 2 MAXTRANS 255 
    STORAGE ( INITIAL 64K NEXT 0K MINEXTENTS 1 MAXEXTENTS 
    2147483645 PCTINCREASE 0 FREELISTS 1 FREELIST GROUPS 1) 
    LOGGING;

COMMIT;

CREATE INDEX advdb.psccontrolididx
    ON advdb.pscordercontrol  ('pscid')
TABLESPACE TSADVD0 PCTFREE 10 INITRANS 2 MAXTRANS 255 
    STORAGE ( INITIAL 64K NEXT 0K MINEXTENTS 1 MAXEXTENTS 
    2147483645 PCTINCREASE 0 FREELISTS 1 FREELIST GROUPS 1) 
    LOGGING;

COMMIT;

CREATE INDEX advdb.pscopenididx
    ON advdb.pscordercontrol  ('openid')
TABLESPACE TSADVD0 PCTFREE 10 INITRANS 2 MAXTRANS 255 
    STORAGE ( INITIAL 64K NEXT 0K MINEXTENTS 1 MAXEXTENTS 
    2147483645 PCTINCREASE 0 FREELISTS 1 FREELIST GROUPS 1) 
    LOGGING;

COMMIT;

CREATE INDEX advdb.pscdataididx
    ON advdb.pscorderdata  ('pscid')
TABLESPACE TSADVD0 PCTFREE 10 INITRANS 2 MAXTRANS 255 
    STORAGE ( INITIAL 64K NEXT 0K MINEXTENTS 1 MAXEXTENTS 
    2147483645 PCTINCREASE 0 FREELISTS 1 FREELIST GROUPS 1) 
    LOGGING;

COMMIT;

CREATE INDEX advdb.psctablesleftidx 
    ON advdb.psccusttables  ('tablename', 'left') 
TABLESPACE "USERS" PCTFREE 10 INITRANS 2 MAXTRANS 255 
    STORAGE ( INITIAL 64K NEXT 0K MINEXTENTS 1 MAXEXTENTS 
    2147483645 PCTINCREASE 0 FREELISTS 1 FREELIST GROUPS 1) 
    LOGGING;
    
COMMIT;