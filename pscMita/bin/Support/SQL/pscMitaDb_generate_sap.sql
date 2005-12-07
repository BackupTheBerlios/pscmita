CREATE TABLE advdb.pscsapsystems
(	sapsystemid	number				NOT NULL,
	sapname		varchar(10)			NOT NULL,
	version		number				NOT NULL,
	sapversionid	number				NOT NULL,
	environment	varchar(10)			NOT NULL,
	activ		varchar(1)	DEFAULT 'Y'	NOT NULL,
	sapgateway	varchar(20)			NOT NULL,
	sapservice	varchar(20)			NOT NULL,
	sapid		varchar(20)			NOT NULL,
	sapidclt	varchar(20)			NOT NULL,
	sapuser		varchar(10)			NOT NULL,
	sapowner	varchar(10)			NOT NULL,
	sapsystem	varchar(10)			NOT NULL,
	sapserver	varchar(10)			NOT NULL,
	sapclient	varchar(10)			NOT NULL,
	databas		varchar(20)			NOT NULL,
	databasuser	varchar(20)			NOT NULL,
	databaspwd	varchar(20)			NOT NULL,
	createtime	date		DEFAULT sysdate	NOT NULL,
	createuser	varchar(20)			NULL,
	createid	number			NULL,
	createhost	varchar(20)			NULL
)
TABLESPACE TSADVD0; 
COMMIT;

CREATE TABLE advdb.pscsapversions
(	sapversionid	number				NOT NULL,
	sapversionname	varchar(10)		NOT NULL,
	sapsubversion	varchar(10)		NULL,
	version		number				NOT NULL,
	activ		varchar(1)	DEFAULT 'Y'	NOT NULL,
	createtime	date		DEFAULT sysdate	NOT NULL,
	createuser	varchar(20)			NULL,
	createid	number			NULL,
	createhost	varchar(20)			NULL,
	replacestruct	varchar(255)			NULL,
	replacefunc	varchar(255)			null
)
TABLESPACE TSADVD0; 
COMMIT;

