CREATE OR REPLACE PROCEDURE "ADVDB"."PSC_SAPPOOL_SP"  (
  p_TABLE in varchar2,
  p_INTSAPNO in varchar2,
  p_INTCOMBONO  in number,
  p_CLIENTNO in varchar2,
  p_FILEPATH in varchar2,
  p_INTADNO in number,
  p_ORDERVNO in number,
  p_INTMONR in number,
  p_INTBOXNO in varchar2,
  p_TOPTEXT in varchar2,
  p_TDEPTH in number,
  p_INTPOSNR in number
    )
AS
	sapcount number;
	BEGIN
	select count(*) into sapcount from p_TABLE where sapno=p_INTSAPNO and combono=p_INTCOMBONO and clientno=p_CLIENTNO;
	IF sapcount >0 THEN
  		update p_TABLE set adno=p_INTADNO,vno=p_ORDERVNO,posno=p_INTMONR, boxno=p_INTBOXNO, timestamp=sysdate, pubno=p_INTPOSNR
  		where sapno=p_INTSAPNO and combono=p_INTCOMBONO and clientno=p_CLIENTNO;
	ELSE
  		insert into p_TABLE (sapno,adno,combono,filename,vno,posno,pubno,clientno,boxno,sortword,ysizemm)
  	alues (p_INTSAPNO,p_INTADNO,p_INTCOMBONO,p_FILEPATH,p_ORDERVNO,p_INTMONR,p_INTPOSNR,p_CLIENTNO,p_INTBOXNO,p_TOPTEXT,p_TDEPTH);
	END IF;
	END;