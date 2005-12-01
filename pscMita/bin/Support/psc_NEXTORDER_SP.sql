CREATE OR REPLACE PROCEDURE advdb.PSC_NEXTORDER_SP  (
 f_myid IN NUMBER,
 f_sapsystem IN NUMBER,
 f_sapversion IN NUMBER,
 f_pscid OUT NUMBER,
 f_processed OUT NUMBER
)
AS
  t_count NUMBER;
  t_id NUMBER;
  f_avm VARCHAR(10);
  f_version NUMBER;
  CURSOR my_cursor IS
    SELECT pscid, avm, version, processed FROM pscordercontrol
      WHERE status = 'N'
        AND (
          (openid IS NULL OR openid = f_myid)
          AND activ = 'Y'
          OR (openid = f_myid AND activ = 'W')
        )
        AND sapversionid = f_sapversion
        AND sapsystemid = f_sapsystem
        ORDER BY createtime DESC
        FOR UPDATE;
        
  BEGIN
    f_pscid := -1;
    f_avm := '0000000000';
    f_version := -1;
     
    OPEN my_cursor;
    LOOP
      FETCH my_cursor into t_id, f_avm, f_version, f_processed;
      EXIT WHEN my_cursor%NOTFOUND OR
        my_cursor%NOTFOUND IS NULL;
      SELECT COUNT(avm) into t_count FROM pscordercontrol
        WHERE status = 'N'
          AND openid <> f_myid
          AND activ = 'W'
          AND avm = f_avm
          AND sapversionid = f_sapversion
          AND sapsystemid = f_sapsystem;
          
      IF t_count = 0 THEN
        f_pscid := t_id;     
        UPDATE pscordercontrol SET activ = 'W', openid = f_myid
            WHERE CURRENT OF my_cursor;
        COMMIT;
        CLOSE my_cursor; 
         EXIT; 
      END IF;
    END LOOP;
  END;