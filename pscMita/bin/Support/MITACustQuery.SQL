[SAPSYSTEMID=1]

SetOldVersion_1 = BEGIN sx_UpdateToOldVer(#APPADNO#); END;

InsertAd_1 = BEGIN sx_NewAd(#APPADNO#,#APPVNO#,'#ADRNAME#','#ADRTEL#',#ADRNO#,'#NAME1#',
InsertAd_2 = '#PHONE#',#CUSNO#,'#MITAPAPER#','#REMARKS#',
InsertAd_3 = '#BOXNO#','#BOXTYPE#',#APPPUBCNT#,'#ADPRODTYPE#','#ADTYPE#',
InsertAd_4 = '#USERID#','#SAPNO#','#MITACLIENTNO#',#PAPERNO#,'#REFAVMNR#','#CAMPAIGN#','#BLED#',
InsertAd_5 = '#STRASSE#','#HAUSNR#','#HSNMR2#','#POSTNO#','#PROVINCE#',
InsertAd_6 = '#PHONE1#','#FAX1#','#FAX#','#COUNTRY#','#EPSCHK#',#APPVNO#,'#MITASAPNO#','#APPADV#'); 
InsertAd_7 = sx_OrderPub(#APPADNO#,#APPVNO#);
InsertAd_8 = UPDATE ad SET pubcnt = #APPPUBCNT# WHERE adno = #APPADNO# AND vnoflag = 'Y';
InsertAd_9 = END;

InsertPub_1 = BEGIN sx_NewPub(#APPADNO#,#APPVNO#,#APPPUBCNT#,#ITEMNO#,'#RUNDATE#',
InsertPub_2 = '#ADTYPE#',#WIDTH#,#COLS#,#WORDS#,#REQWORDS#,#DEPTH#,
InsertPub_3 = '#PPAGE#','#CLASS#','#MITAPAPER#','#APPSTATUS#',#APPDEPTH#,'#APPCREDP#',
InsertPub_4 = #APPFP#,'#REFAVMNR#','#PUBREADY#','#EPSCHK#','#COLOR#',
InsertPub_5 = '#USERID#','#EVENODD#','#POSITION#','#POSREMARKS#',#APPXLOC#,#APPYLOC#,
InsertPub_6 = '#APPTEXT#','#KEYWORD#',#CUSNO#,'#FORMA#','#NOTIZ#','#ORDERTYPE#','#CAMPANHA#',#SCHEDNO#,
InsertPub_7 = #PSVNO#, #PAPVNO#, #VNO#); 
InsertPub_8 = sx_NewTrack(#APPADNO#,#APPPUBCNT#,#APPVNO#,'#PRODSTEP#','#USERID#',#WIDTH#,#DEPTH#);
InsertPub_9 = sx_NewText(#APPADNO#,#APPVNO#,'#APPADV#','#APPTEXT#',
InsertPub_10 = '#ADTYPE#','#KEYWORD#','#APPTEXT#','#TEXTTYPE#','#REFAVMNR#'); END;

UpdateAdFinally_1 = UPDATE ad SET startdate = to_date('#APPSTARTDATE#','YYYYMMDD'), 
UpdateAdFinally_2 = enddate = to_date('#APPENDDATE#','YYYYMMDD')
UpdateAdFinally_3 = WHERE adno = #APPADNO# AND vnoflag = 'Y';

NewAdNo_1 = SELECT newadno.nextval FROM dual;

GetFileName_1 = SELECT filename FROM #MITAPOOL# WHERE adno = #APPADNO# AND combono = #MITACOMBONO#

ReadSapPool_1 = SELECT * FROM #MITAPOOL# WHERE sapno = '#MITASAPNO#' AND clientno = '#MITACLIENTNO#'
ReadSapPool_2 = AND combono = #MITACOMBONO#;

GetSapVersion_1 = SELECT vno FROM #MITAPOOL# WHERE sapno LIKE '#MITAORDERNO#%' AND clientno = '#MITACLIENTNO#';

GetReferenceAdNo_1 = SELECT adno FROM #MITAPOOL# WHERE sapno LIKE '#APPREFAVMNO#%' AND clientno = '#MITACLIENTNO#' AND pubno = #APPREFPOSNO#;

UpdateSapPool_1 = BEGIN psc_sappool_sp('#MITAPOOL#','#MITASAPNO#',#APPCOMBONO#,'#MITACLIENTNO#','#APPADV#',#APPADNO#,
UpdateSapPool_2 = #MITAVNO#,#MITAMOTIVNO#,'#BOXNO#','#APPTEXT#',#APPDEPTH#, #MITAPOSNO#); END;

UpdateSapVersion_1 = UPDATE #MITAPOOL# SET vno = #MITAVNO# WHERE sapno LIKE '#MITAORDERNO#%' AND clientno = '#MITACLIENTNO#';

UpdateSapPub_1 = UPDATE #MITAPOOL# SET pubno = #MITAPOSNO# WHERE sapno LIKE '#MITAORDERNO#%' AND clientno = '#MITACLIENTNO#';

WriteReference_1 = UPDATE ad SET pubdays = #APPREFADNO#, branch = 'C' WHERE adno = #APPADNO# AND vnoflag = 'Y';



[SAPSYSTEMID=2]

SetOldVersion_1 = BEGIN sx_UpdateToOldVer(#APPADNO#); END;

InsertAd_1 = BEGIN sx_NewAd(#APPADNO#,#APPVNO#,'#ADRNAME#','#ADRTEL#',#ADRNO#,'#NAME1#',
InsertAd_2 = '#PHONE#',#CUSNO#,'#MITAPAPER#','#REMARKS#',
InsertAd_3 = '#BOXNO#','#BOXTYPE#',#APPPUBCNT#,'#ADPRODTYPE#','#ADTYPE#',
InsertAd_4 = '#USERID#','#SAPNO#','#MITACLIENTNO#',#PAPERNO#,'#REFAVMNR#','#CAMPAIGN#','#BLED#',
InsertAd_5 = '#STRASSE#','#HAUSNR#','#HSNMR2#','#POSTNO#','#PROVINCE#',
InsertAd_6 = '#PHONE1#','#FAX1#','#FAX#','#COUNTRY#','#EPSCHK#',#APPVNO#,'#MITASAPNO#','#APPADV#'); 
InsertAd_7 = sx_OrderPub(#APPADNO#,#APPVNO#);
InsertAd_8 = UPDATE ad SET pubcnt = #APPPUBCNT# WHERE adno = #APPADNO# AND vnoflag = 'Y';
InsertAd_9 = END;

InsertPub_1 = BEGIN sx_NewPub(#APPADNO#,#APPVNO#,#APPPUBCNT#,#ITEMNO#,'#RUNDATE#',
InsertPub_2 = '#ADTYPE#',#WIDTH#,#COLS#,#WORDS#,#REQWORDS#,#DEPTH#,
InsertPub_3 = '#PPAGE#','#CLASS#','#MITAPAPER#','#APPSTATUS#',#APPDEPTH#,'#APPCREDP#',
InsertPub_4 = #APPFP#,'#REFAVMNR#','#PUBREADY#','#EPSCHK#','#COLOR#',
InsertPub_5 = '#USERID#','#EVENODD#','#POSITION#','#POSREMARKS#',#APPXLOC#,#APPYLOC#,
InsertPub_6 = '#APPTEXT#','#KEYWORD#',#CUSNO#,'#FORMA#','#NOTIZ#','#ORDERTYPE#','#CAMPANHA#',#SCHEDNO#,
InsertPub_7 = #PSVNO#, #PAPVNO#, #VNO#); 
InsertPub_8 = sx_NewTrack(#APPADNO#,#APPPUBCNT#,#APPVNO#,'#PRODSTEP#','#USERID#',#WIDTH#,#DEPTH#);
InsertPub_9 = sx_NewText(#APPADNO#,#APPVNO#,'#APPADV#','#APPTEXT#',
InsertPub_10 = '#ADTYPE#','#KEYWORD#','#APPTEXT#','#TEXTTYPE#','#REFAVMNR#'); END;

UpdateAdFinally_1 = UPDATE ad SET startdate = to_date('#APPSTARTDATE#','YYYYMMDD'), 
UpdateAdFinally_2 = enddate = to_date('#APPENDDATE#','YYYYMMDD'),
UpdateAdFinally_3 = WHERE adno = #APPADNO# AND vnoflag = 'Y';

NewAdNo_1 = SELECT newadno.nextval FROM dual;

NewContentNo_1 = SELECT newcontent.nextval FROM dual;

GetFileName_1 = SELECT filename FROM #MITAPOOL# WHERE adno = #APPADNO# AND combono = #MITACOMBONO#

ReadSapPool_1 = SELECT * FROM #MITAPOOL# WHERE sapno = '#MITASAPNO#' AND clientno = '#MITACLIENTNO#'
ReadSapPool_2 = AND combono = #MITACOMBONO#;

GetSapVersion_1 = SELECT vno FROM #MITAPOOL# WHERE sapno LIKE '#MITAORDERNO#%' AND clientno = '#MITACLIENTNO#';

GetReferenceAdNo_1 = SELECT adno FROM #MITAPOOL# WHERE sapno LIKE '#APPREFAVMNO#%' AND clientno = '#MITACLIENTNO#' AND pubno = #APPREFPOSNO#;

UpdateSapPool_1 = BEGIN psc_sappool_sp('#MITAPOOL#','#MITASAPNO#',#APPCOMBONO#,'#MITACLIENTNO#','#APPADV#',#APPADNO#,
UpdateSapPool_2 = #MITAVNO#,#MITAMOTIVNO#,'#BOXNO#','#APPTEXT#',#APPDEPTH#, #MITAPOSNO#); END;

UpdateSapVersion_1 = UPDATE #MITAPOOL# SET vno = #MITAVNO# WHERE sapno LIKE '#MITAORDERNO#%' AND clientno = '#MITACLIENTNO#';

UpdateSapPub_1 = UPDATE #MITAPOOL# SET pubno = #MITAPOSNO# WHERE sapno LIKE '#MITAORDERNO#%' AND clientno = '#MITACLIENTNO#';

WriteReference_1 = UPDATE ad SET pubdays = #APPREFADNO#, branch = 'C' WHERE adno = #APPADNO# AND vnoflag = 'Y';



