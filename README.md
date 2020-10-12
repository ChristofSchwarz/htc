# Code for HTC Tourismuscontrolling GmbH

Customer work of mine. 

To re-publish from a local git copy, run
```
git pull  #if some changes were made elsewhere to this repo
sh push.sh
```
To include this Qlik Script files (*.qvf) from Qlik cloud, use the REST Connector in this way:
```
LIB CONNECT TO 'Igel:REST_GET';

// Code ausgelagert auf https://github.com/ChristofSchwarz/htc
$URLs: LOAD * INLINE [
  https://raw.githubusercontent.com/ChristofSchwarz/htc/main/01_Main.qvs
  https://raw.githubusercontent.com/ChristofSchwarz/htc/main/02_Variables.qvs
  https://raw.githubusercontent.com/ChristofSchwarz/htc/main/03_Preload%20Data.qvs
  https://raw.githubusercontent.com/ChristofSchwarz/htc/main/04_ExcelStruktur.qvs
  https://raw.githubusercontent.com/ChristofSchwarz/htc/main/05_Berichte%20postprocessing.qvs
  https://raw.githubusercontent.com/ChristofSchwarz/htc/main/06_Join%20Werte.qvs
  https://raw.githubusercontent.com/ChristofSchwarz/htc/main/07_Datenart.qvs
  https://raw.githubusercontent.com/ChristofSchwarz/htc/main/08_Betriebe.qvs
  https://raw.githubusercontent.com/ChristofSchwarz/htc/main/09_Kalender.qvs
  https://raw.githubusercontent.com/ChristofSchwarz/htc/main/99_Exit.qvs
] (no labels, delimiter is '\n') WHERE NOT @1 LIKE '//*';

FOR vUrlIdx = 1 TO NoOfRows('$URLs')
  LET vUrl = Peek('@1', vUrlIdx-1, '$URLs');
  LET vVarName = Replace(SubField(vUrl, '/', -1),'%20',' ');
  SCRIPT: LOAD Concat(col_1, CHR(10), RecNo()) AS script;
  SQL SELECT "col_1" FROM CSV (header off, delimiter "\n", quote """") "CSV_source"
  WITH CONNECTION (URL "$(vUrl)");
  LET [$(vVarName)] = Peek('script', -1, 'SCRIPT');
  TRACE; TRACE Created Script in variable <<$(vVarName)>>;
  DROP TABLE SCRIPT;
NEXT vUrlIdx;
DROP TABLE $URLs;

$(01_Main.qvs);
$(02_Variables.qvs);
$(03_Preload Data.qvs);
$(04_ExcelStruktur.qvs);
$(05_Berichte postprocessing.qvs);
$(06_Join Werte.qvs);
$(07_Datenart.qvs);                 
$(08_Betriebe.qvs); 
$(09_Kalender.qvs);                 
$(99_Exit.qvs);
```
The script raw lines are read in the FOR loop using any REST GET connection without authentication, 
then the lines are concatenated into a Line-Break separated string and put into a variable that
is named like the file itself (last part of the url behind the last "/")
