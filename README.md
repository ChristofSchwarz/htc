# Code for HTC Tourismuscontrolling GmbH

Customer work of mine. 

To re-publish from a local git copy, run
```
git pull  #if some changes were made elsewhere to this repo
sh push.sh
```
To include this Qlik Script files (*.qvf) from Qlik cloud, use the REST Connector in this way:
```
LIB CONNECT TO '<<< put your http-REST-GET lib connection >>>';
$URLs: 
LOAD 
  @1 AS $url,
  If(Len(@2), @2, Replace(SubField(@1,'/',-1),'%20',' ')) AS $varName
INLINE [
  https://raw.githubusercontent.com/ChristofSchwarz/htc/main/00_Subs.qvs | 00_Subs.qvs
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
] (no labels, delimiter is '|') WHERE NOT @1 LIKE '//*';

// Keep next block ...
FOR vUrlIdx = 1 TO NoOfRows('$URLs')
  LET vUrl = Peek('$url', vUrlIdx-1, '$URLs');
  LET vVarName = Peek('$varName', vUrlIdx-1, '$URLs');
  $ScriptRows: LOAD Concat(col_1, CHR(10), RecNo()) AS $script;
  SQL SELECT "col_1" FROM CSV (header off, delimiter "\n", quote "\n") "CSV_source"
  WITH CONNECTION (URL "$(vUrl)", HTTPHEADER "Cache-Control" "no-cache");
  LET [$(vVarName)] = Peek('$script', -1, '$ScriptRows');
  LET vVarLen = Num(Len([$(vVarName)])/1024,'# ##0.00','.',' ');
  IF vVarLen = 0 THEN
  	[Script not found at $(vUrl)];  // Throw error
  ENDIF
  TRACE; TRACE Created $(vVarLen) kB of Script in variable <<$(vVarName)>>;
  DROP TABLE $ScriptRows;
NEXT vUrlIdx;
DROP TABLE $URLs;  

$(00_Subs.qvs);
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
