@echo off

IF NOT EXIST C:\IDA6.5 GOTO NO65
echo Installing for 6.5
del C:\IDA6.5\plugins\idasrvr.plw
copy D:\_RE_Plugins\IDASrvr\bin\idasrvr.plw C:\IDA6.5\plugins\

:NO65

IF NOT EXIST C:\IDA6 GOTO NO6
echo Installing for 6
del C:\IDA6\plugins\idasrvr.plw
copy D:\_RE_Plugins\IDASrvr\bin\idasrvr.plw C:\IDA6\plugins\

:NO6


IF NOT EXIST C:\IDA GOTO NO5
echo Installing for c:\IDA
del C:\IDA\plugins\idasrvr.plw
copy D:\_RE_Plugins\IDASrvr\bin\idasrvr.plw C:\IDA\plugins\

:NO5

if not exist D:\_Installs\iDef\github\IDACompare goto NO
echo Installing for IDAcompare
copy D:\_RE_Plugins\IDASrvr\bin\idasrvr.plw D:\_Installs\iDef\github\IDACompare\

:NO

pause