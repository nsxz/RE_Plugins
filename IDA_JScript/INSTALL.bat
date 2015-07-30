@echo off
Echo registering dependancies
regsvr32 vbdevkit.dll
regsvr32 spsubclass.dll
regsvr32 scivb2.ocx
echo Make sure to manually install the IDASRVR.plw IDA plugin..
pause

