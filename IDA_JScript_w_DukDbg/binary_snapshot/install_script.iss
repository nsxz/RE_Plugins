[Setup]
AppName=IdaJSDBG
AppVerName=IdaJSDBG 0.0.1
DefaultDirName=c:\IdaJSDBG
DefaultGroupName=IdaJSDBG
UninstallDisplayIcon={app}\unins000.exe
OutputDir=./
OutputBaseFilename=IdaJSDBG_Setup


[Dirs]
Name: {app}\COM
Name: {app}\scripts


[Files]
Source: dukDbg.ocx; DestDir: {app}; Flags: regserver
Source: IDA_JScript.exe; DestDir: {app}
Source: spSubclass.dll; DestDir: {app}; Flags: regserver
Source: SciLexer.dll; DestDir: {app}
Source: scivb2.ocx; DestDir: {app}; Flags: regserver
Source: vbDevKit.dll; DestDir: {app}; Flags: regserver
Source: Duk4VB.dll; DestDir: {app}
Source: ..\COM\ida.js; DestDir: {app}\COM\
Source: ..\COM\list.js; DestDir: {app}\COM\
Source: ..\COM\TextBox.js; DestDir: {app}\COM\
Source: ..\COM\remote.js; DestDir: {app}\COM\
Source: ..\scripts\extractFuncNames.idajs; DestDir: {app}\scripts\
Source: ..\scripts\extractNamesRange.idajs; DestDir: {app}\scripts\
Source: ..\scripts\funcCalls.idajs; DestDir: {app}\scripts\
Source: ..\api.txt; DestDir: {app}
Source: ..\beautify.js; DestDir: {app}
Source: ..\java.hilighter; DestDir: {app}
Source: ..\userlib.js; DestDir: {app}
Source: ..\..\IDASrvr\bin\IDASrvr.plw; DestDir: {app}
Source: MSCOMCTL.OCX; DestDir: {win}; Flags: regserver uninsneveruninstall
Source: richtx32.ocx; DestDir: {win}; Flags: regserver uninsneveruninstall

[Icons]
Name: {group}\IDA_Jscript; Filename: {app}\IDA_JScript.exe
Name: {group}\Uninstall; Filename: {app}\unins000.exe
;Name: {userdesktop}\IDA_Jscript; Filename: {app}\IDA_Jscript.exe; IconIndex: 0


[Messages]
FinishedLabel=Remember to install the plw into your IDA plugins directory.
