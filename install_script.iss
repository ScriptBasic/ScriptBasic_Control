[Setup]
AppName=SB_Debug
AppVerName=ScriptBasic Debugger .1
DefaultDirName=c:\ScriptBasic\Debugger
DefaultGroupName=ScriptBasic
UninstallDisplayIcon={app}\unins000.exe
OutputDir=./
OutputBaseFilename=SBDebug_Setup


[Dirs]
Name: {app}\dependancies
Name: {app}\include
Name: {app}\modules
Name: {app}\scripts


[Files]
Source: dependancies\calltips.txt; DestDir: {app}\dependancies\
Source: dependancies\MSCOMCTL.OCX; DestDir: {win}\system32\; Flags: regserver
Source: dependancies\SciLexer.dll; DestDir: {app}\dependancies\
Source: dependancies\scivb_lite.ocx; DestDir: {app}\dependancies\; Flags: regserver
Source: dependancies\spSubclass.dll; DestDir: {app}\dependancies\; Flags: regserver
Source: dependancies\VB.bin; DestDir: {app}\dependancies\
Source: engine\sb_engine.dll; DestDir: {app}
Source: include\COM.bas; DestDir: {app}\include\
Source: include\curl.bas; DestDir: {app}\include\
Source: include\mysql.bas; DestDir: {app}\include\
Source: include\nt.bas; DestDir: {app}\include\
Source: include\odbc.bas; DestDir: {app}\include\
Source: include\re.bas; DestDir: {app}\include\
Source: modules\curl.dll; DestDir: {app}\modules\
Source: modules\mysql.dll; DestDir: {app}\modules\
Source: modules\nt.dll; DestDir: {app}\modules\
Source: modules\odbc.dll; DestDir: {app}\modules\
Source: modules\re.dll; DestDir: {app}\modules\
Source: README.txt; DestDir: {app}
Source: SBDebug.exe; DestDir: {app}
Source: SBDebug.pdb; DestDir: {app}
Source: scripts\com_voice_test.sb; DestDir: {app}\scripts\

[Icons]
Name: {userdesktop}\SBDebug.exe; Filename: {app}\SBDebug.exe; IconIndex: 0

[Run]

Filename: {app}\README.txt; StatusMsg: View README; Flags: shellexec
