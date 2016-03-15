# Microsoft Developer Studio Project File - Name="ARS DLL" - Package Owner=<4>
# Microsoft Developer Studio Generated Build File, Format Version 6.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Dynamic-Link Library" 0x0102

CFG=ARS DLL - Win32 Debug
!MESSAGE This is not a valid makefile. To build this project using NMAKE,
!MESSAGE use the Export Makefile command and run
!MESSAGE 
!MESSAGE NMAKE /f "ARS DLL.mak".
!MESSAGE 
!MESSAGE You can specify a configuration when running NMAKE
!MESSAGE by defining the macro CFG on the command line. For example:
!MESSAGE 
!MESSAGE NMAKE /f "ARS DLL.mak" CFG="ARS DLL - Win32 Debug"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "ARS DLL - Win32 Release" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE "ARS DLL - Win32 Debug" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE 

# Begin Project
# PROP AllowPerConfigDependencies 0
# PROP Scc_ProjName ""
# PROP Scc_LocalPath ""
CPP=cl.exe
MTL=midl.exe
RSC=rc.exe

!IF  "$(CFG)" == "ARS DLL - Win32 Release"

# PROP BASE Use_MFC 6
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "Release"
# PROP BASE Intermediate_Dir "Release"
# PROP BASE Target_Dir ""
# PROP Use_MFC 6
# PROP Use_Debug_Libraries 0
# PROP Output_Dir "Release"
# PROP Intermediate_Dir "Release"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MD /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_WINDLL" /D "_AFXDLL" /Yu"stdafx.h" /FD /c
# ADD CPP /nologo /MD /W3 /GX /Zd /O2 /I "..\BackupService" /I "..\..\ARSystem 5.1\include" /I "..\..\CBlockSocket Class" /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_WINDLL" /D "_AFXDLL" /D "_MBCS" /D "_AFXEXT" /Fr /Yu"stdafx.h" /FD /c
# SUBTRACT CPP /X
# ADD BASE MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x409 /d "NDEBUG" /d "_AFXDLL"
# ADD RSC /l 0x409 /d "NDEBUG" /d "_AFXDLL"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 /nologo /subsystem:windows /dll /machine:I386
# ADD LINK32 arapi51.lib BlockingSocket_R.lib Ws2_32.lib /nologo /subsystem:windows /dll /pdb:none /map /machine:I386 /def:".\ARS DLL.def" /libpath:"..\..\ARSystem 5.1\lib" /libpath:"..\..\CBlockSocket Class\Lib" /MAPINFO:LINES /MAPINFO:EXPORTS
# Begin Special Build Tool
SOURCE="$(InputPath)"
PostBuild_Cmds=xcopy /y "Release\ARS DLL.dll" "..\BackupService\Release\"	xcopy /y "Release\ARS DLL.dll" "..\BackupClient\Release\"
# End Special Build Tool

!ELSEIF  "$(CFG)" == "ARS DLL - Win32 Debug"

# PROP BASE Use_MFC 6
# PROP BASE Use_Debug_Libraries 1
# PROP BASE Output_Dir "Debug"
# PROP BASE Intermediate_Dir "Debug"
# PROP BASE Target_Dir ""
# PROP Use_MFC 6
# PROP Use_Debug_Libraries 1
# PROP Output_Dir "Debug"
# PROP Intermediate_Dir "Debug"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MDd /W3 /Gm /GX /ZI /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_WINDLL" /D "_AFXDLL" /Yu"stdafx.h" /FD /GZ /c
# ADD CPP /nologo /MD /W3 /Gm /GX /ZI /Od /I "..\BackupService" /I "..\..\ARSystem 5.1\include" /I "..\..\CBlockSocket Class" /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_WINDLL" /D "_AFXDLL" /D "_MBCS" /D "_AFXEXT" /FR /Yu"stdafx.h" /FD /GZ /c
# ADD BASE MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x409 /d "_DEBUG" /d "_AFXDLL"
# ADD RSC /l 0x409 /d "_DEBUG" /d "_AFXDLL"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 /nologo /subsystem:windows /dll /debug /machine:I386 /pdbtype:sept
# ADD LINK32 arapi51.lib BlockingSocket_D.lib Ws2_32.lib /nologo /subsystem:windows /dll /debug /machine:I386 /nodefaultlib:"MSVCRTD.LIB" /pdbtype:sept /libpath:"..\..\ARSystem 5.1\lib" /libpath:"..\..\CBlockSocket Class\Lib"
# SUBTRACT LINK32 /incremental:no /nodefaultlib
# Begin Special Build Tool
SOURCE="$(InputPath)"
PostBuild_Desc=Copying files to Test folder
PostBuild_Cmds=xcopy /y "Debug\ARS DLL.dll" "..\BackupService\Debug\"	xcopy /y "Debug\ARS DLL.dll" "..\BackupClient\Debug\"
# End Special Build Tool

!ENDIF 

# Begin Target

# Name "ARS DLL - Win32 Release"
# Name "ARS DLL - Win32 Debug"
# Begin Group "Source Files"

# PROP Default_Filter "cpp;c;cxx;rc;def;r;odl;idl;hpj;bat"
# Begin Source File

SOURCE=.\AckMsg.cpp
# End Source File
# Begin Source File

SOURCE=".\ARS DLL.cpp"
# End Source File
# Begin Source File

SOURCE=".\ARS DLL.def"

!IF  "$(CFG)" == "ARS DLL - Win32 Release"

# PROP Exclude_From_Build 1

!ELSEIF  "$(CFG)" == "ARS DLL - Win32 Debug"

!ENDIF 

# End Source File
# Begin Source File

SOURCE=".\ARS DLL.rc"
# End Source File
# Begin Source File

SOURCE=.\ARSConnection.cpp
# End Source File
# Begin Source File

SOURCE=.\ARSException.cpp
# End Source File
# Begin Source File

SOURCE=.\BackupJob.cpp
# End Source File
# Begin Source File

SOURCE=.\BlockingSocketFile.cpp
# End Source File
# Begin Source File

SOURCE=.\CommPacket.cpp
# End Source File
# Begin Source File

SOURCE=.\Daily.cpp
# End Source File
# Begin Source File

SOURCE=.\DumpARX.cpp
# End Source File
# Begin Source File

SOURCE=.\EntryId.cpp
# End Source File
# Begin Source File

SOURCE=.\Field.cpp
# End Source File
# Begin Source File

SOURCE=.\FieldList.cpp
# End Source File
# Begin Source File

SOURCE=.\Form.cpp
# End Source File
# Begin Source File

SOURCE=.\FormList.cpp
# End Source File
# Begin Source File

SOURCE=.\Globals.cpp
# End Source File
# Begin Source File

SOURCE=.\Monthly.cpp
# End Source File
# Begin Source File

SOURCE=.\Record.cpp
# End Source File
# Begin Source File

SOURCE=.\RecordList.cpp
# End Source File
# Begin Source File

SOURCE=.\Schedule.cpp
# End Source File
# Begin Source File

SOURCE=.\StdAfx.cpp
# ADD CPP /Yc"stdafx.h"
# End Source File
# Begin Source File

SOURCE=.\Weekly.cpp
# End Source File
# End Group
# Begin Group "Header Files"

# PROP Default_Filter "h;hpp;hxx;hm;inl"
# Begin Source File

SOURCE=.\AckMsg.h
# End Source File
# Begin Source File

SOURCE=.\AREErrorNumbers.h
# End Source File
# Begin Source File

SOURCE=.\ARSConnection.h
# End Source File
# Begin Source File

SOURCE=.\ARSException.h
# End Source File
# Begin Source File

SOURCE=.\BackupJob.h
# End Source File
# Begin Source File

SOURCE=.\BlockingSocketFile.h
# End Source File
# Begin Source File

SOURCE=.\CommPacket.h
# End Source File
# Begin Source File

SOURCE=.\Daily.h
# End Source File
# Begin Source File

SOURCE=.\DaysOfMonth.h
# End Source File
# Begin Source File

SOURCE=.\DumpARX.h
# End Source File
# Begin Source File

SOURCE=.\Form.h
# End Source File
# Begin Source File

SOURCE=.\FormList.h
# End Source File
# Begin Source File

SOURCE=.\Globals.h
# End Source File
# Begin Source File

SOURCE=.\Monthly.h
# End Source File
# Begin Source File

SOURCE=.\Record.h
# End Source File
# Begin Source File

SOURCE=.\RecordList.h
# End Source File
# Begin Source File

SOURCE=.\Resource.h
# End Source File
# Begin Source File

SOURCE=.\Schedule.h
# End Source File
# Begin Source File

SOURCE=.\StdAfx.h
# End Source File
# Begin Source File

SOURCE=.\ToDos.h
# End Source File
# Begin Source File

SOURCE=.\Weekly.h
# End Source File
# End Group
# Begin Group "Resource Files"

# PROP Default_Filter "ico;cur;bmp;dlg;rc2;rct;bin;rgs;gif;jpg;jpeg;jpe"
# Begin Source File

SOURCE=".\res\ARS DLL.rc2"
# End Source File
# End Group
# Begin Source File

SOURCE=.\ReadMe.txt
# End Source File
# End Target
# End Project
