# Microsoft Developer Studio Project File - Name="OPCAuto" - Package Owner=<4>
# Microsoft Developer Studio Generated Build File, Format Version 6.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Dynamic-Link Library" 0x0102

CFG=OPCAuto - Win32 Debug
!MESSAGE This is not a valid makefile. To build this project using NMAKE,
!MESSAGE use the Export Makefile command and run
!MESSAGE 
!MESSAGE NMAKE /f "OPCAuto.mak".
!MESSAGE 
!MESSAGE You can specify a configuration when running NMAKE
!MESSAGE by defining the macro CFG on the command line. For example:
!MESSAGE 
!MESSAGE NMAKE /f "OPCAuto.mak" CFG="OPCAuto - Win32 Debug"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "OPCAuto - Win32 Debug" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE "OPCAuto - Win32 Release" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE 

# Begin Project
# PROP AllowPerConfigDependencies 0
CPP=cl.exe
MTL=midl.exe
RSC=rc.exe

!IF  "$(CFG)" == "OPCAuto - Win32 Debug"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 1
# PROP BASE Output_Dir "Debug"
# PROP BASE Intermediate_Dir "Debug"
# PROP BASE Ignore_Export_Lib 0
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 1
# PROP Output_Dir "Debug"
# PROP Intermediate_Dir "Debug"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MTd /W3 /Gm /GX /ZI /Od /I "..\..\Include" /I "$(CommonProgramFiles)\OPC Foundation\Include" /D "_DEBUG" /D "_UNICODE" /D "WIN32" /D "_WINDOWS" /D "_USRDLL" /Yu"stdafx.h" /FD /c
# ADD CPP /nologo /MTd /W3 /Gm /GX /ZI /Od /I "..\..\..\Include" /I "$(CommonProgramFiles)\OPC Foundation\Include" /D "_DEBUG" /D "WIN32" /D "_WINDOWS" /D "_USRDLL" /Yu"StdAfx.h" /FD /c
# ADD BASE MTL /nologo /D "_DEBUG" /mktyplib203 /o "NUL" /win32
# ADD MTL /nologo /D "_DEBUG" /o "NUL" /win32
# SUBTRACT MTL /mktyplib203
# ADD BASE RSC /l 0x409 /d "_DEBUG"
# ADD RSC /l 0x409 /d "_DEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 odbccp32.lib ole32.lib oleaut32.lib uuid.lib /nologo /base:"0x1f500000" /subsystem:windows /dll /debug /machine:I386 /out:"DebugU/OPCDAAuto.dll" /pdbtype:sept
# ADD LINK32 odbccp32.lib ole32.lib oleaut32.lib uuid.lib version.lib Shlwapi.lib /nologo /base:"0x1f500000" /subsystem:windows /dll /debug /machine:I386 /out:"Debug/OPCDAAuto.dll" /pdbtype:sept
# Begin Custom Build - Registering ActiveX Control...
OutDir=.\Debug
TargetPath=.\Debug\OPCDAAuto.dll
InputPath=.\Debug\OPCDAAuto.dll
SOURCE="$(InputPath)"

"$(OutDir)\regsvr32.trg" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	regsvr32 /s /c "$(TargetPath)" 
	echo regsvr32 exec. time > "$(OutDir)\regsvr32.trg" 
	
# End Custom Build

!ELSEIF  "$(CFG)" == "OPCAuto - Win32 Release"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "Release"
# PROP BASE Intermediate_Dir "Release"
# PROP BASE Ignore_Export_Lib 0
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 0
# PROP Output_Dir "Release"
# PROP Intermediate_Dir "Release"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MT /W3 /GX /O1 /I "..\..\Include" /I "$(CommonProgramFiles)\OPC Foundation\Include" /D "NDEBUG" /D "_UNICODE" /D "_ATL_STATIC_REGISTRY" /D "_ATL_MIN_CRT" /D "WIN32" /D "_WINDOWS" /D "_USRDLL" /Yu"stdafx.h" /FD /c
# ADD CPP /nologo /MT /W3 /GX /O1 /I "..\..\..\Include" /I "$(CommonProgramFiles)\OPC Foundation\Include" /D "NDEBUG" /D "_ATL_STATIC_REGISTRY" /D "_ATL_MIN_CRT" /D "WIN32" /D "_WINDOWS" /D "_USRDLL" /Yu"StdAfx.h" /FD /c
# ADD BASE MTL /nologo /D "NDEBUG" /mktyplib203 /o "NUL" /win32
# ADD MTL /nologo /D "NDEBUG" /o "NUL" /win32
# SUBTRACT MTL /mktyplib203
# ADD BASE RSC /l 0x409 /d "NDEBUG"
# ADD RSC /l 0x409 /d "NDEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 odbccp32.lib ole32.lib oleaut32.lib uuid.lib /nologo /base:"0x1f500000" /subsystem:windows /dll /machine:I386 /out:"ReleaseUMinDependency/OPCDAAuto.dll"
# ADD LINK32 odbccp32.lib ole32.lib oleaut32.lib uuid.lib version.lib Shlwapi.lib /nologo /base:"0x1f500000" /subsystem:windows /dll /machine:I386 /out:"Release/OPCDAAuto.dll"
# Begin Custom Build - Registering ActiveX Control...
OutDir=.\Release
TargetPath=.\Release\OPCDAAuto.dll
InputPath=.\Release\OPCDAAuto.dll
SOURCE="$(InputPath)"

"$(OutDir)\regsvr32.trg" : $(SOURCE) "$(INTDIR)" "$(OUTDIR)"
	regsvr32 /s /c "$(TargetPath)" 
	echo regsvr32 exec. time > "$(OutDir)\regsvr32.trg" 
	
# End Custom Build

!ENDIF 

# Begin Target

# Name "OPCAuto - Win32 Debug"
# Name "OPCAuto - Win32 Release"
# Begin Group "Source Files"

# PROP Default_Filter "cpp;c;cxx;rc;def;r;odl;idl;hpj;bat"
# Begin Source File

SOURCE=.\OPCAuto.cpp
# End Source File
# Begin Source File

SOURCE=.\OPCAuto.def
# End Source File
# Begin Source File

SOURCE=.\OPCAuto.idl
# PROP BASE Ignore_Default_Tool 1
# ADD MTL /tlb "./OPCAuto.tlb" /h "./OPCAuto.h" /iid "./OPCAuto_i.c"
# End Source File
# Begin Source File

SOURCE=.\OPCAuto.rc
# End Source File
# Begin Source File

SOURCE=.\OPCAutoServer.cpp
# End Source File
# Begin Source File

SOURCE=.\OPCBrowser.cpp
# End Source File
# Begin Source File

SOURCE=.\OPCGroup.cpp
# End Source File
# Begin Source File

SOURCE=.\OPCGroupCallback.cpp
# End Source File
# Begin Source File

SOURCE=.\OPCGroups.cpp
# End Source File
# Begin Source File

SOURCE=.\OPCItem.cpp
# End Source File
# Begin Source File

SOURCE=.\OPCItems.cpp
# End Source File
# Begin Source File

SOURCE=.\OPCTrace.cpp
# End Source File
# Begin Source File

SOURCE=.\OPCUtils.cpp
# End Source File
# Begin Source File

SOURCE=.\StdAfx.cpp
# ADD BASE CPP /Yc"stdafx.h"
# ADD CPP /Yc"stdafx.h"
# End Source File
# End Group
# Begin Group "Header Files"

# PROP Default_Filter "h;hpp;hxx;hm;inl"
# Begin Source File

SOURCE=.\OPCActivator.h
# End Source File
# Begin Source File

SOURCE=.\OPCAutoServer.h
# End Source File
# Begin Source File

SOURCE=.\OPCBrowser.h
# End Source File
# Begin Source File

SOURCE=.\OPCGroup.h
# End Source File
# Begin Source File

SOURCE=.\OPCGroupEvent.h
# End Source File
# Begin Source File

SOURCE=.\OPCGroups.h
# End Source File
# Begin Source File

SOURCE=.\OPCItem.h
# End Source File
# Begin Source File

SOURCE=.\OPCItems.h
# End Source File
# Begin Source File

SOURCE=.\OPCTrace.h
# End Source File
# Begin Source File

SOURCE=.\OPCUtils.h
# End Source File
# Begin Source File

SOURCE=.\Resource.h
# End Source File
# Begin Source File

SOURCE=.\StdAfx.h
# End Source File
# End Group
# End Target
# End Project
