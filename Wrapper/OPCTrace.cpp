//==============================================================================
// TITLE: OPCTrace.cpp
//
// CONTENTS:
// 
// The implementation of a debug tracing class.
//
// (c) Copyright 2003-2004 The OPC Foundation
// ALL RIGHTS RESERVED.
//
// DISCLAIMER:
//  This code is provided by the OPC Foundation solely to assist in 
//  understanding and use of the appropriate OPC Specification(s) and may be 
//  used as set forth in the License Grant section of the OPC Specification.
//  This code is provided as-is and without warranty or support of any sort
//  and is subject to the Warranty and Liability Disclaimers which appear
//  in the printed OPC Specification.
//
// MODIFICATION LOG:
//
// Date       By    Notes
// ---------- ---   -----
// 2003/10/27 RSA   Added from code donated by Siemens AG.

#include "stdafx.h"
#include "OPCTrace.h"
#include "OPCUtils.h"

//==============================================================================
// Local Declarations

#define REG_CONFIG_PATH _T("SOFTWARE\\OPC Foundation\\")
#define REG_MAX_ENTRIES _T("MaxEntries")
#define REG_TRACE_LEVEL _T("TraceLevel")

#define DEFAULT_MAX_ENTRIES 10000
#define DEFAULT_TRACE_LEVEL 0x00000000
#define MAX_MSG_LENGTH      32768

static const LPCTSTR TRACE_ENTRY        = _T("<tr class='%s'><td width='65'>%08d</td><td width='85'>%02d:%02d:%02d.%03d</td><td width='40'>%04X</td><td width='15'>%s</td><td width='*'>%s</td></tr>\n");
static const LPCTSTR TRACE_ENTRY_ERROR  = _T("<tr class='ERROR'><td width='65'>%08d</td><td width='85'>%02d:%02d:%02d.%03d</td><td width='40'>%04X</td><td width='15'>FE</td><td width='*'>Trace exception (invalid param?)</td></tr>\n");
static const LPCTSTR TRACE_BACKUP_ERROR = _T("<tr class='ERROR'><td width='65'>%08d</td><td width='85'>%02d:%02d:%02d.%03d</td><td width='40'>%04X</td><td width='15'>FE</td><td width='*'>Could not backup log file '%s'.</td></tr>\n");

static COPCTrace g_cTrace;

//==============================================================================
// OpcGetTrace

COPCTrace& GetTrace() 
{
	return g_cTrace;
}

//==============================================================================
// COPCTrace

// Constructor
COPCTrace::COPCTrace()
{
	m_pExeInfo = NULL;
	m_pDllInfo = NULL;
	m_szRegKey = NULL;
	m_szFilePath = NULL;
	m_hKey = NULL;
	m_hEvent = NULL;
	m_dwMaxEntries = 0;
	m_dwTraceLevel = 0;
	m_dwLineNo = 0;
}

// Destructor
COPCTrace::~COPCTrace()
{
}

// DeleteRegistryKeys
void COPCTrace::DeleteRegistryKeys()
{	
	// create the critical section.
	InitializeCriticalSection(&m_csLock);

	COPCLock cLock(m_csLock);

	// build the registry key path.
	LPTSTR szRegKey = (LPTSTR)CoTaskMemAlloc(sizeof(TCHAR)*(MAX_PATH+1));
	memset(szRegKey, 0, sizeof(TCHAR)*(MAX_PATH+1));
    
	_tcscpy(szRegKey, REG_CONFIG_PATH);
	_tcscat(szRegKey, m_pDllInfo->Name);

	// delete the registry key.
	OpcRegDeleteKey(HKEY_LOCAL_MACHINE, szRegKey);

	// free the string.
	CoTaskMemFree(szRegKey);
}

// Initialize
void COPCTrace::Initialize(HINSTANCE hModule)
{	
	// create the critical section.
	InitializeCriticalSection(&m_csLock);

	COPCLock cLock(m_csLock);

	m_pExeInfo = GetVersionInfo(NULL);
	m_pDllInfo = GetVersionInfo(hModule);

	// build the registry key path.
	m_szRegKey = (LPTSTR)CoTaskMemAlloc(sizeof(TCHAR)*(MAX_PATH+1));
	memset(m_szRegKey, 0, sizeof(TCHAR)*(MAX_PATH+1));
    
	_tcscpy(m_szRegKey, REG_CONFIG_PATH);
	_tcscat(m_szRegKey, m_pDllInfo->Name);
	_tcscat(m_szRegKey, _T("\\Trace\\"));
	_tcscat(m_szRegKey, m_pExeInfo->Name);
	
	// Read the configuration information from the registry.
	ReadRegistryConfig();

	// Create an event used to check for changes to the registry.
	CreateRegistryEvent();

	// build the log file path.
	m_szFilePath = (LPTSTR)CoTaskMemAlloc(sizeof(TCHAR)*(MAX_PATH+1));
	memset(m_szFilePath, 0, sizeof(TCHAR)*(MAX_PATH+1));
  
	_tcscat(m_szFilePath, _T(".\\"));
	_tcscat(m_szFilePath, m_pExeInfo->Name);
	_tcscat(m_szFilePath, _T("_"));
	_tcscat(m_szFilePath, m_pDllInfo->Name);
	_tcscat(m_szFilePath, _T("_LOG.htm"));

	// initialize line count.
	m_dwLineNo = 0;

	// write header to new log file.
	WriteHeader();
}

// Uninitialize
void COPCTrace::Uninitialize()
{
	COPCLock cLock(m_csLock);

	// free all resources.
	delete m_pExeInfo;
	delete m_pDllInfo;
	CoTaskMemFree(m_szRegKey);
	CoTaskMemFree(m_szFilePath);
	RegCloseKey(m_hKey);
	CloseHandle(m_hEvent);

	// initialize members.
	m_pExeInfo = NULL;
	m_pDllInfo = NULL;
	m_szRegKey = NULL;
	m_szFilePath = NULL;
	m_hKey = NULL;
	m_hEvent = NULL;
	m_dwMaxEntries = 0;
	m_dwTraceLevel = 0;
	m_dwLineNo = 0;

	cLock.Unlock();

	// destroy the critical section.
	DeleteCriticalSection(&m_csLock);
}

// GetTraceLevel
DWORD COPCTrace::GetTraceLevel()
{
	// check if the trace has already been uninitialized.
	if (m_pExeInfo == NULL)
	{
		return 0;
	}

	COPCLock cLock(m_csLock);

	if (m_hEvent != NULL)
	{
		// check for change to registry.
		if (WaitForSingleObject(m_hEvent, 0) == WAIT_OBJECT_0)
		{
			// re-read configuration.
			ReadRegistryConfig();

			// re-create registry change event.
			CreateRegistryEvent();
		}
	}
	
	return m_dwTraceLevel;
}

// Write
void COPCTrace::Write(LPCTSTR szStyle, LPCTSTR szSymbol, LPCTSTR szFormat, va_list arg_ptr)
{
	COPCLock cLock(m_csLock);

	// do nothing if trace level is zero.
	if (m_dwTraceLevel == 0)
	{
		return;
	}
	
	// get current time.
	SYSTEMTIME st;
	GetSystemTime(&st);

	// declare buffers for the entry in the trace log.
	TCHAR szText[MAX_MSG_LENGTH];
	TCHAR szBuffer[MAX_MSG_LENGTH];

	memset(szText, 0, sizeof(TCHAR)*MAX_MSG_LENGTH);
	memset(szBuffer, 0, sizeof(TCHAR)*MAX_MSG_LENGTH);

	// backup log file.
	if (m_dwMaxEntries <= m_dwLineNo)
	{
		// create backup file name.
		LPTSTR szBackupFile = OpcDupStr(m_szFilePath);
		szBackupFile[_tcslen(m_szFilePath)-1] = _T('_');

		// removing existing backup file.
		int iResult = _tremove(szBackupFile);

		// rename current log file.
		if (iResult == 0 || errno == ENOENT)
		{
			iResult = _trename(m_szFilePath, szBackupFile);
		}

		// reset line count.
		m_dwLineNo  = 0;

		// write header to new log file.
		WriteHeader();

		// write error message if backup failed.
		if (iResult != 0)
		{
			FILE* fp = _tfopen(m_szFilePath, _T("a"));
			_ftprintf(fp, TRACE_BACKUP_ERROR, ++m_dwLineNo, st.wHour, st.wMinute, st.wSecond, st.wMilliseconds, GetCurrentThreadId(), szBackupFile);
			fclose(fp);
		}
	}

	// increment line count.
	m_dwLineNo++;
   
	// format and write message.
    try
    {
		_vstprintf(szText, szFormat, arg_ptr);
		_stprintf(szBuffer, TRACE_ENTRY, szStyle, m_dwLineNo, st.wHour, st.wMinute, st.wSecond, st.wMilliseconds, GetCurrentThreadId(), szSymbol, szText);
    }
    catch( ... )
    {
		_stprintf(szBuffer, TRACE_ENTRY_ERROR, m_dwLineNo, st.wHour, st.wMinute, st.wSecond, st.wMilliseconds, GetCurrentThreadId());
    }

	// write message to file.
	FILE* fp = _tfopen(m_szFilePath, _T("a"));
	_ftprintf(fp, szBuffer);
	fclose(fp);
}

// WriteError
void COPCTrace::WriteError(LPCTSTR szFormat, ...)
{
	va_list args;
	va_start(args, szFormat);

	Write(_T("ERROR"), _T("*"), szFormat, args);

	va_end(args);
}

// WriteErrorInfo
void COPCTrace::WriteErrorInfo(LPCTSTR szFormat, ...)
{
	va_list args;
	va_start(args, szFormat);

	Write(_T("ERRORINFO"), _T("*"), szFormat, args);

	va_end(args);
}

// WriteSpecial
void COPCTrace::WriteSpecial(LPCTSTR szFormat, ...)
{
	va_list args;
	va_start(args, szFormat);

	Write(_T("SPECIAL"), _T("!"), szFormat, args);

	va_end(args);
}

// WriteSpecialInfo
void COPCTrace::WriteSpecialInfo(LPCTSTR szFormat, ...)
{
	va_list args;
	va_start(args, szFormat);

	Write(_T("SPECIALINFO"), _T("!"), szFormat, args);

	va_end(args);
}

// WriteIn
void COPCTrace::WriteIn(LPCTSTR szFormat, ...)
{
	va_list args;
	va_start(args, szFormat);

	Write(_T("IN"), _T("&rarr;"), szFormat, args);

	va_end(args);
}

// WriteOut
void COPCTrace::WriteOut(LPCTSTR szFormat, ...)
{
	va_list args;
	va_start(args, szFormat);

	Write(_T("OUT"), _T("&larr;"), szFormat, args);

	va_end(args);
}

// WriteInvoke
void COPCTrace::WriteInvoke(LPCTSTR szFormat, ...)
{
	va_list args;
	va_start(args, szFormat);

	Write(_T("INVOKE"), _T(">"), szFormat, args);

	va_end(args);
}

// WriteInvokeDone
void COPCTrace::WriteInvokeDone(LPCTSTR szFormat, ...)
{
	va_list args;
	va_start(args, szFormat);

	Write(_T("INVOKEDONE"), _T("<"), szFormat, args);

	va_end(args);
}

// WriteValue
void COPCTrace::WriteValue(LPCTSTR szFormat, ...)
{
	va_list args;
	va_start(args, szFormat);

	Write(_T("VALUE"), _T("="), szFormat, args);

	va_end(args);
}

// WriteValueInfo
void COPCTrace::WriteValueInfo(LPCTSTR szFormat, ...)
{
	va_list args;
	va_start(args, szFormat);

	Write(_T("VALUEINFO"), _T("="), szFormat, args);

	va_end(args);
}

// WriteMisc
void COPCTrace::WriteMisc(LPCTSTR szFormat, ...)
{
	va_list args;
	va_start(args, szFormat);

	Write(_T("MISC"), _T("-"), szFormat, args);

	va_end(args);
}

// WriteMiscInfo
void COPCTrace::WriteMiscInfo(LPCTSTR szFormat, ...)
{
	va_list args;
	va_start(args, szFormat);

	Write(_T("MISCINFO"), _T("-"), szFormat, args);

	va_end(args);
}

// ReadRegistryConfig
void COPCTrace::ReadRegistryConfig()
{
	// get the trace level.
	if (!OpcRegGetValue(HKEY_LOCAL_MACHINE, m_szRegKey, REG_TRACE_LEVEL, &m_dwTraceLevel))
	{
		m_dwTraceLevel = DEFAULT_TRACE_LEVEL;
		OpcRegSetValue(HKEY_LOCAL_MACHINE, m_szRegKey, REG_TRACE_LEVEL, m_dwTraceLevel);
	}

	// get the maximum number of log entries.
	if (!OpcRegGetValue(HKEY_LOCAL_MACHINE, m_szRegKey, REG_MAX_ENTRIES, &m_dwMaxEntries))
	{
		m_dwMaxEntries = DEFAULT_MAX_ENTRIES;
		OpcRegSetValue(HKEY_LOCAL_MACHINE, m_szRegKey, REG_MAX_ENTRIES, m_dwMaxEntries);
	}
}

// CreateRegistryEvent
void COPCTrace::CreateRegistryEvent()
{
	COPCLock cLock(m_csLock);
	
	LONG lResult = ERROR_SUCCESS;

	if (m_hKey == NULL)
	{
		// open the key to watch for changes.
		lResult = RegOpenKeyEx(HKEY_LOCAL_MACHINE, m_szRegKey, NULL, KEY_ALL_ACCESS, &m_hKey);

		if (lResult != ERROR_SUCCESS)
		{
			return;
		}

		// create a event to detect registry changes.
		if (m_hEvent == NULL)
		{
			m_hEvent = CreateEvent(NULL, TRUE, FALSE, NULL);

			if (m_hEvent == NULL)
			{
				RegCloseKey(m_hKey);
				m_hKey = NULL;
				return;
			}
		}
	}

	// reset the registry change event.
	ResetEvent(m_hEvent);

	// register interest in change notifications.
	lResult = RegNotifyChangeKeyValue(
		m_hKey,				        // handle to key to watch
		TRUE,					    // subkey notification option
		REG_NOTIFY_CHANGE_LAST_SET,	// notify value changes
		m_hEvent,			        // handle to event to be signaled
		TRUE                        // asynchronous reporting option
	);		

	if (lResult != ERROR_SUCCESS)
	{
		RegCloseKey(m_hKey);
		CloseHandle(m_hEvent);

		m_hKey = NULL;
		m_hEvent = NULL;
	}				
}

// WriteHeader
void COPCTrace::WriteHeader()
{
	// do nothing if trace level is zero.
	if (m_dwTraceLevel == 0)
	{
		return;
	}
	
	// over write existing file.
	FILE* fp = _tfopen(m_szFilePath, _T("w"));

	if (fp == NULL)
	{
		return;
	}

	_ftprintf(fp, _T("<html>\n  <head>\n"));
	_ftprintf(fp, _T("  <title>"));
	_ftprintf(fp, m_pDllInfo->Name);
	_ftprintf(fp, _T("- TRACE</title>\n"));
	_ftprintf(fp, _T("<meta name='generator' content='Trace generated' />\n"));
	_ftprintf(fp, _T("<meta name='author' content='OPC Foundation' />\n"));
	_ftprintf(fp, _T("<style>\n"));
	_ftprintf(fp, _T("	.ERROR{COLOR:red;FONT-STYLE:normal; FONT-WEIGHT:bold;BACKGROUND-COLOR:white;FONT-FAMILY:Arial;FONT-SIZE:10pt}\n"));
	_ftprintf(fp, _T("	.ERRORINFO{COLOR:red;FONT-STYLE:normal; FONT-WEIGHT:normal;BACKGROUND-COLOR:#E0E0E0;FONT-FAMILY:Arial;FONT-SIZE:10pt}\n"));
	_ftprintf(fp, _T("	.SPECIAL{COLOR:red;FONT-STYLE:normal; FONT-WEIGHT:normal;BACKGROUND-COLOR:white;FONT-FAMILY:Arial;FONT-SIZE:10pt}\n"));
	_ftprintf(fp, _T("	.SPECIALINFO{COLOR:darkblue;FONT-STYLE:normal; FONT-WEIGHT:normal;BACKGROUND-COLOR:#E0E0E0;FONT-FAMILY:Arial;FONT-SIZE:10pt}\n"));
	_ftprintf(fp, _T("	.IN{COLOR:green;FONT-STYLE:normal; FONT-WEIGHT:normal;BACKGROUND-COLOR:white;FONT-FAMILY:Arial;FONT-SIZE:10pt}\n"));
	_ftprintf(fp, _T(" 	.OUT{COLOR:green;FONT-STYLE:normal; FONT-WEIGHT:normal;BACKGROUND-COLOR:#E0E0E0;FONT-FAMILY:Arial;FONT-SIZE:10pt}\n"));
	_ftprintf(fp, _T("	.INVOKE{COLOR:magenta;FONT-STYLE:normal; FONT-WEIGHT:normal;BACKGROUND-COLOR:white;FONT-FAMILY:Arial;FONT-SIZE:10pt}\n"));
	_ftprintf(fp, _T("	.INVOKEDONE{COLOR:magenta;FONT-STYLE:normal; FONT-WEIGHT:normal;BACKGROUND-COLOR:#E0E0E0;FONT-FAMILY:Arial;FONT-SIZE:10pt}\n"));
	_ftprintf(fp, _T("	.VALUE{COLOR:blue;FONT-STYLE:normal; FONT-WEIGHT:normal;BACKGROUND-COLOR:white;FONT-FAMILY:Arial;FONT-SIZE:10pt}\n"));
	_ftprintf(fp, _T("	.VALUEINFO{COLOR:blue;FONT-STYLE:normal; FONT-WEIGHT:normal;BACKGROUND-COLOR:#E0E0E0;FONT-FAMILY:Arial;FONT-SIZE:10pt}\n"));
	_ftprintf(fp, _T("	.MISC{COLOR:darkblue;FONT-STYLE:normal; FONT-WEIGHT:normal;BACKGROUND-COLOR:white;FONT-FAMILY:Arial;FONT-SIZE:10pt}\n"));
	_ftprintf(fp, _T("	.MISCINFO{COLOR:darkblue;FONT-STYLE:normal; FONT-WEIGHT:normal;BACKGROUND-COLOR:#E0E0E0;FONT-FAMILY:Arial;FONT-SIZE:10pt}\n"));
	_ftprintf(fp, _T("	.HEADER{COLOR:black;FONT-STYLE:normal; FONT-WEIGHT:bold;BACKGROUND-COLOR:white;FONT-FAMILY:Arial;FONT-SIZE:10pt}\n"));
	_ftprintf(fp, _T("	.xERROR{COLOR:red;FONT-STYLE:normal; FONT-WEIGHT:bold;BACKGROUND-COLOR:white;FONT-FAMILY:Arial;FONT-SIZE:10pt}\n"));
	_ftprintf(fp, _T("	.xERRORINFO{COLOR:red;FONT-STYLE:normal; FONT-WEIGHT:normal;BACKGROUND-COLOR:#E0E0E0;FONT-FAMILY:Arial;FONT-SIZE:10pt}\n"));
	_ftprintf(fp, _T("	.xSPECIAL{COLOR:red;FONT-STYLE:normal; FONT-WEIGHT:normal;BACKGROUND-COLOR:white;FONT-FAMILY:Arial;FONT-SIZE:10pt}\n"));
	_ftprintf(fp, _T("	.xSPECIALINFO{COLOR:darkblue;FONT-STYLE:normal; FONT-WEIGHT:normal;BACKGROUND-COLOR:#E0E0E0;FONT-FAMILY:Arial;FONT-SIZE:10pt}\n"));
	_ftprintf(fp, _T("	.xIN{COLOR:green;FONT-STYLE:normal; FONT-WEIGHT:normal;BACKGROUND-COLOR:white;FONT-FAMILY:Arial;FONT-SIZE:10pt}\n"));
	_ftprintf(fp, _T(" 	.xOUT{COLOR:green;FONT-STYLE:normal; FONT-WEIGHT:normal;BACKGROUND-COLOR:#E0E0E0;FONT-FAMILY:Arial;FONT-SIZE:10pt}\n"));
	_ftprintf(fp, _T("	.xINVOKE{COLOR:magenta;FONT-STYLE:normal; FONT-WEIGHT:normal;BACKGROUND-COLOR:white;FONT-FAMILY:Arial;FONT-SIZE:10pt}\n"));
	_ftprintf(fp, _T("	.xINVOKEDONE{COLOR:magenta;FONT-STYLE:normal; FONT-WEIGHT:normal;BACKGROUND-COLOR:#E0E0E0;FONT-FAMILY:Arial;FONT-SIZE:10pt}\n"));
	_ftprintf(fp, _T("	.xVALUE{COLOR:blue;FONT-STYLE:normal; FONT-WEIGHT:normal;BACKGROUND-COLOR:white;FONT-FAMILY:Arial;FONT-SIZE:10pt}\n"));
	_ftprintf(fp, _T("	.xVALUEINFO{COLOR:blue;FONT-STYLE:normal; FONT-WEIGHT:normal;BACKGROUND-COLOR:#E0E0E0;FONT-FAMILY:Arial;FONT-SIZE:10pt}\n"));
	_ftprintf(fp, _T("	.xMISC{COLOR:darkblue;FONT-STYLE:normal; FONT-WEIGHT:normal;BACKGROUND-COLOR:white;FONT-FAMILY:Arial;FONT-SIZE:10pt}\n"));
	_ftprintf(fp, _T("	.xMISCINFO{COLOR:darkblue;FONT-STYLE:normal; FONT-WEIGHT:normal;BACKGROUND-COLOR:#E0E0E0;FONT-FAMILY:Arial;FONT-SIZE:10pt}\n"));
	_ftprintf(fp, _T("</style>\n"));
	_ftprintf(fp, _T("<script id='clientEventHandlersJS' language='javascript'>\n"));
	_ftprintf(fp, _T("<!--\n"));
	_ftprintf(fp, _T("  function ChangeButton(MyButton, MyStyle, ValueOn, ValueOff){if(document.styleSheets[0].rules[MyStyle].style.display=='none'){document.styleSheets[0].rules[MyStyle].style.display='';document.all[MyButton].value=ValueOn;}else{document.styleSheets[0].rules[MyStyle].style.display='none';document.all[MyButton].value=ValueOff;}}\n"));
	_ftprintf(fp, _T("  function btn0_onclick(){ChangeButton('btn0', 0, 'Hide Error', 'Show Error');}\n"));
	_ftprintf(fp, _T("  function btn1_onclick(){ChangeButton('btn1', 1, 'Hide Error Info', 'Show Error Info');}\n"));
	_ftprintf(fp, _T("  function btn2_onclick(){ChangeButton('btn2', 2, 'Hide Special', 'Show Special');}\n"));
	_ftprintf(fp, _T("  function btn3_onclick(){ChangeButton('btn3', 3, 'Hide Special Info', 'Show Special Info');}\n"));
	_ftprintf(fp, _T("  function btn4_onclick(){ChangeButton('btn4', 4, 'Hide In', 'Show In');}\n"));
	_ftprintf(fp, _T("  function btn5_onclick(){ChangeButton('btn5', 5, 'Hide Out', 'Show Out');}\n"));
	_ftprintf(fp, _T("  function btn6_onclick(){ChangeButton('btn6', 6, 'Hide Invoke', 'Show Invoke');}\n"));
	_ftprintf(fp, _T("  function btn7_onclick(){ChangeButton('btn7', 7, 'Hide Invoke Done', 'Show Invoke Done');}\n"));
	_ftprintf(fp, _T("  function btn8_onclick(){ChangeButton('btn8', 8, 'Hide Value', 'Show Value');}\n"));
	_ftprintf(fp, _T("  function btn9_onclick(){ChangeButton('btn9', 9, 'Hide Value Info', 'Show Value Info');}\n"));
	_ftprintf(fp, _T("  function btn10_onclick(){ChangeButton('btn10', 10, 'Hide Misc', 'Show Misc');}\n"));
	_ftprintf(fp, _T("  function btn11_onclick(){ChangeButton('btn11', 11, 'Hide Misc Info', 'Show Misc Info');}\n"));
	_ftprintf(fp, _T("//-->\n"));
	_ftprintf(fp, _T("</script>\n"));
	_ftprintf(fp, _T("</head>\n"));
	_ftprintf(fp, _T("<body>\n"));
	_ftprintf(fp, _T("<h1 style='font-family:Arial' style='font-size:16pt'>Traceoutput</h1>\n"));
	_ftprintf(fp, _T("<table style='font-family:Arial' style='font-size:10pt' bgcolor='white' align='justify' cellspacing='1' width='100%%' border='0' borderColorDark='white' borderColorLight='white' background='' borderColor='white'>\n"));
	_ftprintf(fp, _T("  <tr><td bgcolor='#e0e0e0' width='100'>EXE-File</td><td width='*'>"));
	_ftprintf(fp, m_pExeInfo->OriginalFileName);
	_ftprintf(fp, _T("</td></tr>\n"));
	_ftprintf(fp, _T("  <tr><td bgcolor='#e0e0e0' width='100'>EXE-Build</td><td width='*'>"));
	_ftprintf(fp, m_pExeInfo->BuildInfo);
	_ftprintf(fp, _T("</td></tr>\n"));
	_ftprintf(fp, _T("  <tr><td bgcolor='#e0e0e0' width='100'>EXE-Version</td><td width='*'>"));
	_ftprintf(fp, m_pExeInfo->FileVersion);
	_ftprintf(fp, _T("</td></tr>\n"));
	_ftprintf(fp, _T("  <tr><td bgcolor='#e0e0e0' width='100'>EXE-Path</td><td width='*'>"));
	_ftprintf(fp, m_pExeInfo->Path);
	_ftprintf(fp, _T("</td></tr>\n"));
	_ftprintf(fp, _T("  <tr><td bgcolor='#e0e0e0' width='100'>EXE-Comment</td><td width='*'>"));
	_ftprintf(fp, m_pExeInfo->Comments);
	_ftprintf(fp, _T("</td></tr>\n"));
	_ftprintf(fp, _T("  <tr><td bgcolor='#e0e0e0' width='100'>DLL-Start</td><td width='*'>"));
	_ftprintf(fp, m_pDllInfo->OriginalFileName);
	_ftprintf(fp, _T("</td></tr>\n"));
	_ftprintf(fp, _T("  <tr><td bgcolor='#e0e0e0' width='100'>DLL-Build</td><td width='*'>"));
	_ftprintf(fp, m_pDllInfo->BuildInfo);
	_ftprintf(fp, _T("</td></tr>\n"));
	_ftprintf(fp, _T("  <tr><td bgcolor='#e0e0e0' width='100'>DLL-Version</td><td width='*'>"));
	_ftprintf(fp, m_pDllInfo->FileVersion);
	_ftprintf(fp, _T("</td></tr>\n"));
	_ftprintf(fp, _T("  <tr><td bgcolor='#e0e0e0' width='100'>DLL-Path</td><td width='*'>"));
	_ftprintf(fp, m_pDllInfo->Path);
	_ftprintf(fp, _T("</td></tr>\n"));
	_ftprintf(fp, _T("  <tr><td bgcolor='#e0e0e0' width='100'>DLL-Comment</td><td width='*'>"));
	_ftprintf(fp, m_pDllInfo->Comments);
	_ftprintf(fp, _T("</td></tr>\n"));
	_ftprintf(fp, _T("  <tr><td bgcolor='#e0e0e0' width='100'>DLL-Trace</td><td width='*'>TraceLevel = %u"), m_dwTraceLevel);
	_ftprintf(fp, _T("</td></tr>\n"));
	_ftprintf(fp, _T("</table><hr />\n"));
	_ftprintf(fp, _T("<table style='font-family:Arial' style='FONT-SIZE:10pt' border='0' align='justify' cellspacing='1' width='576'><tr>\n"));
	
	_ftprintf(fp, _T("<td width='115'><BUTTON id='btn0' class='xERROR' onclick='btn0_onclick()' style='WIDTH: 115px'>Hide Error</BUTTON></td>\n"));
	_ftprintf(fp, _T("<td width='115'><BUTTON id='btn2' class='xSPECIAL' onclick='btn2_onclick()' style='WIDTH: 115px'>Hide Special</BUTTON></td>\n"));
	_ftprintf(fp, _T("<td width='115'><BUTTON id='btn4' class='xIN' onclick='btn4_onclick()' style='WIDTH: 115px'>Hide In</BUTTON></td>\n"));
	_ftprintf(fp, _T("<td width='115'><BUTTON id='btn6' class='xINVOKE' onclick='btn6_onclick()' style='WIDTH: 115px'>Hide Invoke</BUTTON></td>\n"));
	_ftprintf(fp, _T("<td width='115'><BUTTON id='btn8' class='xVALUE' onclick='btn8_onclick()' style='WIDTH: 115px'>Hide Value</BUTTON></td>\n"));
	_ftprintf(fp, _T("<td width='115'><BUTTON id='btn10' class='xMISC' onclick='btn10_onclick()' style='WIDTH: 115px'>Hide Misc</BUTTON></td>\n"));

	_ftprintf(fp, _T("</tr></table><table style='font-family:Arial' style='font-size:10pt' border='0' align='justify' cellspacing='1' width='576'><tr>\n"));

	_ftprintf(fp, _T("<td width='115'><BUTTON id='btn1' class='xERRORINFO' onclick='btn1_onclick()' style='WIDTH: 115px'>Hide Error Info</BUTTON></td>\n"));
	_ftprintf(fp, _T("<td width='115'><BUTTON id='btn3' class='xSPECIALINFO' onclick='btn3_onclick()' style='WIDTH: 115px'>Hide Special Info</BUTTON></td>\n"));
	_ftprintf(fp, _T("<td width='115'><BUTTON id='btn5' class='xOUT' onclick='btn5_onclick()' style='WIDTH: 115px'>Hide Out</BUTTON></td>\n"));
	_ftprintf(fp, _T("<td width='115'><BUTTON id='btn7' class='xINVOKEDONE' onclick='btn7_onclick()' style='WIDTH: 115px'>Hide Invoke Done</BUTTON></td>\n"));
	_ftprintf(fp, _T("<td width='115'><BUTTON id='btn9' class='xVALUEINFO' onclick='btn9_onclick()' style='WIDTH: 115px'>Hide Value Info</BUTTON></td>\n"));
	_ftprintf(fp, _T("<td width='115'><BUTTON id='btn11' class='xMISCINFO' onclick='btn11_onclick()' style='WIDTH: 115px'>Hide Misc Info</BUTTON></td>\n"));
	_ftprintf(fp, _T("</tr></table><hr/><table cols='4' style='font-family:Arial' style='font-size:12pt' border='0' align='justify' cellspacing='1' width='100%%'>\n"));
	_ftprintf(fp, _T("<tr class='HEADER'><td width='65'>Line No</td><td width='85'>Time</td><td width='40'>TID</td><td width='*'>Description</td></tr>\n"));
	_ftprintf(fp, _T("</table><table cols='5' style='font-family:Arial' style='font-size:10pt' border='0' align='justify' cellspacing='1' width='100%%'>\n"));

	fflush(fp);
	fclose(fp);
}

// GetVersionInfo
COPCVersionInfo* COPCTrace::GetVersionInfo(HINSTANCE hModule)
{
	// retrieve location of exe
	TCHAR szModulePath[_MAX_PATH+1];
	memset(szModulePath, 0, sizeof(TCHAR)*(_MAX_PATH+1));

	DWORD dwResult = GetModuleFileName(hModule, szModulePath, _MAX_PATH);

	if (dwResult == 0)
	{
		return NULL;
	}
	
	// create version info structure.
	COPCVersionInfo* pVersionInfo = new COPCVersionInfo();
	pVersionInfo->Path = OpcDupStr(szModulePath);

	// get the file name.
	TCHAR szFileName[_MAX_PATH+1];
	memset(szFileName, 0, sizeof(TCHAR)*(_MAX_PATH+1));

	_tsplitpath(szModulePath, NULL, NULL, szFileName, NULL);
	pVersionInfo->Name = OpcDupStr(szFileName);
	
	// get the size of the version info resource block.
	DWORD dwHandle = 0;

	DWORD dwSize = GetFileVersionInfoSize(szModulePath, &dwHandle);
	
	if (dwSize == 0)
	{
		return pVersionInfo;
	}
	
	// get the version info resource block.
	BYTE* pInfo = new BYTE[dwSize];
	memset(pInfo, 0, dwSize);

	if (!GetFileVersionInfo(szModulePath, dwHandle, dwSize, (LPVOID)pInfo))
	{
		delete pInfo;
		return pVersionInfo;
	}

	UINT uBytesRead = 0;

	// get the common version information.
	VS_FIXEDFILEINFO* pFixedFileInfo = NULL;

	if (VerQueryValue(pInfo, _T("\\"), (LPVOID*)&pFixedFileInfo, &uBytesRead))
	{
		switch (pFixedFileInfo->dwFileFlags)
		{
			case VS_FF_DEBUG:        { pVersionInfo->BuildInfo = OpcDupStr(_T("Debug"));         break; }
			case VS_FF_PRERELEASE:   { pVersionInfo->BuildInfo = OpcDupStr(_T("Pre-Release"));   break; }
			case VS_FF_PATCHED:      { pVersionInfo->BuildInfo = OpcDupStr(_T("Patched"));       break; }
			case VS_FF_PRIVATEBUILD: { pVersionInfo->BuildInfo = OpcDupStr(_T("Private Build")); break; }
			case VS_FF_SPECIALBUILD: { pVersionInfo->BuildInfo = OpcDupStr(_T("Special Build")); break; }
			case VS_FF_INFOINFERRED: { pVersionInfo->BuildInfo = OpcDupStr(_T("Unknown"));       break; }
			default:                 { pVersionInfo->BuildInfo = OpcDupStr(_T("Release"));       break; }
		}
	}
		
	// declare pointer to structure used to store enumerated languages and code pages.
	struct LANGANDCODEPAGE { WORD wLanguage; WORD wCodePage; }* pTranslation = NULL;

	// read the list of languages and code pages.
	if (VerQueryValue(pInfo, _T("\\VarFileInfo\\Translation"), (LPVOID*)&pTranslation, &uBytesRead))
	{
		TCHAR szPrefix[_MAX_PATH+1];
		TCHAR szSubBlock[_MAX_PATH+1];

		memset(szPrefix, 0, sizeof(TCHAR)*(_MAX_PATH+1));
		memset(szSubBlock, 0, sizeof(TCHAR)*(_MAX_PATH+1));

		_sntprintf(szPrefix, MAX_PATH, _T("\\StringFileInfo\\%04x%04x"), pTranslation->wLanguage, pTranslation->wCodePage);

		TCHAR* pStringInfo = NULL;

		// file version.
		_sntprintf(szSubBlock, MAX_PATH, _T("%s\\FileVersion"), szPrefix);

		if (VerQueryValue(pInfo, szSubBlock, (LPVOID*)&pStringInfo, &uBytesRead))
		{
			pVersionInfo->FileVersion = OpcDupStr(pStringInfo);
		}

		// file description.
		_sntprintf(szSubBlock, MAX_PATH, _T("%s\\FileDescription"), szPrefix);

		if (VerQueryValue(pInfo, szSubBlock, (LPVOID*)&pStringInfo, &uBytesRead))
		{
			pVersionInfo->FileDecription = OpcDupStr(pStringInfo);
		}

		// original file name.
		_sntprintf(szSubBlock, MAX_PATH, _T("%s\\OriginalFilename"), szPrefix);

		if (VerQueryValue(pInfo, szSubBlock, (LPVOID*)&pStringInfo, &uBytesRead))
		{
			pVersionInfo->OriginalFileName = OpcDupStr(pStringInfo);
		}

		// comments.
		_sntprintf(szSubBlock, MAX_PATH, _T("%s\\Comments"), szPrefix);

		if (VerQueryValue(pInfo, szSubBlock, (LPVOID*)&pStringInfo, &uBytesRead))
		{
			pVersionInfo->Comments = OpcDupStr(pStringInfo);
		}
	}
	
	// ensure all fields have a valid value.
	if (pVersionInfo->BuildInfo == NULL)        pVersionInfo->BuildInfo = OpcDupStr(_T(""));
	if (pVersionInfo->FileVersion == NULL)      pVersionInfo->FileVersion = OpcDupStr(_T(""));
	if (pVersionInfo->FileDecription == NULL)   pVersionInfo->FileDecription = OpcDupStr(_T(""));
	if (pVersionInfo->OriginalFileName == NULL) pVersionInfo->OriginalFileName = OpcDupStr(_T(""));
	if (pVersionInfo->Comments == NULL)         pVersionInfo->Comments = OpcDupStr(_T(""));

	delete pInfo;
	return pVersionInfo;
}


