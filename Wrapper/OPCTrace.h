//============================================================================
// TITLE: OPCTrace.h
//
// CONTENTS:
// 
// A class that provides support for debugging output to a file.
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

#ifndef __OPCTrace_H__
#define __OPCTrace_H__

//============================================================================
// MACRO:   TRACE_MASK_XXX
// PURPOSE: The masks used to control what is written to the log file.

#define TRACE_MASK_ERROR	    0x0001  // any error messages.
#define TRACE_MASK_ERROR_INFO   0x0002  // additional detail about errors conditions.
#define TRACE_MASK_SPECIAL      0x0004  // any warning/informational messages.
#define TRACE_MASK_SPECIAL_INFO 0x0008  // additional detail about warning/informational conditions.
#define TRACE_MASK_INOUT  	    0x0010  // entry/exit from public functions.
#define TRACE_MASK_INVOKE       0x0020	// before/after remote server calls.
#define TRACE_MASK_VALUE 	    0x0040  // in/out function parameters.
#define TRACE_MASK_VALUE_INFO   0x0080  // additional detail about in/out function parameters (e.g. array contents).
#define TRACE_MASK_MISC			0x0100	// any other information.
#define TRACE_MASK_MISC_INFO	0x0200	// additional detail about any other information.

struct COPCVersionInfo;

//============================================================================
// CLASS:   COPCTrace
// PURPOSE: Provides methods to write messages to a logfile.

class COPCTrace
{
public:

	//============================================================================
	// Public Operators

	// Constructor
	COPCTrace();

	// Destructor
	virtual ~COPCTrace();

	//============================================================================
	// Public Methods

	// initializes tracing by reading the configuration and creating the log files.
	void Initialize(HINSTANCE hModule);

	// stops tracing and releases all resources allocated.
	void Uninitialize();

	// ensures all registry keys created by the DLL are deleted.
	void DeleteRegistryKeys();

	// returns the current trace level.
	DWORD GetTraceLevel();

	// writes a message to the log file.
	void Write(LPCTSTR szStyle, LPCTSTR szSymbol, LPCTSTR szFormat, va_list arg_ptr);
	
	// convenience functions for writing special types of messages.
	void WriteError(LPCTSTR szFormat, ...);
	void WriteErrorInfo(LPCTSTR szFormat, ...);
	void WriteSpecial(LPCTSTR szFormat, ...);
	void WriteSpecialInfo(LPCTSTR szFormat, ...);
	void WriteIn(LPCTSTR szFormat, ...);
	void WriteOut(LPCTSTR szFormat, ...);
	void WriteInvoke(LPCTSTR szFormat, ...);
	void WriteInvokeDone(LPCTSTR szFormat, ...);
	void WriteValue(LPCTSTR szFormat, ...);
	void WriteValueInfo(LPCTSTR szFormat, ...);
	void WriteMisc(LPCTSTR szFormat, ...);
	void WriteMiscInfo(LPCTSTR szFormat, ...);

private:

	//============================================================================
	// Private Methods

	// reads the configuration information from the registry.
	void ReadRegistryConfig();

	// creates an event used to check for changes to the registry.
	void CreateRegistryEvent();

	// populates a version info structure for a module.
	COPCVersionInfo* GetVersionInfo(HINSTANCE hModule);

	// writes the header to the log file.
	void WriteHeader();

	//============================================================================
	// Private Members
	
	COPCVersionInfo* m_pExeInfo;
	COPCVersionInfo* m_pDllInfo;

	LPTSTR m_szRegKey;
	LPTSTR m_szFilePath;
			
	HKEY   m_hKey;
	HANDLE m_hEvent;

	DWORD  m_dwMaxEntries;
	DWORD  m_dwTraceLevel;
	DWORD  m_dwLineNo;

	CRITICAL_SECTION m_csLock;
};

//============================================================================
// FUNCTION: GetTrace
// PURPOSE:  Returns a reference to the module trace object.

COPCTrace& GetTrace();

//============================================================================
// MACRO:   TRACE_XXX
// PURPOSE: Macros used to write events to the trace logs.

#define TRACE_ERROR         if ((GetTrace().GetTraceLevel() & TRACE_MASK_ERROR) != 0) GetTrace().WriteError
#define TRACE_ERRORINFO     if ((GetTrace().GetTraceLevel() & TRACE_MASK_ERROR_INFO) != 0) GetTrace().WriteErrorInfo
#define TRACE_SPECIAL       if ((GetTrace().GetTraceLevel() & TRACE_MASK_SPECIAL) != 0) GetTrace().WriteSpecial
#define TRACE_SPECIALINFO   if ((GetTrace().GetTraceLevel() & TRACE_MASK_SPECIAL_INFO) != 0) GetTrace().WriteErrorSpecialInfo
#define TRACE_IN            if ((GetTrace().GetTraceLevel() & TRACE_MASK_INOUT) != 0) GetTrace().WriteIn
#define TRACE_OUT           if ((GetTrace().GetTraceLevel() & TRACE_MASK_INOUT) != 0) GetTrace().WriteOut
#define TRACE_INVOKE		if ((GetTrace().GetTraceLevel() & TRACE_MASK_INVOKE) != 0) GetTrace().WriteInvoke
#define TRACE_INVOKEDONE	if ((GetTrace().GetTraceLevel() & TRACE_MASK_INVOKE) != 0) GetTrace().WriteInvokeDone
#define TRACE_VALUE         if ((GetTrace().GetTraceLevel() & TRACE_MASK_VALUE) != 0) GetTrace().WriteValue
#define TRACE_VALUE_INFO    if ((GetTrace().GetTraceLevel() & TRACE_MASK_VALUE_INFO) != 0) GetTrace().WriteValueInfo
#define TRACE_MISC          if ((GetTrace().GetTraceLevel() & TRACE_MASK_MISC) != 0) GetTrace().WriteMisc
#define TRACE_MISC_INFO     if ((GetTrace().GetTraceLevel() & TRACE_MASK_MISC_INFO) != 0) GetTrace().WriteMiscInfo

//============================================================================
// MACROS:  TRACE_XXX
// PURPOSE: Provide macros for common events.

#define TRACE_FUNCTION_NAME(xName)  static const LPCTSTR _szFunction = _T(xName);
#define TRACE_FUNCTION              _szFunction
#define TRACE_FUNCTION_ENTER(xName) TRACE_FUNCTION_NAME(xName); TRACE_IN(_T(xName));
#define TRACE_FUNCTION_LEAVE()      TRACE_OUT(_T("%s"), _szFunction);
#define TRACE_METHOD_CALL(x)        TRACE_INVOKE(_T("%s, Calling %s"), _szFunction, _T(#x));
#define TRACE_METHOD_ERROR(x,y)     TraceMethodError(_szFunction, _T(#x), y);
#define TRACE_INTERFACE_ERROR(x)    TraceInterfaceError(_szFunction, _T(#x));

//============================================================================
// CLASS:   COPCLock
// PURPOSE: A convenience class used to enter/leave a critical section.
 
class COPCLock
{
public:

	// Constructor
	COPCLock(CRITICAL_SECTION& csLock)
	{ 
		m_dwCount = 0;
		m_pLock = &csLock; 
		
		Lock(); 
	}

	// Destructor
	~COPCLock() 
	{ 
		while (m_dwCount > 0) Unlock();
	}

	// Lock
	void Lock() 
	{ 
		EnterCriticalSection(m_pLock); 
		m_dwCount++;
	}

	// Unlock
	void Unlock() 
	{ 
		if (m_dwCount > 0) 
		{
			LeaveCriticalSection(m_pLock);
			m_dwCount--;
		}
	}

private:

	DWORD             m_dwCount;
	CRITICAL_SECTION* m_pLock;
};

//============================================================================
// STRUCTURE: COPCVersionInfo
// PURPOSE:   Stores information extracted from the version resource block.

struct COPCVersionInfo
{
	LPTSTR Name;
	LPTSTR Path;
	LPTSTR BuildInfo;
	LPTSTR FileVersion;
	LPTSTR FileDecription;
	LPTSTR OriginalFileName;
	LPTSTR Comments;

	// Constructor
	COPCVersionInfo()
	{
		Name = NULL;
		Path = NULL;
		BuildInfo = NULL;
		FileVersion = NULL;
		FileDecription = NULL;
		OriginalFileName = NULL;
		Comments = NULL;
	}

	// Destructor
	~COPCVersionInfo()
	{
		CoTaskMemFree(Name);
		CoTaskMemFree(Path);
		CoTaskMemFree(BuildInfo);
		CoTaskMemFree(FileVersion);
		CoTaskMemFree(FileDecription);
		CoTaskMemFree(OriginalFileName);
		CoTaskMemFree(Comments);
	}
};

#endif // __OPCTrace_H__
