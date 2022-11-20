//============================================================================
// TITLE: OPCUtils.h
//
// CONTENTS:
// 
// The declaration various utility functions.
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
// 2003/10/27 RSA   Initial implementation.

#ifndef _OPCUtils_H__
#define _OPCUtils_H__

//============================================================================
// FUNCTION: OpcDupStr
// PURPOSE:  Duplicates a string (free with CoTaskMemFree).

extern LPSTR OpcDupStr(LPCSTR szSrc);
extern LPWSTR OpcDupStr(LPCWSTR szSrc);

//============================================================================
// FUNCTION: OpcFind
// PURPOSE:  Returns the index of the first occurrance of the target string in the source.

extern int OpcFind(LPCTSTR szSrc, LPCTSTR szTarget);

//============================================================================
// FUNCTION: OpcFind
// PURPOSE:  Returns the index of the last occurrance of the target string in the source.

extern int OpcReverseFind(LPCTSTR szSrc, LPCTSTR szTarget);

//==============================================================================
// FUNCTION: OpcSubStr
// PURPOSE:  Returns a sub-string of a string (free with CoTaskMemFree).

extern LPTSTR OpcSubStr(LPCTSTR tsSrc, UINT uStart, UINT uCount = -1);

//==============================================================================
// FUNCTION: OpcRegGetValue
// PURPOSE:  Looks up a DWORD value in the registry.

extern bool OpcRegGetValue(
	HKEY    hBaseKey,
	LPCTSTR szSubKey,
	LPCTSTR szValueName,
	DWORD*  pdwValue
);

//==============================================================================
// FUNCTION: OpcRegSetValue
// PURPOSE:  Sets a DWORD value in the registry.

extern bool OpcRegSetValue(
	HKEY    hBaseKey,
	LPCTSTR szSubKey,
	LPCTSTR szValueName,
	DWORD   dwValue
);

//==============================================================================
// FUNCTION: OpcRegDeleteKey
// PURPOSE:  Deletes a registry key and all of its subkeys.

extern bool OpcRegDeleteKey(HKEY hBaseKey, LPCTSTR tsSubKey);

//============================================================================
// FUNCTION: GetUtcTime
// PURPOSE:  Converts a UTC FILETIME value to a GetUtcTime VARIANT DATE value.

extern DATE GetUtcTime(const FILETIME& ftUtcTime);

//============================================================================
// FUNCTION: GetMinTime
// PURPOSE:  Returns the minimum FILETIME value (e.g. 1601/01/01)

extern FILETIME GetMinTime();

//============================================================================
// FUNCTION: VariantToAutomation
// PURPOSE:  Converts a VARIANT type to an Automation compatible type. 

extern HRESULT VariantToAutomation(VARIANT* varValue);

//============================================================================
// FUNCTION: EnumVARIANT
// PURPOSE:  A class to enumerate a list of VARIANTS.

typedef CComObject<CComEnum<IEnumVARIANT, &IID_IEnumVARIANT, VARIANT, _Copy<VARIANT> > > EnumVARIANT;

//============================================================================
// FUNCTIONS: CheckArgumentXXX
// PURPOSE:   Verifies an function argument.

bool CheckByValArg(LPCTSTR szFunction, LPCTSTR szArg, SHORT cValue);
bool CheckByValArg(LPCTSTR szFunction, LPCTSTR szArg, LONG cValue);
bool CheckByValArg(LPCTSTR szFunction, LPCTSTR szArg, DATE cValue);
bool CheckByValArg(LPCTSTR szFunction, LPCTSTR szArg, BSTR cValue);
bool CheckByValArg(LPCTSTR szFunction, LPCTSTR szArg, LPCSTR cValue);
bool CheckByValArg(LPCTSTR szFunction, LPCTSTR szArg, LPCWSTR cValue);
bool CheckByValBoolArg(LPCTSTR szFunction, LPCTSTR szArg, VARIANT_BOOL cValue);
bool CheckByValVariantArg(LPCTSTR szFunction, LPCTSTR szArg, VARIANT& cValue, VARTYPE vtType = VT_VARIANT, bool bOptional = false, LONG lLength = -1);

bool CheckByRefArg(LPCTSTR szFunction, LPCTSTR szArg, SHORT* pValue, bool bCheckValue = true);
bool CheckByRefArg(LPCTSTR szFunction, LPCTSTR szArg, LONG* pValue, bool bCheckValue = true);
bool CheckByRefArg(LPCTSTR szFunction, LPCTSTR szArg, DATE* pValue, bool bCheckValue = true);
bool CheckByRefArg(LPCTSTR szFunction, LPCTSTR szArg, BSTR* pValue, bool bCheckValue = true);
bool CheckByRefArg(LPCTSTR szFunction, LPCTSTR szArg, VARIANT* pValue, bool bCheckValue = true);
bool CheckByRefArg(LPCTSTR szFunction, LPCTSTR szArg, void* pValue, bool bCheckValue = true);
bool CheckByRefArg(LPCTSTR szFunction, LPCTSTR szArg, void** pValue, bool bCheckValue = true);
bool CheckByRefArg(LPCTSTR szFunction, LPCTSTR szArg, SAFEARRAY** pValue, bool bCheckValue = true);
bool CheckByRefBoolArg(LPCTSTR szFunction, LPCTSTR szArg, VARIANT_BOOL* cValue, bool bCheckValue = true);

bool CheckByRefArrayArg(LPCTSTR szFunction, LPCTSTR szArg, VARIANT* cValue, VARTYPE vtType = VT_VARIANT, LONG lLength = -1);
bool CheckByRefArrayArg(LPCTSTR szFunction, LPCTSTR szArg, SAFEARRAY** cValue, VARTYPE vtType = VT_VARIANT, LONG lLength = -1);

#define CheckByValInArg(x)                          CheckByValArg(_szFunction, _T(#x), x)
#define CheckByValInVariantArg(x, xType, xOptional) CheckByValVariantArg(_szFunction, _T(#x), x, xType, xOptional)
#define CheckByValInArrayArg(x, xType, xLength)     CheckByValVariantArg(_szFunction, _T(#x), x, VT_ARRAY | xType, true, xLength)
#define CheckByRefInArg(x)                          CheckByRefArg(_szFunction, _T(#x), x, true)
#define CheckByRefInArrayArg(x, xType, xLength)     CheckByRefArrayArg(_szFunction, _T(#x), x, xType, xLength)
#define CheckByRefOutArg(x)                         CheckByRefArg(_szFunction, _T(#x), x, false)
#define TraceByRefOutArg(x)                         CheckByRefArg(_szFunction, _T(#x), x, true)

void TraceMethodError(LPCTSTR szFunction, LPCTSTR szMethod, HRESULT hError);
void TraceInterfaceError(LPCTSTR szFunction, LPCTSTR szInterface);

//============================================================================
// FUNCTION: GetErrorString
// PURPOSE:  Attempts to lookup a string on the local machine.

HRESULT GetErrorString( 
    HRESULT dwError,
    LCID    dwLocale,
	bool    bUseAnyLocale,
    LPTSTR* ppString
);

#endif // _OPCUtils_H__
