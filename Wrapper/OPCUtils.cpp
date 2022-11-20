//============================================================================
// TITLE: OPCUtils.cpp
//
// CONTENTS:
// 
// The implementation various utility functions.
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

#include "StdAfx.h"
#include "OPCUtils.h"
#include "OPCTrace.h"

//============================================================================
// Local Declarations

#define DATA_ACCESS_PROXYSTUB  _T("opcproxy.dll")
#define MAX_ERROR_STRING       1024
#define LOCALE_ENGLISH_US      MAKELCID(MAKELANGID(LANG_ENGLISH, SUBLANG_ENGLISH_US), SORT_DEFAULT)
#define LOCALE_ENGLISH_NEUTRAL MAKELCID(MAKELANGID(LANG_ENGLISH, SUBLANG_NEUTRAL), SORT_DEFAULT)

//============================================================================
// VarTypeToStr

LPTSTR VarTypeToStr(VARTYPE vtType)
{
	static LPCTSTR g_szVarTypes[] =
	{  
		_T("VT_EMPTY"),
		_T("VT_NULL"),
		_T("VT_I2"),
		_T("VT_I4"),
		_T("VT_R4"),
		_T("VT_R8"),
		_T("VT_CY"),
		_T("VT_DATE"),
		_T("VT_BSTR"),
		_T("VT_DISPATCH"),
		_T("VT_ERROR"),
		_T("VT_BOOL"),
		_T("VT_VARIANT"),
		_T("VT_UNKNOWN"),
		_T("VT_DECIMAL"),
		_T("<invalid>"),
		_T("VT_I1"),
		_T("VT_UI1"),
		_T("VT_UI2"),
		_T("VT_UI4")
	};

	static DWORD g_dwMaxVarType = sizeof(g_szVarTypes)/sizeof(LPCTSTR);

	TCHAR szBuffer[256];
	memset(szBuffer, 0, sizeof(szBuffer));

	if ((vtType & VT_ARRAY) == 0)
	{
		if (vtType >= 0 && vtType < g_dwMaxVarType)
		{
			_stprintf(szBuffer, _T("%s [%d]"), g_szVarTypes[vtType], vtType);
		}
		else
		{
			_stprintf(szBuffer, _T("<unknown> [%d]"), vtType);
		}
	}
	else
	{
		vtType = (vtType & VT_TYPEMASK);

		if (vtType >= 0 && vtType < g_dwMaxVarType)
		{
			_stprintf(szBuffer, _T("VT_ARRAY | %s [%d]"), g_szVarTypes[vtType], vtType);
		}
		else
		{
			_stprintf(szBuffer, _T("VT_ARRAY | <unknown> [%d]"), vtType);
		}
	}

	LPTSTR szResult = (LPTSTR)CoTaskMemAlloc(sizeof(TCHAR)*(_tcslen(szBuffer)+1));
	_tcscpy(szResult, szBuffer);
	return szResult;
}

//============================================================================
// OpcDupStr

LPSTR OpcDupStr(LPCSTR szSrc)
{
	if (szSrc == NULL)
	{
		return NULL;
	}

	LPSTR szDst = (LPSTR)CoTaskMemAlloc(sizeof(CHAR)*(strlen(szSrc)+1));
	strcpy(szDst, szSrc);
	return szDst;
}

LPWSTR OpcDupStr(LPCWSTR szSrc)
{
	if (szSrc == NULL)
	{
		return NULL;
	}

	LPWSTR szDst = (LPWSTR)CoTaskMemAlloc(sizeof(WCHAR)*(wcslen(szSrc)+1));
	wcscpy(szDst, szSrc);
	return szDst;
}

//==============================================================================
// OpcFind

int OpcFind(LPCTSTR szSrc, LPCTSTR szTarget)
{
    if (szSrc == NULL || szTarget == NULL || _tcslen(szTarget) == 0)
    {
        return -1;
    }

    int iLength = _tcslen(szTarget);
    int iLast   = _tcslen(szSrc) - iLength + 1;

    for (int ii = 0; ii < iLast; ii++)
    {
        if (_tcsncmp(szSrc+ii, szTarget, iLength) == 0)
        {
            return ii;
        }
    }

    return -1;
}

//==============================================================================
// OpcReverseFind

int OpcReverseFind(LPCTSTR szSrc, LPCTSTR szTarget)
{
    if (szSrc == NULL || szTarget == NULL || _tcslen(szTarget) == 0)
    {
        return -1;
    }

    int iLength = _tcslen(szTarget);
    int iLast   = _tcslen(szSrc) - iLength;

    for (int ii = iLast; ii >=0; ii--)
    {
        if (_tcsncmp(szSrc+ii, szTarget, iLength) == 0)
        {
            return ii;
        }
    }

    return -1;
}

//==============================================================================
// OpcSubStr

LPTSTR OpcSubStr(LPCTSTR szSrc, UINT uStart, UINT uCount)
{
	UINT uLength = (szSrc == NULL)?0:_tcslen(szSrc);

    if (uCount == 0 || uStart >= uLength)
    {
        return NULL;
    }

    uLength = (uCount == -1 || uCount > uLength)?uLength-uStart:uCount;

	LPTSTR szDst = (LPTSTR)CoTaskMemAlloc((uLength+1)*sizeof(TCHAR));
	memset(szDst, 0, (uLength+1)*sizeof(TCHAR));
    _tcsncpy(szDst, szSrc+uStart, uLength);
    
	return szDst;
}

//==============================================================================
// OpcRegGetValue

bool OpcRegGetValue(
	HKEY    hBaseKey,
	LPCTSTR szSubKey,
	LPCTSTR szValueName,
	DWORD*  pdwValue
)
{
	_ASSERTE(pdwValue != NULL);

    *pdwValue = NULL;
    
    HKEY  hKey  = NULL;
    BYTE* pData = NULL;

    // open registry key.
	LONG lResult = RegOpenKeyEx(
        hBaseKey,
        szSubKey,
        NULL,
        KEY_READ,
        &hKey
    );

    if (lResult != ERROR_SUCCESS)
    {
		return false;
    }

	// fetch registry value.
	DWORD dwLength = sizeof(DWORD);

    lResult = ::RegQueryValueEx(
        hKey,
        szValueName,
        NULL,
        NULL,
        (BYTE*)pdwValue,
        &dwLength
    );

    if (lResult != ERROR_SUCCESS)
    {
		RegCloseKey(hKey);
		return false;
    }

	// close registry key.
	RegCloseKey(hKey);
	return true;
}


//==============================================================================
// OpcRegSetValue

bool OpcRegSetValue(
	HKEY    hBaseKey,
	LPCTSTR szSubKey,
	LPCTSTR szValueName,
	DWORD   dwType,
	BYTE*   pValue,
	DWORD   dwLength
)
{
	_ASSERTE(pValue != NULL);
	_ASSERTE(dwLength > 0);

	// open or create the registry key.
    HKEY hKey = NULL;
    DWORD dwDisposition = NULL;

    LONG lResult = RegCreateKeyEx(
        hBaseKey, 
        szSubKey, 
        NULL, 
        NULL,
        REG_OPTION_NON_VOLATILE,
        KEY_WRITE,
        NULL,
        &hKey,
        &dwDisposition
    );

    if (lResult != ERROR_SUCCESS)
    {
		return false;
    }

	// update registry value.
    lResult = RegSetValueEx(
        hKey,
        szValueName,
        NULL,
        dwType,
        pValue,
        dwLength
    );

    if (lResult != ERROR_SUCCESS)
    {
		RegCloseKey(hKey);
		return false;
    }

	// close registry key.
	RegCloseKey(hKey);
    return true;
}

//==============================================================================
// OpcRegSetValue

bool OpcRegSetValue(
	HKEY    hBaseKey,
	LPCTSTR szSubKey,
	LPCTSTR szValueName,
	DWORD   dwValue
)
{
    return OpcRegSetValue(
		hBaseKey, 
		szSubKey, 
		szValueName,
		REG_DWORD, 
        (BYTE*)&dwValue,
        sizeof(DWORD)
    );
}

//==============================================================================
// OpcRegDeleteKey

bool OpcRegDeleteKey(
	HKEY    hBaseKey,
	LPCTSTR tsSubKey
)
{
    bool bResult = false;

	HKEY   hKey      = NULL;
    LPTSTR tsKeyPath = NULL;
    LPTSTR tsKeyName = NULL;

    // begin block that could exit with a goto.
    {
        // parse the sub key path.
        int iIndex = OpcReverseFind(tsSubKey, _T("\\"));

        if (iIndex != -1)
        {
            tsKeyPath = OpcSubStr(tsSubKey, 0, iIndex);
            tsKeyName = OpcSubStr(tsSubKey, iIndex+1);
        }
		else
		{
			tsKeyName = OpcDupStr(tsSubKey);
		}

        // safety check - don't delete root keys
        if (tsKeyName == NULL || _tcslen(tsKeyName) == 0)
        {
			goto cleanup;
        }

        // open the key for modifications.
        LONG lResult = RegOpenKeyEx(
            hBaseKey, 
            tsSubKey, 
			NULL,
            KEY_ALL_ACCESS, 
            &hKey
        );

        if (lResult != ERROR_SUCCESS)
        {
			goto cleanup;
        }

        // determine the number of sub keys.
        DWORD dwCount = 0;
        DWORD dwMaxLength = 0;

        lResult = RegQueryInfoKey(
            hKey,         // handle to key
            NULL,         // class buffer
            NULL,         // size of class buffer
            NULL,         // reserved
            &dwCount,     // number of subkeys
            &dwMaxLength, // longest subkey name
            NULL,         // longest class string
            NULL,         // number of value entries
            NULL,         // longest value name
            NULL,         // longest value data
            NULL,         // descriptor length
            NULL          // last write time
        );

        if (lResult != ERROR_SUCCESS)
        {
			goto cleanup;
        }

        // recursively delete sub keys. 
        LPTSTR tsName = new TCHAR[dwMaxLength+1];

        while (dwCount > 0 && (lResult == ERROR_SUCCESS || lResult == ERROR_MORE_DATA))
        {
            DWORD dwLength = dwMaxLength+1;

            lResult = RegEnumKeyEx(
                hKey, 
                0, 
                tsName, 
                &dwLength,
                NULL,
                NULL,
                NULL,
                NULL);

            if (lResult == ERROR_MORE_DATA || lResult == ERROR_SUCCESS)
            {
				TCHAR tsChildKey[_MAX_PATH+1];
				memset(tsChildKey, 0, sizeof(tsChildKey));

				_tcscpy(tsChildKey, tsSubKey);
				_tcscat(tsChildKey, _T("\\"));
				_tcscat(tsChildKey, tsName);

                OpcRegDeleteKey(hBaseKey, tsChildKey);
            }
        }

        delete [] tsName;

        // close the key before delete.
        RegCloseKey(hKey);

        // open the parent key for delete.
        lResult = RegOpenKeyEx(
            hBaseKey, 
            tsKeyPath, 
            NULL,
            KEY_ALL_ACCESS, 
            &hKey
        );

        if (lResult != ERROR_SUCCESS)
        {
			goto cleanup;
        }

        // delete key
        lResult = RegDeleteKey(hKey, tsKeyName);

        if (lResult != ERROR_SUCCESS)
        {
			goto cleanup;
        }
    }

	// cleanup properly for exiting function.
	cleanup:
    {
        if (hKey != NULL)      RegCloseKey(hKey);
        if (tsKeyPath != NULL) CoTaskMemFree(tsKeyPath);
        if (tsKeyName != NULL) CoTaskMemFree(tsKeyName);
    }

    return bResult;
}

//============================================================================
// ConvertArray

/*
LONG lLBound = 0;
LONG lUBound = 0;

SafeArrayGetLBound(Value.parray, 1, &lLBound);
SafeArrayGetUBound(Value.parray, 1, &lUBound);

SAFEARRAYBOUND cBounds;
cBounds.lLbound   = lLBound;
cBounds.cElements = lUBound - lLBound + 1;

VariantClear(&m_value);

m_value.vt     = VT_ARRAY | VT_CY;
m_value.parray = SafeArrayCreate(VT_CY, 1, &cBounds);

for (LONG ii = lLBound; ii <= lUBound; ii++)
{
	DECIMAL decValue;
	memset(&decValue, 0, sizeof(decValue));
	SafeArrayGetElement(Value.parray, &ii, &decValue);

	CY cyValue;
	memset(&cyValue, 0, sizeof(cyValue));
	VarCyFromDec(&decValue, &cyValue);
	SafeArrayPutElement(m_value.parray, &ii, &cyValue);
}
*/

//============================================================================
// GetUtcTime
// 
// These functions were provided by the because they preserve the milliseconds 
// in the conversion between FILETIME and DATE and they are more efficient.

inline DATE FileTimeToDate(FILETIME *pft)
{
	return (double)((double)(*(__int64 *)pft) / 8.64e11) - (double)(363 + (1899 - 1601) * 365 + (24 + 24 + 24));
}

inline void FileTimeToDate(FILETIME *pft, DATE *pdate)
{
	*pdate = FileTimeToDate(pft);
}

inline FILETIME DateToFileTime(DATE *pdate)
{
	__int64 temp = (__int64)((*pdate + (double)(363 + (1899 - 1601) * 365 + (24 + 24 + 24))) * 8.64e11);
	return *(FILETIME *)&temp;
}

inline void DateToFileTime(DATE *pdate, FILETIME *pft)
{
	*pft = DateToFileTime(pdate);
}

DATE GetUtcTime(const FILETIME& ftUtcTime)
{
	return FileTimeToDate((FILETIME*)&ftUtcTime);
}

//============================================================================
// GetMinTime

FILETIME GetMinTime()
{
	FILETIME ftMinDate;
	memset(&ftMinDate, 0, sizeof(ftMinDate));
	return ftMinDate;
}

//============================================================================
// ElemsConvert

template <class D, class S>
inline void ElemsConvert( D *pDst, S *pSrc, long nbEntries)
{
	void *pEnd = pSrc + nbEntries;
	
	for( ; pSrc < pEnd; pSrc++, pDst++) 
	{
		*pDst = *pSrc;
	}
}

//============================================================================
// ArrayToAutomation

static HRESULT ArrayToAutomation(VARIANT &var)
{
	// Change the type of the array elements
	HRESULT hr = S_OK;

	VARTYPE vtDst = VT_EMPTY;
	switch( var.vt) {
	case VT_ARRAY | VT_I1:
		vtDst = VT_I2;
		break;
	case VT_ARRAY | VT_UI2:
		vtDst = VT_I4;
		break;
	case VT_ARRAY | VT_UI4:
		vtDst = VT_R8;
		break;
	default:
		return E_INVALIDARG;
	}
	SAFEARRAY *pSrc = var.parray;
	if( 0 != pSrc->cLocks) {
		return E_UNEXPECTED;	//If it has locks we can't convert it !
	}
	int cDims = pSrc->cDims;
	//Can't use sab on pSrc beacause its in reverse order !
	SAFEARRAYBOUND *psab = new SAFEARRAYBOUND[ cDims];
	if( NULL == psab) {
		return E_OUTOFMEMORY;
	}
	long nbEntries = 1;
	SAFEARRAYBOUND *psabSrc = pSrc->rgsabound;
	SAFEARRAYBOUND *psabDst = psab + cDims - 1;
	void *pEnd = psabSrc + cDims;
	for( ; psabSrc < pEnd; psabSrc++, psabDst--) {
		*psabDst = *psabSrc;
		nbEntries *= psabSrc->cElements;
	}
	// initialize the safearray
	SAFEARRAY *pDst = ::SafeArrayCreate( vtDst, cDims, psab);
	delete [] psab;
	if( NULL == pDst) {
		return E_OUTOFMEMORY;
	}
	void *pvSrc;
	hr = ::SafeArrayAccessData( pSrc, &pvSrc);
	if( FAILED( hr))
		goto cleanupAndExit;
	void *pvDst;
	hr = ::SafeArrayAccessData( pDst, &pvDst);
	if( FAILED( hr))	//Big problems if this happens
		goto cleanupAndExit;
	// Fill the safearray
	switch( var.vt) {
	case VT_ARRAY | VT_I1:	//Convert to I2
		ElemsConvert< short, char>( (short *)pvDst, (char *)pvSrc, nbEntries);
		break;
	case VT_ARRAY | VT_UI2: //Convert to I4
		ElemsConvert< long, unsigned short>( (long *)pvDst, (unsigned short *)pvSrc, nbEntries);
		break;
	case VT_ARRAY | VT_UI4: //Convert to R8
		ElemsConvert< double, unsigned long>( (double *)pvDst, (unsigned long *)pvSrc, nbEntries);
		break;
	}
	hr = ::SafeArrayUnaccessData( pDst);	//No problems expected
	if( FAILED( hr))
		goto cleanupAndExit;
	hr = ::SafeArrayUnaccessData( pSrc);
	if( FAILED( hr))
		goto cleanupAndExit;
	// set the VARIANT with the new array
	hr = ::VariantClear( &var);
	if( FAILED( hr))
		goto cleanupAndExit;
	var.vt = VT_ARRAY | vtDst;
	var.parray = pDst;

	return S_OK;
cleanupAndExit:
	SafeArrayDestroy( pDst);
	return hr;
}

//============================================================================
// VariantToAutomation

HRESULT VariantToAutomation(VARIANT* varValue)
{
	if (varValue == NULL)
	{
		return E_POINTER;
	}

	switch (varValue->vt)
	{
		case VT_I1:               return VariantChangeType(varValue, varValue, 0, VT_I2);
		case VT_UI2:              return VariantChangeType(varValue, varValue, 0, VT_I4);
		case VT_UI4:              return VariantChangeType(varValue, varValue, 0, VT_R8);
		case (VT_ARRAY | VT_I1):  return ArrayToAutomation(*varValue);
		case (VT_ARRAY | VT_UI2): return ArrayToAutomation(*varValue);
		case (VT_ARRAY | VT_UI4): return ArrayToAutomation(*varValue);
	}

	return S_OK;
}

//============================================================================
// CheckArgument

bool CheckByValArg(LPCTSTR szFunction, LPCTSTR szArg, SHORT cValue)
{	
	TRACE_VALUE(_T("%s, Argument '%s'='%d'"), szFunction, szArg, cValue);
	return true;
}

bool CheckByValArg(LPCTSTR szFunction, LPCTSTR szArg, LONG cValue)
{	
	TRACE_VALUE(_T("%s, Argument '%s'='%d'"), szFunction, szArg, cValue);
	return true;
}

bool CheckByValArg(LPCTSTR szFunction, LPCTSTR szArg, DATE cValue)
{	
	USES_CONVERSION;

	BSTR szValue = NULL;
	VarBstrFromDate(cValue, LOCALE_SYSTEM_DEFAULT, NULL, &szValue);

	if (szValue != NULL)
	{
		TRACE_VALUE(_T("%s, Argument '%s'='%s'"), szFunction, szArg, OLE2T(szValue));
		SysFreeString(szValue);
	}
	else
	{
		TRACE_VALUE(_T("%s, Argument '%s'='<bad date>'"), szFunction, szArg);
	}

	return true;
}

bool CheckByValArg(LPCTSTR szFunction, LPCTSTR szArg, BSTR cValue)
{	
	USES_CONVERSION;
	TRACE_VALUE(_T("%s, Argument '%s'='%s'"), szFunction, szArg, OLE2T(cValue));
	return true;
}

bool CheckByValArg(LPCTSTR szFunction, LPCTSTR szArg, LPCSTR cValue)
{	
	USES_CONVERSION;
	TRACE_VALUE(_T("%s, Argument '%s'='%s'"), szFunction, szArg, A2T((LPSTR)cValue));
	return true;
}

bool CheckByValArg(LPCTSTR szFunction, LPCTSTR szArg, LPCWSTR cValue)
{	
	USES_CONVERSION;
	TRACE_VALUE(_T("%s, Argument '%s'='%s'"), szFunction, szArg, W2T(cValue));
	return true;
}

bool CheckByValBoolArg(LPCTSTR szFunction, LPCTSTR szArg, VARIANT_BOOL cValue)
{	
	if (cValue == VARIANT_FALSE)
	{
		TRACE_VALUE(_T("%s, Argument '%s'='False'"), szFunction, szArg);
	}
	else
	{
		TRACE_VALUE(_T("%s, Argument '%s'='True'"), szFunction, szArg);
	}

	return true;
}

bool CheckByRefArg(LPCTSTR szFunction, LPCTSTR szArg, SHORT* pValue, bool bCheckValue)
{
	if (pValue == NULL)
	{
		TRACE_ERROR(_T("%s, %s is NULL."), szFunction, szArg);
		return false;
	}

	if (bCheckValue)
	{
		return CheckByValArg(szFunction, szArg, *pValue);
	}

	return true;
}

bool CheckByRefArg(LPCTSTR szFunction, LPCTSTR szArg, LONG* pValue, bool bCheckValue)
{
	if (pValue == NULL)
	{
		TRACE_ERROR(_T("%s, %s is NULL."), szFunction, szArg);
		return false;
	}

	if (bCheckValue)
	{
		return CheckByValArg(szFunction, szArg, *pValue);
	}

	return true;
}

bool CheckByRefArg(LPCTSTR szFunction, LPCTSTR szArg, DATE* pValue, bool bCheckValue)
{
	if (pValue == NULL)
	{
		TRACE_ERROR(_T("%s, %s is NULL."), szFunction, szArg);
		return false;
	}

	if (bCheckValue)
	{
		return CheckByValArg(szFunction, szArg, *pValue);
	}

	return true;
}

bool CheckByRefArg(LPCTSTR szFunction, LPCTSTR szArg, BSTR* pValue, bool bCheckValue)
{	
	if (pValue == NULL)
	{
		TRACE_ERROR(_T("%s, %s is NULL."), szFunction, szArg);
		return false;
	}

	if (bCheckValue)
	{
		return CheckByValArg(szFunction, szArg, *pValue);
	}

	return true;
}

bool CheckByRefArg(LPCTSTR szFunction, LPCTSTR szArg, VARIANT* pValue, bool bCheckValue)
{	
	if (pValue == NULL)
	{
		TRACE_ERROR(_T("%s, %s is NULL."), szFunction, szArg);
		return false;
	}

	if (bCheckValue)
	{
		return CheckByValVariantArg(szFunction, szArg, *pValue);
	}

	return true;
}

bool CheckByRefArg(LPCTSTR szFunction, LPCTSTR szArg, void* pValue, bool bCheckValue)
{	
	if (pValue == NULL)
	{
		TRACE_ERROR(_T("%s, %s is NULL."), szFunction, szArg);
		return false;
	}

	return true;
}

bool CheckByRefArg(LPCTSTR szFunction, LPCTSTR szArg, void** ppValue, bool bCheckValue)
{	
	if (ppValue == NULL)
	{
		TRACE_ERROR(_T("%s, %s is NULL."), szFunction, szArg);
		return false;
	}

	if (bCheckValue)
	{
		if (*ppValue == NULL)
		{
			TRACE_ERROR(_T("%s, %s contains a NULL."), szFunction, szArg);
			return false;
		}
	}

	return true;
}

bool CheckByRefArg(LPCTSTR szFunction, LPCTSTR szArg, SAFEARRAY** ppValue, bool bCheckValue)
{	
	if (ppValue == NULL)
	{
		TRACE_ERROR(_T("%s, %s is NULL."), szFunction, szArg);
		return false;
	}

	if (bCheckValue)
	{
		if (*ppValue == NULL)
		{
			TRACE_ERROR(_T("%s, %s contains a NULL."), szFunction, szArg);
			return false;
		}

		return CheckByRefArrayArg(szFunction, szArg, ppValue);
	}

	return true;
}

bool CheckByRefBoolArg(LPCTSTR szFunction, LPCTSTR szArg, VARIANT_BOOL* pValue, bool bCheckValue)
{	
	if (pValue == NULL)
	{
		TRACE_ERROR(_T("%s, %s is NULL."), szFunction, szArg);
		return false;
	}

	if (bCheckValue)
	{
		return CheckByValBoolArg(szFunction, szArg, *pValue);
	}

	return true;
}

bool CheckByRefArrayArg(LPCTSTR szFunction, LPCTSTR szArg, VARIANT* pValue, VARTYPE vtType, LONG lLength)
{
	if (pValue == NULL)
	{
		TRACE_ERROR(_T("%s, %s is NULL."), szFunction, szArg);
		return false;
	}

	if ((pValue->vt & VT_ARRAY) == 0)
	{
		TRACE_ERROR(_T("%s, %s is not an array."), szFunction, szArg);
		return false;
	}

	return CheckByRefArrayArg(szFunction, szArg, &pValue->parray, vtType, lLength);
}

bool CheckByRefArrayArg(LPCTSTR szFunction, LPCTSTR szArg, SAFEARRAY** ppValue, VARTYPE vtType, LONG lLength)
{
	USES_CONVERSION;

	if (ppValue == NULL)
	{
		TRACE_ERROR(_T("%s, %s is NULL."), szFunction, szArg);
		return false;
	}

	VARTYPE vtActualType = VT_EMPTY;
    SafeArrayGetVartype(*ppValue, &vtActualType);

	if (vtType != VT_VARIANT && vtActualType != vtType)
	{
		LPTSTR szActualType   = VarTypeToStr(vtActualType);
		LPTSTR szExpectedType = VarTypeToStr(vtType);

		TRACE_ERROR(_T("%s, %s contains '%s' instead of '%s'."), szFunction, szArg, szActualType, szExpectedType);

		CoTaskMemFree(szActualType);
		CoTaskMemFree(szExpectedType);
		return false;
	}

	UINT uDims = SafeArrayGetDim(*ppValue);

	if (uDims != 1)
	{
		TRACE_ERROR(_T("%s, %s has %d dimensions."), szFunction, szArg, uDims);
		return false;
	}
	
	LONG lLBound = 0;
	LONG lUBound = 0;

	SafeArrayGetLBound(*ppValue, 1, &lLBound);
	SafeArrayGetUBound(*ppValue, 1, &lUBound);

	// handle arrays passed from VB .NET that have an extra element prepended.
	if (lLBound == 0)
	{
		lLBound++;
	}

	// only reject arrays that have fewer than the expected number of elements.
	if (lLength != -1 && (lUBound - lLBound + 1) < lLength)
	{		
		TRACE_ERROR(_T("%s, %s has length %d instead if %d."), szFunction, szArg, (lUBound - lLBound + 1), lLength);
		return false;
	}
		
	LPTSTR szActualType = VarTypeToStr(vtActualType);
	TRACE_VALUE(_T("%s, Argument '%s'=Array of '%s', Size='%d'"), szFunction, szArg, szActualType, (lUBound - lLBound + 1));
	CoTaskMemFree(szActualType);

	// skip iteration through the array values if trace output is disabled.
	if ((GetTrace().GetTraceLevel() & TRACE_MASK_VALUE_INFO) == 0)
	{
		return true;
	}

	for (LONG ii = lLBound; ii <= lUBound; ii++)
	{
		HRESULT hResult = S_OK;

		VARIANT varSrc;
		VariantInit(&varSrc);

		VARIANT varDst;
		VariantInit(&varDst);

		if (vtActualType != VT_VARIANT)
		{
			varSrc.vt = vtActualType;

			switch (vtActualType)
			{
				case VT_I1:    { SafeArrayGetElement(*ppValue, &ii, &varSrc.cVal);    break; }	
				case VT_UI1:   { SafeArrayGetElement(*ppValue, &ii, &varSrc.bVal);    break; }	
				case VT_I2:    { SafeArrayGetElement(*ppValue, &ii, &varSrc.iVal);    break; }	
				case VT_UI2:   { SafeArrayGetElement(*ppValue, &ii, &varSrc.uiVal);   break; }	
				case VT_I4:    { SafeArrayGetElement(*ppValue, &ii, &varSrc.lVal);    break; }
				case VT_UI4:   { SafeArrayGetElement(*ppValue, &ii, &varSrc.ulVal);   break; }			
				case VT_R4:    { SafeArrayGetElement(*ppValue, &ii, &varSrc.fltVal);  break; }	
				case VT_R8:    { SafeArrayGetElement(*ppValue, &ii, &varSrc.dblVal);  break; }	
				case VT_CY:    { SafeArrayGetElement(*ppValue, &ii, &varSrc.cyVal);   break; }	
				case VT_BOOL:  { SafeArrayGetElement(*ppValue, &ii, &varSrc.boolVal); break; }	
				case VT_BSTR:  { SafeArrayGetElement(*ppValue, &ii, &varSrc.bstrVal); break; }	
				case VT_DATE:  { SafeArrayGetElement(*ppValue, &ii, &varSrc.date);    break; }	
				case VT_ERROR: { SafeArrayGetElement(*ppValue, &ii, &varSrc.scode);   break; }

				default:
				{
					hResult = E_FAIL;
					break;
				}
			}

			if (SUCCEEDED(hResult))
			{
				hResult = VariantChangeTypeEx(&varDst, &varSrc, LOCALE_SYSTEM_DEFAULT, VARIANT_ALPHABOOL, VT_BSTR);
			}
		}
		else
		{
			SafeArrayGetElement(*ppValue, &ii, &varSrc); 

			hResult = VariantChangeTypeEx(&varDst, &varSrc, LOCALE_SYSTEM_DEFAULT, VARIANT_ALPHABOOL, VT_BSTR);

			if (FAILED(hResult))
			{
				LPTSTR szVarType = VarTypeToStr(varSrc.vt);

				varDst.vt      = VT_BSTR;
				varDst.bstrVal = SysAllocString(T2OLE(szVarType));

				CoTaskMemFree(szVarType);

				hResult = S_OK;
			}
		}

		VariantClear(&varSrc);

		if (SUCCEEDED(hResult))
		{
			TRACE_VALUE_INFO(_T("%s, Element [%02d]='%s'"), szFunction, ii, OLE2T(varDst.bstrVal));
			VariantClear(&varDst);
		}
	}

	return true;
}

bool CheckByValVariantArg(LPCTSTR szFunction, LPCTSTR szArg, VARIANT& cValue, VARTYPE vtType, bool bOptional, LONG lLength)
{
	USES_CONVERSION;

	// In some cases, VB will pass VT_BYREF values to a method. These values must be 
	// converted to a regular value to simplify the rest of the method logic. Note that
	// this is only possible becayse

	if (cValue.vt & VT_BYREF)
	{
		if ((cValue.vt & VT_ARRAY) != 0)
		{
			cValue.vt     = (cValue.vt & VT_TYPEMASK) | VT_ARRAY;
			cValue.parray = *cValue.pparray;
		}
		else
		{
			switch (cValue.vt & VT_TYPEMASK)
			{
				case VT_BOOL: { cValue.boolVal = *cValue.pboolVal; break; } 
				case VT_I1:   { cValue.cVal    = *cValue.pcVal;    break; } 
				case VT_UI1:  { cValue.bVal    = *cValue.pbVal ;   break; } 
				case VT_I2:   { cValue.iVal    = *cValue.piVal;    break; } 
				case VT_UI2:  { cValue.uiVal   = *cValue.puiVal;   break; } 
				case VT_I4:   { cValue.lVal    = *cValue.plVal;    break; } 
				case VT_UI4:  { cValue.ulVal   = *cValue.pulVal;   break; } 
				
				#if _MSC_VER >= 1300
				case VT_I8:   { cValue.llVal   = *cValue.pllVal;   break; } 
				case VT_UI8:  { cValue.ullVal  = *cValue.pullVal;  break; } 
				#endif // _MSC_VER >= 1300
				
				case VT_R4:   { cValue.fltVal  = *cValue.pfltVal;  break; } 
				case VT_R8:   { cValue.dblVal  = *cValue.pdblVal;  break; } 
				case VT_CY:   { cValue.cyVal   = *cValue.pcyVal;   break; } 
				case VT_DATE: { cValue.date    = *cValue.pdate;    break; } 
				case VT_BSTR: { cValue.bstrVal = *cValue.pbstrVal; break; } 
				default:      { return false; }
			}
		
			cValue.vt = (cValue.vt & VT_TYPEMASK);
		}
	}

	if (vtType != VT_VARIANT && cValue.vt != vtType)
	{
		if (bOptional && (cValue.vt == VT_ERROR || cValue.vt == VT_EMPTY))
		{
			TRACE_VALUE(_T("%s, Argument '%s' was not specified."), szFunction, szArg);
			return true;
		}

		LPTSTR szActualType   = VarTypeToStr(cValue.vt);
		LPTSTR szExpectedType = VarTypeToStr(vtType);

		TRACE_ERROR(_T("%s, %s is a '%s' instead of a '%s'."), szFunction, szArg, szActualType, szExpectedType);

		CoTaskMemFree(szActualType);
		CoTaskMemFree(szExpectedType);
		return false;
	}

	if ((cValue.vt & VT_ARRAY) != 0)
	{
		CheckByRefArrayArg(szFunction, szArg, &cValue.parray, (VT_TYPEMASK & vtType), lLength);
	}
	else
	{
		VARIANT varDst;
		VariantInit(&varDst);

		HRESULT hResult = VariantChangeTypeEx(&varDst, &cValue, LOCALE_SYSTEM_DEFAULT, VARIANT_ALPHABOOL, VT_BSTR);

		if (FAILED(hResult))
		{
			LPTSTR szVarType = VarTypeToStr(cValue.vt);
			TRACE_VALUE(_T("%s, Argument '%s'='%s'"), szFunction, szArg, szVarType);
			CoTaskMemFree(szVarType);
		}
		else
		{
			TRACE_VALUE(_T("%s, Argument '%s'='%s'"), szFunction, szArg, OLE2T(varDst.bstrVal));
			VariantClear(&varDst);
		}
	}

	return true;
}

//============================================================================
// FormatMessage

static bool FormatMessage(
	LANGID  langID, 
	DWORD   dwCode, 
	LPCTSTR szModuleName, 
	TCHAR*  tsBuffer, 
	DWORD   dwLength
)
{
	// attempt to load string from module message table.
	DWORD dwResult = FormatMessage(
		FORMAT_MESSAGE_FROM_HMODULE | FORMAT_MESSAGE_IGNORE_INSERTS | FORMAT_MESSAGE_MAX_WIDTH_MASK,
		::GetModuleHandle(szModuleName), 
		dwCode,
		langID, 
		tsBuffer,
		dwLength-1,
		NULL 
	);

	// attempt to load string from system message table.
	if (dwResult == 0)
	{
		dwResult = FormatMessage(
			FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS | FORMAT_MESSAGE_MAX_WIDTH_MASK,
			NULL, 
			dwCode,
			langID, 
			tsBuffer,
			dwLength-1,
			NULL 
		);
	}

	return (dwResult != 0);
}

//============================================================================
// GetErrorString

HRESULT GetErrorString( 
    HRESULT dwError,
    LCID    dwLocale,
	bool    bUseAnyLocale,
    LPTSTR* ppString
)
{  
    // check arguments.
    if (ppString == NULL)
    {
        return E_INVALIDARG;
    }

    *ppString = NULL;

	LANGID langID = LANGIDFROMLCID(dwLocale);

	// lookup languages for 'special' locales.
	switch (dwLocale)
	{
		case LOCALE_SYSTEM_DEFAULT:
		{
			langID = GetSystemDefaultLangID();
			break;
		}

		case LOCALE_USER_DEFAULT:
		{
			langID = GetUserDefaultLangID();
			break;
		}
	}

    // initialize buffer for message.
    TCHAR tsMsg[MAX_ERROR_STRING+1];
    memset(tsMsg, 0, sizeof(tsMsg));

	if (!FormatMessage(langID, dwError, DATA_ACCESS_PROXYSTUB, tsMsg, MAX_ERROR_STRING+1))
	{			
		// give up if the requested locale is not supported.
		if (!bUseAnyLocale)
		{
			return E_INVALIDARG;
		}

		// check if the locale explicitly requested a particular language.
		if (dwLocale == LOCALE_SYSTEM_DEFAULT || dwLocale == LOCALE_USER_DEFAULT)
		{
			// try to load message using special locales.
			if (!FormatMessage(LANGIDFROMLCID(dwLocale), dwError, DATA_ACCESS_PROXYSTUB, tsMsg, MAX_ERROR_STRING+1))
			{
				// use US english if a default locale was requested but language is not supported by server.
				if (!FormatMessage(LANGIDFROMLCID(LOCALE_ENGLISH_US), dwError, DATA_ACCESS_PROXYSTUB, tsMsg, MAX_ERROR_STRING+1))
				{
					return E_INVALIDARG;
				}
			}
		}
		else
		{
			// return US english for variants for english.
			if (PRIMARYLANGID(langID) == LANG_ENGLISH)
			{	
				if (!FormatMessage(LANGIDFROMLCID(LOCALE_ENGLISH_US), dwError, DATA_ACCESS_PROXYSTUB, tsMsg, MAX_ERROR_STRING+1))
				{
					return E_INVALIDARG;
				}
			}

			// locale is not supported at all.
			else
			{
				return E_INVALIDARG;
			}
		}
	}

	// allocate string for return.
    *ppString = OpcDupStr(tsMsg);
	   
	return S_OK;
}

//============================================================================
// TraceMethodError

void TraceMethodError(LPCTSTR szFunction, LPCTSTR szMethod, HRESULT hError)
{
	LPTSTR tsError = NULL;

	if (FAILED(GetErrorString(hError, LOCALE_SYSTEM_DEFAULT, true, &tsError)))
	{
		tsError = NULL;
	}

	TRACE_ERROR(_T("%s, %s Failed, Error=[0x%08X] '%s'"), szFunction, szMethod, hError, tsError);
	CoTaskMemFree(tsError);
}

//============================================================================
// TraceInterfaceError

void TraceInterfaceError(LPCTSTR szFunction, LPCTSTR szInterface)
{
	TRACE_ERROR(_T("%s, '%s' interface not supported."), szFunction, szInterface);
}
