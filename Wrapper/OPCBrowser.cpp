//==============================================================================
// TITLE: OPCBrowser.cpp
//
// CONTENTS:
//
//  This is the Browser object.
//
// (c) Copyright 1998-2004 The OPC Foundation
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
// 1998/03/07 JH    First release.
// 2003/03/06 RSA   Updated formatting to complying with coding guidelines.
// 2003/10/27 RSA   Merged bug fixes received from members.

#include "StdAfx.h"
#include "OPCBrowser.h"
#include "OPCUtils.h"
#include "OPCTrace.h"

//==============================================================================
// Local Declarations

#define BLOCK_SIZE 100

//==============================================================================
// Constructor

COPCBrowser::COPCBrowser()
: 
	m_filter("")
{
   m_dataType     = VT_EMPTY;
   m_accessRights = OPC_WRITEABLE | OPC_READABLE;
}

//==============================================================================
// Destructor

COPCBrowser::~COPCBrowser()
{
   ClearNames();
}

//==============================================================================
// Organization Property

STDMETHODIMP COPCBrowser::get_Organization(LONG* pOrganization)
{
	TRACE_FUNCTION_ENTER("COPCBrowser::get_Organization");

	// check arguments.
	if (!CheckByRefOutArg(pOrganization)) return E_INVALIDARG;

	// fetch namespace organization from the server.
	OPCNAMESPACETYPE eNamespaceType = OPC_NS_HIERARCHIAL;

	TRACE_METHOD_CALL(QueryOrganization);

	HRESULT hResult = m_pOPCBrowser->QueryOrganization(&eNamespaceType);
	
	if (FAILED(hResult))
	{
		TRACE_METHOD_ERROR(QueryOrganization, hResult);
		return hResult;
	}

	*pOrganization = (LONG)eNamespaceType;
	
	// write return values to log.
	TraceByRefOutArg(pOrganization);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// Filter Property

STDMETHODIMP COPCBrowser::get_Filter(BSTR* pFilter)
{
	USES_CONVERSION;

	TRACE_FUNCTION_ENTER("COPCBrowser::get_Filter");

	// check arguments.
	if (!CheckByRefOutArg(pFilter)) return E_INVALIDARG;

	*pFilter = m_filter.Copy();

	// write return values to log.
	TraceByRefOutArg(pFilter);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

STDMETHODIMP COPCBrowser::put_Filter(BSTR Filter)
{
	USES_CONVERSION;

	TRACE_FUNCTION_NAME("COPCBrowser::put_Filter");

	// check arguments.
	if (!CheckByValInArg(Filter)) return E_INVALIDARG;

	if (Filter == NULL)
	{
		m_filter = _T("");
	}
	else
	{
		m_filter = Filter;
	}

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// DataType Property

STDMETHODIMP COPCBrowser::get_DataType(SHORT* pDataType)
{
	TRACE_FUNCTION_ENTER("COPCBrowser::get_DataType");

	// check arguments.
	if (!CheckByRefOutArg(pDataType)) return E_INVALIDARG;

	*pDataType = m_dataType;

	// write return values to log.
	TraceByRefOutArg(pDataType);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

STDMETHODIMP COPCBrowser::put_DataType(SHORT DataType)
{
	TRACE_FUNCTION_NAME("COPCBrowser::put_DataType");

	// check arguments.
	if (!CheckByValInArg(DataType)) return E_INVALIDARG;

	m_dataType = DataType;

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// AccessRights Property

STDMETHODIMP COPCBrowser::get_AccessRights(LONG* pAccessRights)
{
	TRACE_FUNCTION_ENTER("COPCBrowser::get_AccessRights");

	// check arguments.
	if (!CheckByRefOutArg(pAccessRights)) return E_INVALIDARG;

	*pAccessRights = m_accessRights;

	// write return values to log.
	TraceByRefOutArg(pAccessRights);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

STDMETHODIMP COPCBrowser::put_AccessRights(LONG AccessRights)
{
	TRACE_FUNCTION_NAME("COPCBrowser::put_AccessRights");

	// check arguments.
	if (!CheckByValInArg(AccessRights)) return E_INVALIDARG;
	
	m_accessRights = AccessRights;

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// CurrentPosition Property

STDMETHODIMP COPCBrowser::get_CurrentPosition(BSTR* pCurrentPosition)
{
	USES_CONVERSION;

	TRACE_FUNCTION_ENTER("COPCBrowser::get_CurrentPosition");
	
	// check arguments.
	if (!CheckByRefOutArg(pCurrentPosition)) return E_INVALIDARG;

	// get the item id returned when by passing an empty name to GetItemID().
	LPWSTR pName = NULL;
	
	TRACE_METHOD_CALL(GetItemID);

	HRESULT hResult = m_pOPCBrowser->GetItemID(L"", &pName);
	
	if (SUCCEEDED(hResult))
	{
		*pCurrentPosition = SysAllocString(pName);
		CoTaskMemFree(pName);
	}

	// write return values to log.
	TraceByRefOutArg(pCurrentPosition);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// Count Property

STDMETHODIMP COPCBrowser::get_Count(LONG* pCount)
{
	TRACE_FUNCTION_ENTER("COPCBrowser::get_Count");

	// check arguments.
	if (!CheckByRefOutArg(pCount)) return E_INVALIDARG;

	*pCount = m_names.size();

	// write return values to log.
	TraceByRefOutArg(pCount);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// _NewEnum Property

STDMETHODIMP COPCBrowser::get__NewEnum(IUnknown** ppUnk)
{
	TRACE_FUNCTION_ENTER("COPCBrowser::get__NewEnum");

	// check arguments.
	if (!CheckByRefOutArg(ppUnk)) return E_INVALIDARG;

	*ppUnk = NULL;

	EnumVARIANT* pEnumerator = new EnumVARIANT();

	HRESULT hResult = S_OK;

	// create an empty IEnumVariant object.
	if (m_names.size() == 0)
	{
		CComVariant varEmpty(_T(""));
		hResult = pEnumerator->Init(&varEmpty, &varEmpty, NULL, AtlFlagCopy);
	}

	// create an array of variants containing BSTRs
	else
	{
		VARIANT* pArray = (VARIANT*)malloc(m_names.size()*sizeof(VARIANT));

		list<CComBSTR*>::iterator iter = m_names.begin();
		
		for (LONG ii = 0; iter != m_names.end(); ii++)
		{
			CComBSTR* pName = *iter;

			// and put it into a variant
			VariantInit(&pArray[ii]);
			
			pArray[ii].vt      = VT_BSTR;
			pArray[ii].bstrVal = SysAllocString(*pName);
			
			iter++;
		}

		// create an IEnumVariant object initialized with these variants
		hResult = pEnumerator->Init(&pArray[0], &pArray[m_names.size()], NULL, AtlFlagCopy);
		
		// clear the local variant array.
		for (LONG ii = 0; ii < (LONG)m_names.size(); ii++)
		{
			VariantClear(&pArray[ii]);
		}

		free(pArray);
	}

	// query for an interface to return to the caller.
	if (SUCCEEDED(hResult))
	{
		hResult = pEnumerator->QueryInterface(IID_IEnumVARIANT, (void**)ppUnk);
	}

	// delete enumerator on error.
	if (FAILED(hResult))
	{
		delete pEnumerator;
	   	TRACE_ERROR("%s, 'hResult=[0x%08X]'", TRACE_FUNCTION, hResult);
		return hResult;
	}

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// Item Property

STDMETHODIMP COPCBrowser::Item(VARIANT ItemSpecifier, BSTR* pItem)
{
	TRACE_FUNCTION_ENTER("COPCBrowser::Item");

	// check arguments.
	if (!CheckByValInVariantArg(ItemSpecifier, VT_VARIANT, false)) return E_INVALIDARG;
	if (!CheckByRefOutArg(pItem))                                  return E_INVALIDARG;
	
	if (ItemSpecifier.vt != VT_I4 && ItemSpecifier.vt != VT_I2) 
	{ 
	   	TRACE_ERROR("%s, 'ItemSpecifier.vt != VT_I4 && ItemSpecifier.vt != VT_I2'", TRACE_FUNCTION);
		return E_INVALIDARG; 
	}

	LONG lIndex = (ItemSpecifier.vt == VT_I4)?ItemSpecifier.lVal:ItemSpecifier.iVal;

	// the indexer is one based.
	if (lIndex < 1 || lIndex > (LONG)m_names.size())
	{
	   	TRACE_ERROR("%s, 'lIndex < 1 || lIndex > (LONG)m_names.size()'", TRACE_FUNCTION);
		return E_INVALIDARG; 
	}

	// search list for item.
	list<CComBSTR*>::iterator iter = m_names.begin();
	
	for (LONG ii = 1; iter != m_names.end(); ii++)
	{
		CComBSTR* pName = *iter;

		if (ii == lIndex)
		{
			*pItem = pName->Copy();

			TraceByRefOutArg(pItem);

			TRACE_FUNCTION_LEAVE();
			return S_OK;
		}
				
		iter++;
	}

	TRACE_ERROR("%s, 'Index not Found'", TRACE_FUNCTION);
	return E_FAIL;
}

//==============================================================================
// ShowBranches Method

STDMETHODIMP COPCBrowser::ShowBranches()
{
	USES_CONVERSION;

	TRACE_FUNCTION_ENTER("COPCBrowser::ShowBranches");

	// clear existing names.
	ClearNames();

	// Get the Branch names
	IEnumString* ipEnumString = NULL;
	
	TRACE_METHOD_CALL(BrowseOPCItemIDs);

	HRESULT hResult = m_pOPCBrowser->BrowseOPCItemIDs(
		OPC_BRANCH,
		m_filter,
		VT_EMPTY,
		0,
		&ipEnumString
	);

	if (FAILED(hResult))
	{
		TRACE_METHOD_ERROR(BrowseOPCItemIDs, hResult);
		return hResult;
	}

	// some servers return an invalid enumerator instead of an empty list.
	if (ipEnumString == NULL)
	{
		TRACE_FUNCTION_LEAVE();
		return S_OK;
	}

	// enumerate all the names, saving them in the list
	LPWSTR pNames[BLOCK_SIZE];
	memset(pNames, 0, sizeof(pNames));

	ULONG ulCount = 0;
	
	do
	{
		hResult = ipEnumString->Next(BLOCK_SIZE, pNames, &ulCount);

		for (ULONG ii = 0; ii < ulCount; ii++)
		{
			TRACE_MISC(_T("%s, Branch, '%s'"), TRACE_FUNCTION, OLE2T(pNames[ii]));

			m_names.push_back(new CComBSTR(pNames[ii]));
			CoTaskMemFree(pNames[ii]);
		}
	}
	while (hResult == S_OK);

	// release enumerator.
	ipEnumString->Release();

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// ShowLeafs Method

STDMETHODIMP COPCBrowser::ShowLeafs(VARIANT Flat)
{
	USES_CONVERSION;

	TRACE_FUNCTION_ENTER("COPCBrowser::ShowLeafs");

	// check arguments.
	if (!CheckByValInVariantArg(Flat, VT_BOOL, true)) return E_INVALIDARG;

	// clear existing names.
	ClearNames();

	// Get the left names
	IEnumString* ipEnumString = NULL;
	
	HRESULT hResult = S_OK;
	
	TRACE_METHOD_CALL(BrowseOPCItemIDs);

	// browse the items as a flat space.
	if (Flat.vt == VT_BOOL && Flat.boolVal == VARIANT_TRUE)
	{
		hResult = m_pOPCBrowser->BrowseOPCItemIDs(
			OPC_FLAT,
			m_filter,
			m_dataType,
			m_accessRights,
			&ipEnumString
		);
	}

	// browse the leafs at the current position.
	else
	{
		hResult = m_pOPCBrowser->BrowseOPCItemIDs(
			OPC_LEAF,
			m_filter,
			m_dataType,
			m_accessRights,
			&ipEnumString
		);
	}

	if (FAILED(hResult))
	{
		TRACE_METHOD_ERROR(BrowseOPCItemIDs, hResult);
		return hResult;
	}

	// some servers return an invalid enumerator instead of an empty list.
	if (ipEnumString == NULL)
	{
		TRACE_FUNCTION_LEAVE();
		return S_OK;
	}

	// enumerate all the names, saving them in the list
	LPWSTR pNames[BLOCK_SIZE];
	memset(pNames, 0, sizeof(pNames));

	ULONG ulCount = 0;
	
	do
	{
		hResult = ipEnumString->Next(BLOCK_SIZE, pNames, &ulCount);

		for (ULONG ii = 0; ii < ulCount; ii++)
		{
			TRACE_MISC(_T("%s, Leaf, '%s'"), TRACE_FUNCTION, OLE2T(pNames[ii]));
			
			m_names.push_back(new CComBSTR(pNames[ii]));
			CoTaskMemFree(pNames[ii]);
		}
	}
	while (hResult == S_OK);

	// release enumerator.
	ipEnumString->Release();

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// MoveUp Method

STDMETHODIMP COPCBrowser::MoveUp()
{
	TRACE_FUNCTION_ENTER("COPCBrowser::MoveUp");

	// clear existing names.
	ClearNames();
   
	// change browse position.
	TRACE_METHOD_CALL(ChangeBrowsePosition);

	HRESULT hResult = m_pOPCBrowser->ChangeBrowsePosition(OPC_BROWSE_UP, L"");

	if (FAILED(hResult))
	{	   	
		TRACE_METHOD_ERROR(ChangeBrowsePosition, hResult);
		return hResult;
	}

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// MoveToRoot Method

STDMETHODIMP COPCBrowser::MoveToRoot()
{
	TRACE_FUNCTION_ENTER("COPCBrowser::MoveToRoot");

	// clear existing names.
	ClearNames();
   
	// change browse position the hard way since V1.0 servers don't support OPC_BROWSE_TO.
	HRESULT hResult = S_OK;
	
	TRACE_METHOD_CALL(ChangeBrowsePosition);

	while (SUCCEEDED(hResult))
	{
		hResult = m_pOPCBrowser->ChangeBrowsePosition(OPC_BROWSE_UP, L"");
	}

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// MoveDown Method

STDMETHODIMP COPCBrowser::MoveDown(BSTR Branch)
{
	TRACE_FUNCTION_ENTER("COPCBrowser::MoveDown");

	// clear existing names.
	ClearNames();
   
	// change browse position.
	TRACE_METHOD_CALL(ChangeBrowsePosition);

	HRESULT hResult = m_pOPCBrowser->ChangeBrowsePosition(OPC_BROWSE_DOWN, Branch);

	if (FAILED(hResult))
	{	   		
		TRACE_METHOD_ERROR(ChangeBrowsePosition, hResult);
		return hResult;
	}

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// MoveTo Method

STDMETHODIMP COPCBrowser::MoveTo(SAFEARRAY** ppBranches)
{
	TRACE_FUNCTION_ENTER("COPCBrowser::MoveTo");
   		
	// check arguments.
	if (!CheckByRefInArg(ppBranches)) return E_INVALIDARG;

	// clear existing names.
	ClearNames();

	HRESULT hResult = S_OK;
	
	LONG lBound = 0;
	LONG uBound = 0;

	hResult = SafeArrayGetLBound(*ppBranches, 1, &lBound);
	
	if (FAILED(hResult))
	{
		TRACE_ERROR("%s, 'SafeArrayGetLBound Failed' [hResult=0x%08X]", TRACE_FUNCTION, hResult);
		return hResult;
	}

	hResult = SafeArrayGetUBound(*ppBranches, 1, &uBound);
	
	if (FAILED(hResult))
	{
		TRACE_ERROR("%s, 'SafeArrayGetUBound Failed' [hResult=0x%08X]", TRACE_FUNCTION, hResult);
		return hResult;
	}

	// move to root
	hResult = MoveToRoot();
	
	if (FAILED(hResult))
	{	   	
		TRACE_FUNCTION_LEAVE();
		return hResult;
	}

	for (LONG ii = lBound; ii <= uBound; ii++)
	{ 
		// fetch branch name from array.
		BSTR bstrBranch = NULL;

		hResult = SafeArrayGetElement(*ppBranches, &ii, &bstrBranch);

		if (FAILED(hResult))
		{
			TRACE_ERROR("%s, 'hResult=[0x%08X]'", TRACE_FUNCTION, hResult);
			return hResult;
		}

		// move to branch.
		if (bstrBranch != NULL && bstrBranch[0] != 0)
		{
			TRACE_METHOD_CALL(ChangeBrowsePosition);
			hResult = m_pOPCBrowser->ChangeBrowsePosition(OPC_BROWSE_DOWN, bstrBranch);
		}

		// free memory and check error.
		SysFreeString(bstrBranch);

		if (FAILED(hResult))
		{	   	
			TRACE_METHOD_ERROR(ChangeBrowsePosition, hResult);
			return hResult;
		}
	}

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// GetItemID Method

STDMETHODIMP COPCBrowser::GetItemID(BSTR Leaf, BSTR* pItemID)
{
	TRACE_FUNCTION_ENTER("COPCBrowser::GetItemID");
	
	// check arguments.
	if (!CheckByValInArg(Leaf))     return E_INVALIDARG;
	if (!CheckByRefOutArg(pItemID)) return E_INVALIDARG;

	// fetch item id from server.
	LPWSTR pName = NULL;

	TRACE_METHOD_CALL(GetItemID);

	HRESULT hResult = m_pOPCBrowser->GetItemID(Leaf, &pName);
	
	if (FAILED(hResult))
	{
		TRACE_METHOD_ERROR(GetItemID, hResult);
		return hResult;

	}
	
	*pItemID = SysAllocString(pName);
	CoTaskMemFree(pName);

	TraceByRefOutArg(pItemID);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// GetAccessPaths Method

STDMETHODIMP COPCBrowser::GetAccessPaths(BSTR ItemID, VARIANT* pAccessPaths)
{
	TRACE_FUNCTION_ENTER("COPCBrowser::GetAccessPaths");

	// check arguments.
	if (!CheckByValInArg(ItemID))        return E_INVALIDARG;
	if (!CheckByRefOutArg(pAccessPaths)) return E_INVALIDARG;

	VariantInit(pAccessPaths);

	// fetch the access path enumerator.
	IEnumString* ipPaths = NULL;

	TRACE_METHOD_CALL(BrowseAccessPaths);

	HRESULT hResult = m_pOPCBrowser->BrowseAccessPaths(ItemID, &ipPaths);
	
	if (FAILED(hResult))
	{
		TRACE_METHOD_ERROR(BrowseAccessPaths, hResult);
		return hResult;
	}

	// fetch the access paths.
	list<LPWSTR> paths;

	do
	{
		ULONG  ulFetched = 0;
		LPWSTR szPath    = NULL;

		TRACE_METHOD_CALL(IEnumString::Next);

		hResult = ipPaths->Next(1, &szPath, &ulFetched);

		if (ulFetched < 1)
		{
			break;
		}

		if (SUCCEEDED(hResult))
		{
			paths.push_back(szPath);
		}
	}
	while (SUCCEEDED(hResult));

	// create returned array.
	if (paths.size() > 0)
	{
		SAFEARRAYBOUND cBounds;
		cBounds.lLbound   = 1;
		cBounds.cElements = paths.size();

		SAFEARRAY* pArray = SafeArrayCreate(VT_BSTR, 1, &cBounds);

		list<LPWSTR>::iterator iter = paths.begin();

		for(LONG ii = 1; iter != paths.end(); ii++)
		{
			LPWSTR szPath = *iter;
			BSTR bstrPath = SysAllocString(szPath);
			
			SafeArrayPutElement(pArray, &ii, bstrPath);
			
            CoTaskMemFree(szPath);
            SysFreeString(bstrPath);

			iter++;
		}

		paths.clear();

		pAccessPaths->vt     = VT_ARRAY | VT_BSTR;
		pAccessPaths->parray = pArray;
	}

	TraceByRefOutArg(pAccessPaths);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// ClearNames Method

void COPCBrowser::ClearNames()
{
   list<CComBSTR*>::iterator iter = m_names.begin();

   for(;iter != m_names.end(); iter++)
   {
      CComBSTR* pName = *iter;
      delete pName;
   }

   m_names.clear();
}
