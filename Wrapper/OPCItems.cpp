//==============================================================================
// TITLE: OPCItems.cpp
//
// CONTENTS:
//
//  This is the OPCItems collection in the Group object.
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
// 2004/10/22 RSA   Added check for errors during item activation in AddItems.
//

#include "StdAfx.h"
#include "OPCItems.h"
#include "OPCGroup.h"
#include "OPCUtils.h"
#include "OPCTrace.h"

//==============================================================================
// Constructor

COPCItems::COPCItems()
{
   m_pParent         = NULL;
   m_defaultDataType = VT_EMPTY;
   m_defaultActive   = VARIANT_TRUE;
	
   m_defaultAccessPath.Empty();
}

//==============================================================================
// Destructor

COPCItems::~COPCItems()
{
	Disconnect();
}

//==============================================================================
// Initialize

void COPCItems::Initialize(COPCGroup* pParent)
{
	m_pParent = pParent;
}

//==============================================================================
// Connect

void COPCItems::Connect(IUnknown* ipUnknown)
{
	TRACE_FUNCTION_ENTER("COPCItems::Connect");

	m_pOPCGroup = ipUnknown;
	m_pOPCItem  = ipUnknown;

	TRACE_FUNCTION_LEAVE();
}

//==============================================================================
// Disconnect

void COPCItems::Disconnect()
{
	TRACE_FUNCTION_ENTER("COPCItems::Disconnect");

	RemoveAllItems();

	m_pOPCItem.Release();
	m_pOPCGroup.Release();

	TRACE_FUNCTION_LEAVE();
}

//==============================================================================
// RemoveAllItems Method

void COPCItems::RemoveAllItems()
{
	TRACE_FUNCTION_NAME("COPCItems::RemoveAllItems");

	HRESULT hResult = S_OK;

	// create array of handles
	int iNumItems = m_items.size();

	if (iNumItems > 0)
	{
		OPCHANDLE* phHandles = (OPCHANDLE*)malloc(iNumItems*sizeof(OPCHANDLE));
		memset(phHandles, 0, iNumItems*sizeof(OPCHANDLE));

		list<COPCItem*>::iterator iter = m_items.begin();

		for (int ii = 0; ii < iNumItems && iter != m_items.end(); ii++)
		{
			COPCItem* pItem = *iter;
			phHandles[ii] = pItem->GetServerHandle();
			pItem->Release();
			iter++;
		}

		HRESULT* pErrors = NULL;

		TRACE_METHOD_CALL(RemoveItems);

		hResult = m_pOPCItem->RemoveItems(iNumItems, phHandles, &pErrors);

		if (SUCCEEDED(hResult))
		{
			CoTaskMemFree(pErrors);
		}

		free(phHandles);
	}

	// clear array of items.
	m_items.clear();
}


//==============================================================================
// LookupItem Method

COPCItem* COPCItems::LookupItem(OPCHANDLE ServerHandle)
{
	TRACE_FUNCTION_NAME("COPCItems::LookupItem");

	// do a direct cast.
	COPCItem* pItem = (COPCItem*)ServerHandle;

	// check memory address to ensure cast returned a valid object.
	if (pItem->MagicNumber == MAGIC_NUMBER)
	{
		return pItem;
	}

	// report an error.
	TRACE_ERROR(_T("%s, Item Handle [0x%08X] failed the magic number test."), TRACE_FUNCTION, ServerHandle);

	// do it the in-efficient way (should never be necessary but left the code in in case).
	list<COPCItem*>::iterator iter = m_items.begin();

	HRESULT hResult = E_FAIL;

	while (iter != m_items.end())
	{
		pItem = *iter;

		if ((OPCHANDLE)pItem == ServerHandle)
		{
			return pItem;
		}

		++iter;
	}

	TRACE_ERROR("%s, ServerHandle='0x%08X' not found.'", TRACE_FUNCTION, ServerHandle);
	return NULL;
}

//==============================================================================
// GetInvalidHandle Method

OPCHANDLE COPCItems::GetInvalidHandle()
{
	TRACE_FUNCTION_NAME("COPCItems::GetInvalidHandle");

	OPCHANDLE hMaxHandle = 0;

	list<COPCItem*>::iterator iter = m_items.begin();

	while (iter != m_items.end())
	{
		OPCHANDLE hHandle = (*iter)->GetServerHandle();

		if (hMaxHandle < hHandle)
		{
			hMaxHandle = hHandle;
		}

		iter++;
	}

	return hMaxHandle++;
}

//==============================================================================
// Parent Property

STDMETHODIMP COPCItems::get_Parent(OPCGroup** ppParent)
{
	TRACE_FUNCTION_ENTER("COPCItems::get_Parent");

	// check arguments.
	if (!CheckByRefOutArg(ppParent)) return E_INVALIDARG;

	*ppParent = NULL;

	// get interface pointer
	m_pParent->QueryInterface(IID_IOPCGroup, (void**)ppParent);
	_ASSERTE(*ppParent);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// DefaultRequestedDataType Property

STDMETHODIMP COPCItems::get_DefaultRequestedDataType(SHORT* pDefaultRequestedDataType)
{
	TRACE_FUNCTION_ENTER("COPCItems::get_DefaultRequestedDataType");

	// check arguments.
	if (!CheckByRefOutArg(pDefaultRequestedDataType)) return E_INVALIDARG;

	*pDefaultRequestedDataType = m_defaultDataType;

	TraceByRefOutArg(pDefaultRequestedDataType);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

STDMETHODIMP COPCItems::put_DefaultRequestedDataType(SHORT DefaultRequestedDataType)
{
	TRACE_FUNCTION_ENTER("COPCItems::put_DefaultRequestedDataType");

	// check arguments.
	if (!CheckByValInArg(DefaultRequestedDataType)) return E_INVALIDARG;

	m_defaultDataType = DefaultRequestedDataType;

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// DefaultAccessPath Property

STDMETHODIMP COPCItems::get_DefaultAccessPath(BSTR* pDefaultAccessPath)
{
	TRACE_FUNCTION_ENTER("COPCItems::get_DefaultAccessPath");

	// check arguments.
	if (!CheckByRefOutArg(pDefaultAccessPath)) return E_INVALIDARG;

	*pDefaultAccessPath = m_defaultAccessPath.Copy();

	TraceByRefOutArg(pDefaultAccessPath);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

STDMETHODIMP COPCItems::put_DefaultAccessPath(BSTR DefaultAccessPath)
{
	TRACE_FUNCTION_ENTER("COPCItems::put_DefaultRequestedDataType");

	// check arguments.
	if (!CheckByValInArg(DefaultAccessPath)) return E_INVALIDARG;

	m_defaultAccessPath = DefaultAccessPath;

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// DefaultIsActive Property

STDMETHODIMP COPCItems::get_DefaultIsActive(VARIANT_BOOL* pDefaultIsActive)
{
	TRACE_FUNCTION_ENTER("COPCItems::get_DefaultIsActive");

	// check arguments.
	if (!CheckByRefOutArg(pDefaultIsActive)) return E_INVALIDARG;

	*pDefaultIsActive = (m_defaultActive)?VARIANT_TRUE:VARIANT_FALSE;

	TraceByRefOutArg(pDefaultIsActive);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

STDMETHODIMP COPCItems::put_DefaultIsActive(VARIANT_BOOL DefaultIsActive)
{
	TRACE_FUNCTION_ENTER("COPCItems::put_DefaultRequestedDataType");

	// check arguments.
	if (!CheckByValInArg(DefaultIsActive)) return E_INVALIDARG;

	m_defaultActive = DefaultIsActive;

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// Count Property

STDMETHODIMP COPCItems::get_Count(LONG* pCount)
{
	TRACE_FUNCTION_ENTER("COPCItems::get_Count");

	// check arguments.
	if (!CheckByRefOutArg(pCount)) return E_INVALIDARG;

    *pCount = m_items.size();

	TraceByRefOutArg(pCount);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// _NewEnum Property

STDMETHODIMP COPCItems::get__NewEnum(IUnknown** ppUnknown)
{
	TRACE_FUNCTION_ENTER("COPCItems::get__NewEnum");

	// check arguments.
	if (!CheckByRefOutArg(ppUnknown)) return E_INVALIDARG;

	*ppUnknown = NULL;

	EnumVARIANT* pEnumerator = new EnumVARIANT();

	HRESULT hResult = S_OK;

	// create an empty IEnumVariant object.
	if (m_items.size() == 0)
	{
		CComVariant varEmpty(_T(""));
		hResult = pEnumerator->Init(&varEmpty, &varEmpty, NULL, AtlFlagCopy);
	}
		
	// create an array of variants containing IDispatch pointers.
	else
	{
		VARIANT* pArray = (VARIANT*)malloc(m_items.size()*sizeof(VARIANT));

		list<COPCItem*>::iterator iter = m_items.begin();
		
		for (LONG ii = 0; iter != m_items.end(); ii++)
		{
			COPCItem* pItem = *iter;

			// fetch pointer to group.
			OPCItem* ipItem = NULL;
			pItem->QueryInterface(IID_OPCItem, (void**)&ipItem);
			_ASSERTE(ipItem);
			
			// and put it into a variant
			VariantInit(&pArray[ii]);
			pArray[ii].vt       = VT_DISPATCH;
			pArray[ii].pdispVal = ipItem;
			
			iter++;
		}

		// create an IEnumVariant object initialized with these variants
		hResult = pEnumerator->Init(&pArray[0], &pArray[m_items.size()], NULL, AtlFlagCopy);
		
		// clear the local variant array.
		for (ii = 0; ii < (LONG)m_items.size(); ii++)
		{
			VariantClear(&pArray[ii]);
		}

		free(pArray);
	}

	// query for an interface to return to the caller.
	if (SUCCEEDED(hResult))
	{
		hResult = pEnumerator->QueryInterface(IID_IEnumVARIANT, (void**)ppUnknown);
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

STDMETHODIMP COPCItems::Item(VARIANT ItemSpecifier,  OPCItem** ppItem)
{
	TRACE_FUNCTION_ENTER("COPCItems::Item");

	// check arguments.
	if (!CheckByRefOutArg(ppItem)) return E_INVALIDARG;

	// get indexer.
	LONG lIndex   = 0;
	BSTR bstrName = NULL;

	switch (ItemSpecifier.vt)
	{
		case VT_I2:   { lIndex   = ItemSpecifier.iVal;    break; }
		case VT_I4:	  { lIndex   = ItemSpecifier.lVal;    break; }
		case VT_BSTR: {	bstrName = ItemSpecifier.bstrVal; break; }

		default:
		{
	   		TRACE_ERROR("%s, 'ItemSpecifier is not valid.'", TRACE_FUNCTION);
			return E_INVALIDARG;
		}
	}

	// find matching item.
	list<COPCItem*>::iterator iter = m_items.begin();
	
	for (LONG ii = 0; iter != m_items.end(); ii++)
	{
		COPCItem* pItem = *iter;

		if (((bstrName != NULL) && (wcscmp(pItem->GetItemID(), bstrName) == 0)) || ((bstrName == NULL) && (ii+1 == lIndex)))
		{
			pItem->QueryInterface(IID_OPCItem, (void**)ppItem);
			_ASSERTE(*ppItem);

			TRACE_FUNCTION_LEAVE();
			return S_OK;
		}

		iter++;
	}

	TRACE_ERROR("%s, 'E_INVALIDARG'", TRACE_FUNCTION);
	return E_INVALIDARG;
}

//==============================================================================
// GetOPCItem Property

STDMETHODIMP COPCItems::GetOPCItem(LONG ServerHandle, OPCItem** ppItem)
{
	TRACE_FUNCTION_ENTER("OPCItems::GetOPCItem");

	// check arguments.
	if (!CheckByRefOutArg(ppItem)) return E_INVALIDARG;

	// find matching group.
	list<COPCItem*>::iterator iter = m_items.begin();
	
	for (LONG ii = 0; iter != m_items.end(); ii++)
	{
		COPCItem* pItem = *iter;

		if (ServerHandle == (OPCHANDLE)pItem)
		{
			pItem->QueryInterface(IID_OPCItem, (void**)ppItem);
			_ASSERTE(*ppItem);

			TRACE_FUNCTION_LEAVE();
			return S_OK;
		}

		iter++;
	}

	TRACE_ERROR("%s, 'E_INVALIDARG'", TRACE_FUNCTION);
	return E_INVALIDARG;
}

//==============================================================================
// AddItem Method

STDMETHODIMP COPCItems::AddItem( 
	BSTR      ItemID,
	LONG      ClientHandle,
	OPCItem** ppItem
)
{
	TRACE_FUNCTION_ENTER("OPCItems::AddItem");

	// check arguments.
	if (!CheckByValInArg(ItemID))       return E_INVALIDARG;
	if (!CheckByValInArg(ClientHandle)) return E_INVALIDARG;
	if (!CheckByRefOutArg(ppItem))      return E_INVALIDARG;

	// initialize return parameters.
	*ppItem = NULL;
	
	// check if interface is supported.
	if (!m_pOPCItem)
	{
		TRACE_INTERFACE_ERROR(IOPCItemMgt);
		return E_NOINTERFACE;
	}

	BOOL bIsActive = m_defaultActive;

	SAFEARRAYBOUND cBounds;
	cBounds.lLbound   = 1;
	cBounds.cElements = 1;

	SAFEARRAY* psaItemIDs = SafeArrayCreate(VT_BSTR, 1, &cBounds);
	SafeArrayPutElement(psaItemIDs, &cBounds.lLbound, ItemID);

	// create item definitions.
	OPCITEMDEF* pItemDefs = CreateDefinitions(1, psaItemIDs, NULL, NULL);

	SafeArrayDestroy(psaItemIDs);

	// add item to group.
    OPCITEMRESULT* pResults = NULL;
    HRESULT*       pErrors  = NULL;

	TRACE_METHOD_CALL(AddItems);

	HRESULT hResult = m_pOPCItem->AddItems(1, pItemDefs, &pResults, &pErrors);

	// check method level error.
	if (FAILED(hResult))
	{
		FreeDefinitions(1, pItemDefs, pResults, pErrors);

	   	TRACE_METHOD_ERROR(AddItems, hResult);
		return hResult;
	}

	// check for invalid response from server.
	if (pResults == NULL || pErrors == NULL)
	{
		FreeDefinitions(1, pItemDefs, pResults, pErrors);

	   	TRACE_METHOD_ERROR(AddItems, E_POINTER);
		return E_POINTER;
	}

	// check for item level error.
	if (FAILED(pErrors[0]))
	{
		hResult = pErrors[0];
		FreeDefinitions(1, pItemDefs, pResults, pErrors);
        
	   	TRACE_ERROR("%s, 'AddItems Failed [0x%08X]'", TRACE_FUNCTION, hResult);
		return hResult;
	}

	// initialize new item object.
	COPCItem* pItem = (COPCItem*)pItemDefs->hClient;
	pItem->Initialize(ClientHandle, pItemDefs, pResults, m_pOPCGroup,  m_pParent);
	pItem->AddRef();  

	m_items.push_back(pItem);

	// free memory allocated for operation.
	FreeDefinitions(1, pItemDefs, pResults, pErrors);

	// set item active.
	if (bIsActive)
	{
		pItem->put_IsActive(TRUE);
	}

	// fetch the interface.
	pItem->QueryInterface(IID_OPCItem, (void**)ppItem);
	_ASSERTE(*ppItem != NULL);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// CreateDefinitions Method

OPCITEMDEF* COPCItems::CreateDefinitions(
	LONG       lCount,
	SAFEARRAY* psaItemIDs,
	SAFEARRAY* psaReqTypes,
	SAFEARRAY* psaAccessPaths
)
{
	OPCITEMDEF* pItemDefs = (OPCITEMDEF*)CoTaskMemAlloc(sizeof(OPCITEMDEF)*lCount);
	memset(pItemDefs, 0, sizeof(OPCITEMDEF)*lCount);

	for (LONG ii = 1; ii <= lCount; ii++)
	{		
		COPCItemObject* pItem = NULL;
		COPCItemObject::CreateInstance(&pItem);
		pItem->AddRef();

		BSTR szItemID = NULL;
		SafeArrayGetElement(psaItemIDs, &ii, (void*)&szItemID);

		// create item definition.
		pItemDefs[ii-1].szItemID            = OpcDupStr(szItemID);
		pItemDefs[ii-1].dwBlobSize          = 0;
		pItemDefs[ii-1].pBlob               = NULL;
		pItemDefs[ii-1].bActive             = FALSE;
		pItemDefs[ii-1].hClient             = (OPCHANDLE)pItem;
		pItemDefs[ii-1].vtRequestedDataType = m_defaultDataType;
		pItemDefs[ii-1].szAccessPath        = OpcDupStr(m_defaultAccessPath);

		SysFreeString(szItemID);

		if (psaReqTypes != NULL)
		{
			SHORT sReqType = 0;
			SafeArrayGetElement(psaReqTypes, &ii, (void*)&sReqType);
			pItemDefs[ii-1].vtRequestedDataType = sReqType;
		}

		if (psaAccessPaths != NULL)
		{
			BSTR szAccessPath = NULL;
			SafeArrayGetElement(psaAccessPaths, &ii, (void*)&szAccessPath);
			pItemDefs[ii-1].szAccessPath = OpcDupStr(szAccessPath);
			SysFreeString(szAccessPath);
		}
	}

	return pItemDefs;
}

//==============================================================================
// FreeDefinitions Method

void COPCItems::FreeDefinitions(
	LONG           lCount,
	OPCITEMDEF*    pItemDefs,
    OPCITEMRESULT* pItemResults,
    HRESULT*       pErrors
)
{
	for (LONG ii = 0; ii < lCount; ii++)
	{
		if (pItemDefs != NULL)
		{
			CoTaskMemFree(pItemDefs[ii].szItemID);
			CoTaskMemFree(pItemDefs[ii].szAccessPath);

			if (pItemDefs[ii].hClient != NULL)
			{
				COPCItem* pItem = (COPCItem*)pItemDefs[ii].hClient;
				pItem->Release();
			}

			if (pItemDefs[ii].dwBlobSize > 0)
			{
				CoTaskMemFree(pItemDefs[ii].pBlob);
			}
		}

		if (pItemResults != NULL)
		{
			if (pItemResults[ii].dwBlobSize > 0)
			{
				CoTaskMemFree(pItemResults[ii].pBlob);
			}
		}
	}

	CoTaskMemFree(pItemDefs);
	CoTaskMemFree(pItemResults);
	CoTaskMemFree(pErrors);
}

//==============================================================================
// AddItems Method

STDMETHODIMP COPCItems::AddItems( 
	LONG        NumItems,
	SAFEARRAY** ppItemIDs,
	SAFEARRAY** ppClientHandles,
	SAFEARRAY** ppServerHandles,
	SAFEARRAY** ppErrors,
	VARIANT     RequestedDataTypes, 
	VARIANT     AccessPaths
)
{
	TRACE_FUNCTION_ENTER("OPCItems::AddItems");

	// check arguments.
	if (!CheckByValInArg(NumItems))                                 return E_INVALIDARG;
	if (!CheckByRefInArrayArg(ppItemIDs, VT_BSTR, NumItems))        return E_INVALIDARG;
	if (!CheckByRefInArrayArg(ppClientHandles, VT_I4, NumItems))    return E_INVALIDARG;
	if (!CheckByRefOutArg(ppServerHandles))                         return E_INVALIDARG;
	if (!CheckByRefOutArg(ppErrors))                                return E_INVALIDARG;
	if (!CheckByValInArrayArg(RequestedDataTypes, VT_I2, NumItems)) return E_INVALIDARG;
	if (!CheckByValInArrayArg(AccessPaths, VT_BSTR, NumItems))      return E_INVALIDARG;

	*ppServerHandles = NULL;
	*ppErrors        = NULL;

	// check if interface is supported.
	if (!m_pOPCItem)
	{
	   	TRACE_INTERFACE_ERROR(IOPCItemMgt);
		return E_NOINTERFACE;
	}

	BOOL bIsActive = m_defaultActive;

	// create item definitions.
	OPCITEMDEF* pItemDefs = CreateDefinitions(
		NumItems, 
		*ppItemIDs, 
		(RequestedDataTypes.vt == (VT_I2 | VT_ARRAY))?RequestedDataTypes.parray:NULL,
		(AccessPaths.vt == (VT_BSTR | VT_ARRAY))?AccessPaths.parray:NULL
	);

	// add items to group.
    OPCITEMRESULT* pResults = NULL;
    HRESULT*       pErrors  = NULL;

	TRACE_METHOD_CALL(AddItems);

	HRESULT hResult = m_pOPCItem->AddItems(NumItems, pItemDefs, &pResults, &pErrors);

	// check method level error.
	if (FAILED(hResult))
	{
		FreeDefinitions(NumItems, pItemDefs, pResults, pErrors);

		TRACE_METHOD_ERROR(AddItems, hResult);
		return hResult;
	}

	// check for invalid response from server.
	if (pResults == NULL || pErrors == NULL)
	{
		FreeDefinitions(NumItems, pItemDefs, pResults, pErrors);

		TRACE_METHOD_ERROR(AddItems, E_POINTER);
		return E_POINTER;
	}

	// construct responses.
	SAFEARRAYBOUND cBound;
	cBound.lLbound   = 1;
	cBound.cElements = NumItems;

	*ppServerHandles = SafeArrayCreate(VT_I4, 1, &cBound);
	*ppErrors        = SafeArrayCreate(VT_I4, 1, &cBound);

	for (LONG ii = 1; ii <= NumItems; ii++)
	{
		OPCHANDLE hServer = NULL;

		COPCItem* pItem = (COPCItem*)pItemDefs[ii-1].hClient;

		if (SUCCEEDED(pErrors[ii-1]))
		{	
			// get the local client handle.
			OPCHANDLE hClient = NULL;
			SafeArrayGetElement(*ppClientHandles, &ii, &hClient);

			pItem->Initialize(hClient, &pItemDefs[ii-1], &pResults[ii-1], m_pOPCGroup,  m_pParent);
			pItem->AddRef();  // hold our reference.
			
			m_items.push_back(pItem);

			hServer = (OPCHANDLE)pItem;
		}

		SafeArrayPutElement(*ppServerHandles, &ii, (void*)&hServer);
		SafeArrayPutElement(*ppErrors, &ii, (void*)&pErrors[ii-1]);
	}	
	
	// activate items now that eveything is set up properly. 
	if (bIsActive)
	{
		// build list of valid handles.
		OPCHANDLE* phHandles = (OPCHANDLE*)CoTaskMemAlloc(sizeof(OPCHANDLE)*NumItems);
		COPCItem** ppItems   = (COPCItem**)CoTaskMemAlloc(sizeof(COPCItem*)*NumItems);

		memset(phHandles, 0, sizeof(OPCHANDLE)*NumItems);
		memset(ppItems, 0, sizeof(COPCItem*)*NumItems);

		LONG lIndex = 0;

		for (ii = 0; ii < NumItems; ii++)
		{
			if (SUCCEEDED(pErrors[ii]))
			{
				phHandles[lIndex] = pResults[ii].hServer;
				ppItems[lIndex]   = (COPCItem*)pItemDefs[ii].hClient;
				
				lIndex++;
			}
		}

		// free old errors array.
		CoTaskMemFree(pErrors);
		pErrors = NULL;

		// ignore errors from this call since the items already exist - they just won't be active.
		TRACE_METHOD_CALL(SetActiveState);

		hResult = m_pOPCItem->SetActiveState(lIndex, phHandles , TRUE, &pErrors);

		if (SUCCEEDED(hResult))
		{
			// check item level errors.
			for (ii = 0; ii < lIndex; ii++)
			{
				if (SUCCEEDED(pErrors[ii]))
				{
					ppItems[ii]->SetActiveState(TRUE);
				}
			}
		}
		else
		{
			TRACE_ERROR("%s, Could not activate items after adding them to a group.", TRACE_FUNCTION);
		}
	
		CoTaskMemFree(phHandles);
		CoTaskMemFree(ppItems);
	}

	// free memory allocated for call.
	FreeDefinitions(NumItems, pItemDefs, pResults, pErrors);

	TraceByRefOutArg(ppServerHandles);
	TraceByRefOutArg(ppErrors);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// Validate Method

STDMETHODIMP COPCItems::Validate(
	LONG        NumItems,
	SAFEARRAY** ppItemIDs,
	SAFEARRAY** ppErrors,
	VARIANT     RequestedDataTypes, 
	VARIANT     AccessPaths
)
{
	TRACE_FUNCTION_ENTER("OPCItems::Validate");

	// check arguments.
	if (!CheckByValInArg(NumItems))                                 return E_INVALIDARG;
	if (!CheckByRefInArrayArg(ppItemIDs, VT_BSTR, NumItems))        return E_INVALIDARG;
	if (!CheckByRefOutArg(ppErrors))                                return E_INVALIDARG;
	if (!CheckByValInArrayArg(RequestedDataTypes, VT_I2, NumItems)) return E_INVALIDARG;
	if (!CheckByValInArrayArg(AccessPaths, VT_BSTR, NumItems))      return E_INVALIDARG;

	*ppErrors = NULL;

	// check if interface is supported.
	if (!m_pOPCItem)
	{
	   	TRACE_INTERFACE_ERROR(IOPCItemMgt);
		return E_NOINTERFACE;
	}

	// create item definitions.
	OPCITEMDEF* pItemDefs = CreateDefinitions(
		NumItems, 
		*ppItemIDs, 
		(RequestedDataTypes.vt == (VT_I2 | VT_ARRAY))?RequestedDataTypes.parray:NULL,
		(AccessPaths.vt == (VT_BSTR | VT_ARRAY))?AccessPaths.parray:NULL
	);
	
	// add items to group.
    OPCITEMRESULT* pResults = NULL;
    HRESULT*       pErrors  = NULL;
		
	TRACE_METHOD_CALL(ValidateItems);

	HRESULT hResult = m_pOPCItem->ValidateItems(NumItems, pItemDefs, FALSE, &pResults, &pErrors);

	// check method level error.
	if (FAILED(hResult))
	{
		FreeDefinitions(NumItems, pItemDefs, pResults, pErrors);

		TRACE_METHOD_ERROR(ValidateItems, hResult);
		return hResult;
	}

	// check for invalid response from server.
	if (pResults == NULL || pErrors == NULL)
	{
		FreeDefinitions(NumItems, pItemDefs, pResults, pErrors);

		TRACE_METHOD_ERROR(ValidateItems, E_POINTER);
		return E_POINTER;
	}

	// construct error array.
	SAFEARRAYBOUND cBound;
	cBound.lLbound   = 1;
	cBound.cElements = NumItems;

	*ppErrors = SafeArrayCreate(VT_I4, 1, &cBound);

	for (LONG ii = 1; ii <= NumItems; ii++)
	{
		SafeArrayPutElement(*ppErrors, &ii, (void*)&pErrors[ii-1]);
	}	
	
	// free memory allocated for call.
	FreeDefinitions(NumItems, pItemDefs, pResults, pErrors);

	TraceByRefOutArg(ppErrors);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// Remove Method

STDMETHODIMP COPCItems::Remove(
	LONG        NumItems,
	SAFEARRAY** ppServerHandles,
	SAFEARRAY** ppErrors
)
{
	TRACE_FUNCTION_ENTER("OPCItems::Remove");

	// check arguments.
	if (!CheckByValInArg(NumItems))                                 return E_INVALIDARG;
	if (!CheckByRefInArrayArg(ppServerHandles, VT_I4, NumItems))    return E_INVALIDARG;
	if (!CheckByRefOutArg(ppErrors))                                return E_INVALIDARG;

	*ppErrors = NULL;

	// check if interface is supported.
	if (!m_pOPCItem)
	{
	   	TRACE_ERROR("%s, '!m_pOPCItem'", TRACE_FUNCTION);
		return E_NOINTERFACE;
	}

	// construct list of server handles.
	OPCHANDLE hInvalidHandle = GetInvalidHandle();

	OPCHANDLE* phHandles = (OPCHANDLE*)CoTaskMemAlloc(sizeof(OPCHANDLE)*NumItems);

	for (LONG ii = 1; ii <= NumItems; ii++)
	{
		phHandles[ii-1] = hInvalidHandle;

		OPCHANDLE hServer = NULL;
		SafeArrayGetElement(*ppServerHandles, &ii, &hServer);

		COPCItem* pItem = LookupItem(hServer);

		if (pItem != NULL)
		{
			phHandles[ii-1] = pItem->GetServerHandle();

			// always remove the local item object - even if an error occurs. 
			m_items.remove(pItem);
			pItem->Release();
		}
	}

	// remove the items.
	HRESULT* pErrors = NULL;

	TRACE_METHOD_CALL(RemoveItems);

	HRESULT hResult = m_pOPCItem->RemoveItems(NumItems, phHandles, &pErrors);

	CoTaskMemFree(phHandles);

	if (FAILED(hResult))
	{
		TRACE_METHOD_ERROR(RemoveItems, hResult);
		return hResult;
	}

	if (pErrors == NULL)
	{
		TRACE_METHOD_ERROR(RemoveItems, E_POINTER);
		return E_POINTER;
	}
  
	// construct error array.
	SAFEARRAYBOUND cBound;
	cBound.lLbound   = 1;
	cBound.cElements = NumItems;

	*ppErrors = SafeArrayCreate(VT_I4, 1, &cBound);

	for (ii = 1; ii <= NumItems; ii++)
	{
		SafeArrayPutElement(*ppErrors, &ii, &pErrors[ii-1]);
	}

	CoTaskMemFree(pErrors);

	TraceByRefOutArg(ppErrors);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// SetActive Method

STDMETHODIMP COPCItems::SetActive(
	LONG         NumItems,
	SAFEARRAY**  ppServerHandles,
	VARIANT_BOOL ActiveState,
	SAFEARRAY**  ppErrors
)
{
	TRACE_FUNCTION_ENTER("OPCItems::SetActive");

	// check arguments.
	if (!CheckByValInArg(NumItems))                              return E_INVALIDARG;
	if (!CheckByRefInArrayArg(ppServerHandles, VT_I4, NumItems)) return E_INVALIDARG;
	if (!CheckByValInArg(ActiveState))                           return E_INVALIDARG;
	if (!CheckByRefOutArg(ppErrors))                             return E_INVALIDARG;

	*ppErrors = NULL;

	// check if interface is supported.
	if (!m_pOPCItem)
	{
	   	TRACE_INTERFACE_ERROR(IOPCItemMgt);
		return E_NOINTERFACE;
	}

	// construct list of server handles.
	OPCHANDLE hInvalidHandle = GetInvalidHandle();

	OPCHANDLE* phHandles = (OPCHANDLE*)CoTaskMemAlloc(sizeof(OPCHANDLE)*NumItems);
	COPCItem** ppItems   = (COPCItem**)CoTaskMemAlloc(sizeof(COPCItem*)*NumItems);

	for (LONG ii = 1; ii <= NumItems; ii++)
	{
		phHandles[ii-1] = hInvalidHandle;
		ppItems[ii-1]   = NULL;

		OPCHANDLE hServer = NULL;
		SafeArrayGetElement(*ppServerHandles, &ii, &hServer);

		COPCItem* pItem = LookupItem(hServer);

		if (pItem != NULL)
		{
			phHandles[ii-1] = pItem->GetServerHandle();
			ppItems[ii-1]   = pItem;
		}
	}

	// update the items on the server.
	HRESULT* pErrors = NULL;

	TRACE_METHOD_CALL(SetActiveState);

	HRESULT hResult = m_pOPCItem->SetActiveState(NumItems, phHandles, (ActiveState)?TRUE:FALSE, &pErrors);

	CoTaskMemFree(phHandles);

	if (FAILED(hResult))
	{
		CoTaskMemFree(ppItems);

		TRACE_METHOD_ERROR(SetActiveState, hResult);
		return hResult;
	}

	if (pErrors == NULL)
	{
		CoTaskMemFree(ppItems);

		TRACE_METHOD_ERROR(SetActiveState, E_POINTER);
		return E_POINTER;
	}
  
	// construct error array.
	SAFEARRAYBOUND cBound;
	cBound.lLbound   = 1;
	cBound.cElements = NumItems;

	*ppErrors = SafeArrayCreate(VT_I4, 1, &cBound);

	for (ii = 1; ii <= NumItems; ii++)
	{
		// override error code for invalid items.
		if (ppItems[ii-1] == NULL)
		{
			pErrors[ii-1] = OPC_E_INVALIDHANDLE;
		}

		// update item object.
		if (SUCCEEDED(pErrors[ii-1]))
		{
			ppItems[ii-1]->SetActiveState((ActiveState)?TRUE:FALSE);
		}

		SafeArrayPutElement(*ppErrors, &ii, &pErrors[ii-1]);
	}

	CoTaskMemFree(ppItems);
	CoTaskMemFree(pErrors);

	TraceByRefOutArg(ppErrors);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// SetClientHandles Method

STDMETHODIMP COPCItems::SetClientHandles(
	LONG        NumItems,
	SAFEARRAY** ppServerHandles,
	SAFEARRAY** ppClientHandles,
	SAFEARRAY** ppErrors
)
{
	TRACE_FUNCTION_ENTER("OPCItems::SetClientHandles");

	// check arguments.
	if (!CheckByValInArg(NumItems))                              return E_INVALIDARG;
	if (!CheckByRefInArrayArg(ppServerHandles, VT_I4, NumItems)) return E_INVALIDARG;
	if (!CheckByRefInArrayArg(ppClientHandles, VT_I4, NumItems)) return E_INVALIDARG;
	if (!CheckByRefOutArg(ppErrors))                             return E_INVALIDARG;

	// construct error array.
	SAFEARRAYBOUND cBound;
	cBound.lLbound   = 1;
	cBound.cElements = NumItems;

	*ppErrors = SafeArrayCreate(VT_I4, 1, &cBound);

	// update client handles.
	for (LONG ii = 1; ii <= NumItems; ii++)
	{
		OPCHANDLE hServer = NULL;
		OPCHANDLE hClient = NULL;
		HRESULT   hError  = OPC_E_INVALIDHANDLE;

		SafeArrayGetElement(*ppServerHandles, &ii, &hServer);
		SafeArrayGetElement(*ppClientHandles, &ii, &hClient);

		COPCItem* pItem = LookupItem(hServer);

		if (pItem != NULL)
		{
			pItem->SetClientHandle(hClient);
			hError = S_OK;
		}

		SafeArrayPutElement(*ppErrors, &ii, &hError);
	}

	TraceByRefOutArg(ppErrors);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// SetDataTypes Method

STDMETHODIMP COPCItems::SetDataTypes(
	LONG        NumItems,
	SAFEARRAY** ppServerHandles,
	SAFEARRAY** ppRequestedDataTypes,
	SAFEARRAY** ppErrors
)
{
	TRACE_FUNCTION_ENTER("OPCItems::SetDataTypes");

	// check arguments.
	if (!CheckByValInArg(NumItems))                                   return E_INVALIDARG;
	if (!CheckByRefInArrayArg(ppServerHandles, VT_I4, NumItems))      return E_INVALIDARG;
	if (!CheckByRefInArrayArg(ppRequestedDataTypes, VT_I2, NumItems)) return E_INVALIDARG;
	if (!CheckByRefOutArg(ppErrors))                                  return E_INVALIDARG;

	*ppErrors = NULL;

	// check if interface is supported.
	if (!m_pOPCItem)
	{
		TRACE_INTERFACE_ERROR(IOPCItemMgt);
		return E_NOINTERFACE;
	}

	// construct list of server handles.
	OPCHANDLE hInvalidHandle = GetInvalidHandle();

	OPCHANDLE* phHandles = (OPCHANDLE*)CoTaskMemAlloc(sizeof(OPCHANDLE)*NumItems);
	VARTYPE*   pReqTypes = (VARTYPE*)CoTaskMemAlloc(sizeof(VARTYPE)*NumItems);
	COPCItem** ppItems   = (COPCItem**)CoTaskMemAlloc(sizeof(COPCItem*)*NumItems);

	for (LONG ii = 1; ii <= NumItems; ii++)
	{
		phHandles[ii-1] = hInvalidHandle;
		pReqTypes[ii-1] = VT_EMPTY;
		ppItems[ii-1]   = NULL;

		OPCHANDLE hServer  = NULL;
		LONG      lReqType = VT_EMPTY;

		SafeArrayGetElement(*ppServerHandles, &ii, &hServer);
		SafeArrayGetElement(*ppRequestedDataTypes, &ii, &lReqType);

		COPCItem* pItem = LookupItem(hServer);

		if (pItem != NULL)
		{
			phHandles[ii-1] = pItem->GetServerHandle();
			pReqTypes[ii-1] = (VARTYPE)lReqType;
			ppItems[ii-1]   = pItem;
		}
	}

	// update the items on the server.
	HRESULT* pErrors = NULL;

	TRACE_METHOD_CALL(SetDatatypes);

	HRESULT hResult = m_pOPCItem->SetDatatypes(NumItems, phHandles, pReqTypes, &pErrors);

	CoTaskMemFree(phHandles);

	if (FAILED(hResult))
	{
		CoTaskMemFree(pReqTypes);
		CoTaskMemFree(ppItems);

		TRACE_METHOD_ERROR(SetDatatypes, hResult);
		return hResult;
	}

	if (pErrors == NULL)
	{
		CoTaskMemFree(pReqTypes);
		CoTaskMemFree(ppItems);

		TRACE_METHOD_ERROR(SetDatatypes, E_POINTER);
		return E_POINTER;
	}
  
	// construct error array.
	SAFEARRAYBOUND cBound;
	cBound.lLbound   = 1;
	cBound.cElements = NumItems;

	*ppErrors = SafeArrayCreate(VT_I4, 1, &cBound);

	for (ii = 1; ii <= NumItems; ii++)
	{
		// override error code for invalid items.
		if (ppItems[ii-1] == NULL)
		{
			pErrors[ii-1] = OPC_E_INVALIDHANDLE;
		}

		// update item object.
		if (SUCCEEDED(pErrors[ii-1]))
		{
			ppItems[ii-1]->SetDataType(pReqTypes[ii-1]);
		}

		SafeArrayPutElement(*ppErrors, &ii, &pErrors[ii-1]);
	}

	CoTaskMemFree(pReqTypes);
	CoTaskMemFree(ppItems);
	CoTaskMemFree(pErrors);

	TraceByRefOutArg(ppErrors);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}
