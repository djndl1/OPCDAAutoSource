//==============================================================================
// TITLE: OPCGroups.cpp
//
// CONTENTS:
//
//  This is the Server Object's OPCGroups interface
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
//

#include "StdAfx.h"
#include "OPCGroups.h"
#include "OPCAutoServer.h"
#include "OPCUtils.h"
#include "OPCTrace.h"

//==============================================================================
// Constructor

COPCGroups::COPCGroups()
{
	m_pParent         = NULL;
	m_defaultActive   = VARIANT_TRUE;
	m_defaultUpdate   = 1000;
	m_defaultDeadband = 0.0F;
	m_defaultLocale   = LOCALE_USER_DEFAULT;
	m_defaultTimeBias = 0;
}

//==============================================================================
// Destructor

COPCGroups::~COPCGroups()
{
}

//==============================================================================
// Initialize

void COPCGroups::Initialize(COPCAutoServer* pParent)
{
	m_pParent = pParent;
}

//==============================================================================
// Connect

void COPCGroups::Connect(IUnknown* ipUnknown)
{
	m_pOPCServer = ipUnknown;
}

//==============================================================================
// Disconnect

void COPCGroups::Disconnect()
{
	RemoveAll();
	m_pOPCServer.Release();
}

//==============================================================================
// get_Parent Method

STDMETHODIMP COPCGroups::get_Parent(IOPCAutoServer** ppParent)
{
	TRACE_FUNCTION_ENTER("COPCGroups::get_Parent");

	// check arguments.
	if (!CheckByRefOutArg(ppParent)) return E_INVALIDARG;

	*ppParent = NULL;

	// get Interface pointer
	((COPCAutoServerObject*)m_pParent)->QueryInterface(IID_IOPCAutoServer, (void**)ppParent);
	_ASSERTE(*ppParent);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// DefaultGroupIsActive Property

STDMETHODIMP COPCGroups::get_DefaultGroupIsActive(VARIANT_BOOL* pDefaultGroupIsActive)
{
	TRACE_FUNCTION_ENTER("COPCGroups::get_DefaultGroupIsActive");

	// check arguments.
	if (!CheckByRefOutArg(pDefaultGroupIsActive)) return E_INVALIDARG;

	*pDefaultGroupIsActive = m_defaultActive;

	TraceByRefOutArg(pDefaultGroupIsActive);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

STDMETHODIMP COPCGroups::put_DefaultGroupIsActive(VARIANT_BOOL DefaultGroupIsActive)
{
	TRACE_FUNCTION_ENTER("COPCGroups::put_DefaultGroupIsActive");

	// check arguments.
	if (!CheckByValInArg(DefaultGroupIsActive)) return E_INVALIDARG;

	m_defaultActive = DefaultGroupIsActive;
	
	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// DefaultGroupUpdateRate Property

STDMETHODIMP COPCGroups::get_DefaultGroupUpdateRate(LONG* pDefaultGroupUpdateRate)
{
	TRACE_FUNCTION_ENTER("COPCGroups::get_DefaultGroupUpdateRate");

	// check arguments.
	if (!CheckByRefOutArg(pDefaultGroupUpdateRate)) return E_INVALIDARG;

	*pDefaultGroupUpdateRate = m_defaultUpdate;

	TraceByRefOutArg(pDefaultGroupUpdateRate);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

STDMETHODIMP COPCGroups::put_DefaultGroupUpdateRate(LONG DefaultGroupUpdateRate)
{
	TRACE_FUNCTION_ENTER("COPCGroups::put_DefaultGroupUpdateRate");

	// check arguments.
	if (!CheckByValInArg(DefaultGroupUpdateRate)) return E_INVALIDARG;

	m_defaultUpdate = DefaultGroupUpdateRate;

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// DefaultGroupDeadband Property

STDMETHODIMP COPCGroups::get_DefaultGroupDeadband(float* pDefaultGroupDeadband)
{
	TRACE_FUNCTION_ENTER("COPCGroups::get_DefaultGroupDeadband");

	// check arguments.
	if (!CheckByRefOutArg(pDefaultGroupDeadband)) return E_INVALIDARG;

	*pDefaultGroupDeadband = m_defaultDeadband;

	TraceByRefOutArg(pDefaultGroupDeadband);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

STDMETHODIMP COPCGroups::put_DefaultGroupDeadband(float DefaultGroupDeadband)
{
	TRACE_FUNCTION_ENTER("COPCGroups::put_DefaultGroupDeadband");

	// check arguments.
	if (!CheckByValInArg(DefaultGroupDeadband)) return E_INVALIDARG;

	m_defaultDeadband = DefaultGroupDeadband;

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// DefaultGroupLocaleID Property

STDMETHODIMP COPCGroups::get_DefaultGroupLocaleID(LONG* pDefaultGroupLocaleID)
{
	TRACE_FUNCTION_ENTER("COPCGroups::get_DefaultGroupLocaleID");

	// check arguments.
	if (!CheckByRefOutArg(pDefaultGroupLocaleID)) return E_INVALIDARG;

	*pDefaultGroupLocaleID = m_defaultLocale;

	TraceByRefOutArg(pDefaultGroupLocaleID);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

STDMETHODIMP COPCGroups::put_DefaultGroupLocaleID(LONG DefaultGroupLocaleID)
{
	TRACE_FUNCTION_ENTER("COPCGroups::put_DefaultGroupLocaleID");

	// check arguments.
	if (!CheckByValInArg(DefaultGroupLocaleID)) return E_INVALIDARG;

	m_defaultLocale = DefaultGroupLocaleID;

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// DefaultGroupTimeBias Property

STDMETHODIMP COPCGroups::get_DefaultGroupTimeBias(LONG* pDefaultGroupTimeBias)
{
	TRACE_FUNCTION_ENTER("COPCGroups::get_DefaultGroupTimeBias");

	// check arguments.
	if (!CheckByRefOutArg(pDefaultGroupTimeBias)) return E_INVALIDARG;

	*pDefaultGroupTimeBias = m_defaultTimeBias;

	TraceByRefOutArg(pDefaultGroupTimeBias);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

STDMETHODIMP COPCGroups::put_DefaultGroupTimeBias(LONG DefaultGroupTimeBias)
{
	TRACE_FUNCTION_ENTER("COPCGroups::put_DefaultGroupTimeBias");

	// check arguments.
	if (!CheckByValInArg(DefaultGroupTimeBias)) return E_INVALIDARG;

	m_defaultTimeBias = DefaultGroupTimeBias;

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// Count Property

STDMETHODIMP COPCGroups::get_Count(LONG* pCount)
{
	TRACE_FUNCTION_ENTER("COPCGroups::get_Count");
	
	// check arguments.
	if (!CheckByRefOutArg(pCount)) return E_INVALIDARG;

	*pCount = m_groups.size();

	TraceByRefOutArg(pCount);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// _NewEnum Property

STDMETHODIMP COPCGroups::get__NewEnum(IUnknown** ppUnk)
{
	TRACE_FUNCTION_ENTER("COPCGroups::get__NewEnum");

	// check arguments.
	if (!CheckByRefOutArg(ppUnk)) return E_INVALIDARG;

	EnumVARIANT* pEnumerator = new EnumVARIANT();

	HRESULT hResult = S_OK;

	// create an empty IEnumVariant object.
	if (m_groups.size() == 0)
	{
		CComVariant varEmpty(_T(""));
		hResult = pEnumerator->Init(&varEmpty, &varEmpty, NULL, AtlFlagCopy);
	}
		
	// create an array of variants containing IDispatch pointers.
	else
	{
		VARIANT* pArray = (VARIANT*)malloc(m_groups.size()*sizeof(VARIANT));

		list<COPCGroup*>::iterator iter = m_groups.begin();
		
		for (LONG ii = 0; iter != m_groups.end(); ii++)
		{
			COPCGroup* pGroup = *iter;

			// fetch pointer to group.
			IOPCGroup* ipGroup = NULL;
			pGroup->QueryInterface(IID_IOPCGroup, (void**)&ipGroup);
			_ASSERTE(ipGroup);
			
			// and put it into a variant
			VariantInit(&pArray[ii]);
			pArray[ii].vt       = VT_DISPATCH;
			pArray[ii].pdispVal = ipGroup;
			
			iter++;
		}

		// create an IEnumVariant object initialized with these variants
		hResult = pEnumerator->Init(&pArray[0], &pArray[m_groups.size()], NULL, AtlFlagCopy);
		
		// clear the local variant array.
		for (ii = 0; ii < (LONG)m_groups.size(); ii++)
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

STDMETHODIMP COPCGroups::Item(VARIANT ItemSpecifier, OPCGroup** ppGroup)
{
	TRACE_FUNCTION_ENTER("COPCGroups::Item");

	// check arguments.
	if (!CheckByValInVariantArg(ItemSpecifier, VT_VARIANT, false)) return E_INVALIDARG;
	if (!CheckByRefOutArg(ppGroup))                                return E_INVALIDARG;

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

	// find matching group.
	list<COPCGroup*>::iterator iter = m_groups.begin();
	
	for (LONG ii = 0; iter != m_groups.end(); ii++)
	{
		COPCGroup* pGroup = *iter;

		if (((bstrName != NULL) && (wcscmp(pGroup->GetName(), bstrName) == 0)) || ((bstrName == NULL) && (ii+1 == lIndex)))
		{
			pGroup->QueryInterface(IID_IOPCGroup, (void**)ppGroup);
			_ASSERTE(*ppGroup);

			TRACE_FUNCTION_LEAVE();
			return S_OK;
		}

		iter++;
	}

	TRACE_ERROR("%s, 'E_INVALIDARG'", TRACE_FUNCTION);
	return E_INVALIDARG;
}

//==============================================================================
// Add Method

STDMETHODIMP COPCGroups::Add(VARIANT Name, OPCGroup** ppGroup)
{
	TRACE_FUNCTION_ENTER("COPCGroups::Add");
	
	// check arguments.
	if (!CheckByValInVariantArg(Name, VT_BSTR, true)) return E_INVALIDARG;
	if (!CheckByRefOutArg(ppGroup))                   return E_INVALIDARG;

	*ppGroup = NULL;

	// check if server is connected.
	if (!m_pOPCServer)
	{
	   	TRACE_INTERFACE_ERROR(IOPCServer);
		return E_FAIL;
	}

	// create an Automation group object
	COPCGroupObject* pGroup = NULL;
	
	HRESULT hResult = COPCGroupObject::CreateInstance(&pGroup);
	
	if (FAILED(hResult))
	{
	   	TRACE_ERROR("%s, 'CreateInstance failed - hResult=[0x%08X]'", TRACE_FUNCTION, hResult);
		return hResult;
	}

	// create a group in the server
	CComBSTR bstrName(L"");

	if (Name.scode != DISP_E_PARAMNOTFOUND && Name.vt == VT_BSTR && Name.bstrVal != NULL)
	{
		bstrName = Name.bstrVal;
	}

	OPCHANDLE hServer             = NULL;
	DWORD     dwRevisedUpdateRate = 0;
	IUnknown* ipUnknown           = NULL;

	TRACE_METHOD_CALL(AddGroup);

	hResult = m_pOPCServer->AddGroup(  
		bstrName,
		m_defaultActive,
		m_defaultUpdate,
		(OPCHANDLE)pGroup, // local group pointer is client handle
		&m_defaultTimeBias,
		&m_defaultDeadband,
		m_defaultLocale,
		&hServer,
		&dwRevisedUpdateRate,
		IID_IOPCGroupStateMgt,
		&ipUnknown 
	);
	
	if (FAILED(hResult))
	{
		delete pGroup;
		TRACE_METHOD_ERROR(AddGroup, hResult);
		return hResult;
	}

	// initialize group.
	hResult = pGroup->Initialize(this, ipUnknown);
	ipUnknown->Release();
	ipUnknown = NULL;

	if (FAILED(hResult))
	{
		delete pGroup;
	   	TRACE_ERROR("%s, 'pGroup->Initialize failed - hResult=[0x%08X]'", TRACE_FUNCTION, hResult);
		return hResult;
	}

	// save group.
	m_groups.push_back(pGroup);
	pGroup->AddRef(); // save reference

	// get automation interface pointer
	pGroup->QueryInterface(IID_IOPCGroup, (void**)ppGroup);
	_ASSERTE(*ppGroup);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// GetOPCGroup Method

STDMETHODIMP COPCGroups::GetOPCGroup(VARIANT ItemSpecifier, OPCGroup** ppGroup)
{
	TRACE_FUNCTION_ENTER("COPCGroups::GetOPCGroup");

	// check arguments.
	if (!CheckByValInVariantArg(ItemSpecifier, VT_VARIANT, false)) return E_INVALIDARG;
	if (!CheckByRefOutArg(ppGroup))                                return E_INVALIDARG;

	// get indexer.
	OPCHANDLE hServer  = NULL;
	BSTR      bstrName = NULL;

	switch (ItemSpecifier.vt)
	{
		case VT_I4:	  { hServer  = ItemSpecifier.lVal;    break; }
		case VT_BSTR: {	bstrName = ItemSpecifier.bstrVal; break; }

		default:
		{
	   		TRACE_ERROR("%s, 'ItemSpecifier is not valid.'", TRACE_FUNCTION);
			return E_INVALIDARG;
		}
	}

	// find matching group.
	list<COPCGroup*>::iterator iter = m_groups.begin();
	
	for (LONG ii = 0; iter != m_groups.end(); ii++)
	{
		COPCGroup* pGroup = *iter;

		if (((bstrName != NULL) && (wcscmp(pGroup->GetName(), bstrName) == 0)) || ((bstrName == NULL) && (hServer == (OPCHANDLE)pGroup)))
		{
			pGroup->QueryInterface(IID_IOPCGroup, (void**)ppGroup);
			_ASSERTE(*ppGroup);

			TRACE_FUNCTION_LEAVE();
			return S_OK;
		}

		iter++;
	}

	TRACE_ERROR("%s, 'E_INVALIDARG'", TRACE_FUNCTION);
	return E_INVALIDARG;
}

//==============================================================================
// RemoveAll Method

STDMETHODIMP COPCGroups::RemoveAll()
{
	TRACE_FUNCTION_ENTER("COPCGroups::RemoveAll");

	// check if server is connected.
	if (!m_pOPCServer)
	{
	   	TRACE_INTERFACE_ERROR(IOPCServer);
		return E_FAIL;
	}

	while (m_groups.size() > 0)
	{
		COPCGroup* pGroup = *m_groups.begin();
		_ASSERTE(pGroup != NULL);
	
		// pass the server handle to the remove method.
		CComVariant varHandle((LONG)pGroup); 
		Remove(varHandle);
	}

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// Remove Method

STDMETHODIMP COPCGroups::Remove(VARIANT ItemSpecifier)
{
	USES_CONVERSION;

	TRACE_FUNCTION_ENTER("COPCGroups::Remove");

	// check arguments.
	if (!CheckByValInVariantArg(ItemSpecifier, VT_VARIANT, false)) return E_INVALIDARG;

	// get indexer.
	OPCHANDLE hServer  = NULL;
	BSTR      bstrName = NULL;

	switch (ItemSpecifier.vt)
	{
		case VT_I4:	  { hServer  = ItemSpecifier.lVal;    break; }
		case VT_BSTR: {	bstrName = ItemSpecifier.bstrVal; break; }

		default:
		{
	   		TRACE_ERROR("%s, 'ItemSpecifier is not valid.'", TRACE_FUNCTION);
			return E_INVALIDARG;
		}
	}

	// find matching group.
	list<COPCGroup*>::iterator iter = m_groups.begin();
	
	for (LONG ii = 0; iter != m_groups.end(); ii++)
	{
		COPCGroup* pGroup = *iter;

		if (((bstrName != NULL) && (wcscmp(pGroup->GetName(), bstrName) == 0)) || ((bstrName == NULL) && (hServer == (OPCHANDLE)pGroup)))
		{
			pGroup->put_IsSubscribed(VARIANT_FALSE);
			
			TRACE_METHOD_CALL(RemoveGroup);

			HRESULT hResult = m_pOPCServer->RemoveGroup(pGroup->GetServerHandle(), FALSE);

			if (FAILED(hResult))
			{
				TRACE_METHOD_ERROR(RemoveGroup, hResult);
			}

			m_groups.erase(iter);
			pGroup->Release();

			TRACE_FUNCTION_LEAVE();
			return hResult;
		}

		iter++;
	}

   	TRACE_ERROR("%s, 'E_INVALIDARG'", TRACE_FUNCTION);
	return E_INVALIDARG;
}

//==============================================================================
// ConnectPublicGroup Method

STDMETHODIMP COPCGroups::ConnectPublicGroup(BSTR Name, OPCGroup** ppGroup)
{
	TRACE_FUNCTION_NAME("COPCGroups::ConnectPublicGroup");

	// check arguments.
	if (!CheckByValInArg(Name))     return E_INVALIDARG;
	if (!CheckByRefOutArg(ppGroup)) return E_INVALIDARG;

    *ppGroup = NULL;

	TRACE_ERROR("%s, 'E_NOTIMPL'", TRACE_FUNCTION);
	return E_NOTIMPL;
}

//==============================================================================
// RemovePublicGroup Method

STDMETHODIMP COPCGroups::RemovePublicGroup(VARIANT ItemSpecifier)
{
	TRACE_FUNCTION_NAME("COPCGroups::RemovePublicGroup");

	// check arguments.
	if (!CheckByValInVariantArg(ItemSpecifier, VT_VARIANT, false)) return E_INVALIDARG;

	TRACE_ERROR("%s, 'E_NOTIMPL'", TRACE_FUNCTION);
	return E_NOTIMPL;
}