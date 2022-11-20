//==============================================================================
// TITLE: OPGroup.cpp
//
// CONTENTS:
//
//  This is the Group Object's OPCGroup interface.
//
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
#include "OPCGroup.h"

#include "OPCGroups.h"
#include "OPCAutoServer.h"
#include "OPCItems.h"

#include "OPCUtils.h"
#include "OPCTrace.h"

extern UINT OPCSTMFORMATDATATIME;
extern UINT OPCSTMFORMATWRITECOMPLETE;

//==============================================================================
// Constructor

COPCGroup::COPCGroup()
{
	m_pParent    = NULL;
	m_rate       = 0;
	m_actualRate = 0;
	m_active     = FALSE;
	m_timebias   = 0;
	m_deadband   = 0.0F;
	m_LCID       = LOCALE_SYSTEM_DEFAULT;
	m_client     = (OPCHANDLE)this;
	m_server     = 0;

	m_DataConnection  = 0;
	m_WriteConnection = 0;
	m_usingCP         = FALSE;
	m_subscribed      = VARIANT_FALSE;
	m_asyncReading    = FALSE;
  
	m_pItems = new COPCItemsObject();
	m_pItems->Initialize(this);
	m_pItems->AddRef();
}

//==============================================================================
// Destructor

COPCGroup::~COPCGroup()
{
	Uninitialize();
	m_pItems->Release();
}

//==============================================================================
// Initialize

HRESULT COPCGroup::Initialize(COPCGroups* pParent, IUnknown* pUnknown)
{
	TRACE_FUNCTION_ENTER("COPCGroup::Initialize");
	
	// check arguments.
	if (!CheckByRefInArg(pParent))  return E_INVALIDARG; 
	if (!CheckByRefInArg(pUnknown)) return E_INVALIDARG; 

	// save a reference to the server object,
	m_pParent = pParent;
    ((COPCAutoServerObject*)m_pParent)->AddRef();

	// save reference to group interface pointer.
	m_pOPCGroup = pUnknown;  
	
	// cache references to various group interfaces.
	m_pSyncIO   = m_pOPCGroup;
	m_pAsyncIO2 = m_pOPCGroup;
	m_pAsyncIO  = m_pOPCGroup;

	// pass reference to items object.
	m_pItems->Connect(pUnknown);

	if (!m_pOPCGroup)
	{
		TRACE_INTERFACE_ERROR(IOPCGroupStateMgt);
		return E_INVALIDARG;
	}

	// get initial values for state variables
	LPWSTR    wszName = NULL;
	OPCHANDLE hClient = NULL;

	TRACE_METHOD_CALL(GetState);

	HRESULT hResult = m_pOPCGroup->GetState( 
		&m_actualRate, 
		&m_active, 
		&wszName, 
		&m_timebias,
		&m_deadband, 
		&m_LCID, 
		&hClient, 
		&m_server 
	);

	if (FAILED(hResult))
	{
		TRACE_METHOD_ERROR(GetState, hResult);
		return hResult;
	}

	_ASSERTE(m_client == (OPCHANDLE)this);

	m_name = wszName;
	m_rate = m_actualRate;

	CoTaskMemFree(wszName);

	// check if this server supports OPC 2.0 ConnectionPoints
	IConnectionPointContainer* ipCPC = NULL;

	TRACE_METHOD_CALL(QueryInterface);

	hResult = m_pOPCGroup->QueryInterface(IID_IConnectionPointContainer, (void**)&ipCPC);
	
	if (SUCCEEDED(hResult))
	{
		m_usingCP = TRUE; 
		ipCPC->Release();
	}

	/*
	// check if this server supports OPC 1.0 Data Object
	IDataObject* ipDataObject = NULL;

	TRACE_METHOD_CALL(QueryInterface);

	hResult = m_pOPCGroup->QueryInterface(IID_IDataObject, (void**)&ipDataObject);
	
	if (SUCCEEDED(hResult))
	{
		m_usingCP = FALSE; 
		ipDataObject->Release();
	}
	*/
	
	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// Uninitialize

void COPCGroup::Uninitialize()
{
	TRACE_FUNCTION_ENTER("COPCGroup::Uninitialize");

	// clear subscribed state.
	put_IsSubscribed(VARIANT_FALSE);

	// remove all items.
	m_pItems->Disconnect();

	m_pSyncIO.Release();
	m_pAsyncIO.Release();
	m_pAsyncIO2.Release();
	m_pIDataObject.Release();
	m_pOPCGroup.Release();

	m_name.Empty();

	if (m_pParent != NULL)
	{
		((COPCAutoServerObject*)m_pParent)->Release();
		m_pParent = NULL;
	}
	
	TRACE_FUNCTION_LEAVE();
}

//==============================================================================
// Parent Property

STDMETHODIMP COPCGroup::get_Parent(IOPCAutoServer** ppParent)
{
	TRACE_FUNCTION_ENTER("COPCGroup::get_Parent");

	// check if group has been removed.
	if (m_pParent == NULL) return E_FAIL;

	// check arguments.
	if (!CheckByRefOutArg(ppParent)) return E_INVALIDARG; 

	*ppParent = NULL;

	// get parent of parent.
	m_pParent->get_Parent(ppParent);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// Name Property

STDMETHODIMP COPCGroup::get_Name(BSTR* pName)
{
	TRACE_FUNCTION_ENTER("COPCGroup::get_Name");

	// check if group has been removed.
	if (m_pParent == NULL) return E_FAIL;

	// check arguments.
	if (!CheckByRefOutArg(pName)) return E_INVALIDARG; 

	*pName = NULL;

	// check if group is connected.
	if (!m_pOPCGroup)
	{
		TRACE_INTERFACE_ERROR(IOPCGroupStateMgt);
		return E_FAIL;
	}

	// fetch name from the server.
	LPWSTR    wszName = NULL;
	OPCHANDLE hClient = NULL;

	TRACE_METHOD_CALL(GetState);

	HRESULT hResult = m_pOPCGroup->GetState( 
		&m_actualRate, 
		&m_active, 
		&wszName, 
		&m_timebias,
		&m_deadband, 
		&m_LCID, 
		&hClient, 
		&m_server 
	);

	if (FAILED(hResult))
	{
		TRACE_METHOD_ERROR(GetState, hResult);
		return hResult;
	}

	// cache name locally and allocate return parameter.
	m_name = wszName;
	*pName = m_name.Copy();

	CoTaskMemFree(wszName);

	TraceByRefOutArg(pName);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

STDMETHODIMP COPCGroup::put_Name(BSTR Name)
{
	TRACE_FUNCTION_ENTER("COPCGroup::put_Name");

	// check if group has been removed.
	if (m_pParent == NULL) return E_FAIL;

	// check arguments.
	if (!CheckByValInArg(Name)) return E_INVALIDARG; 

	// check if group is connected.
	if (!m_pOPCGroup)
	{
		TRACE_INTERFACE_ERROR(IOPCGroupStateMgt);
		return E_FAIL;
	}

	// only update the server if the name is actually being changed.
	if (m_name != Name)
	{
		// set name in server.
		TRACE_METHOD_CALL(SetName);

		HRESULT hResult = m_pOPCGroup->SetName((LPWSTR)Name);

		if (FAILED(hResult))
		{
			TRACE_METHOD_ERROR(SetName, hResult);
			return hResult;
		}

		// update local copy.
		m_name = Name;
	}

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// IsPublic Property

STDMETHODIMP COPCGroup::get_IsPublic(VARIANT_BOOL* pIsPublic)
{
	TRACE_FUNCTION_ENTER("COPCGroup::get_IsPublic");

	// check if group has been removed.
	if (m_pParent == NULL) return E_FAIL;

	// check arguments.
	if (!CheckByRefOutArg(pIsPublic)) return E_INVALIDARG; 

	*pIsPublic = VARIANT_FALSE;

	TraceByRefOutArg(pIsPublic);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// IsActive Property

STDMETHODIMP COPCGroup::get_IsActive(VARIANT_BOOL* pIsActive)
{
	TRACE_FUNCTION_ENTER("COPCGroup::get_IsActive");

	// check if group has been removed.
	if (m_pParent == NULL) return E_FAIL;

	// check arguments.
	if (!CheckByRefOutArg(pIsActive)) return E_INVALIDARG; 

	*pIsActive = VARIANT_FALSE;

	// check if group is connected.
	if (!m_pOPCGroup)
	{
		TRACE_INTERFACE_ERROR(IOPCGroupStateMgt);
		return E_FAIL;
	}

	// fetch active state from the server.
	LPWSTR    wszName = NULL;
	OPCHANDLE hClient = NULL;

	TRACE_METHOD_CALL(GetState);

	HRESULT hResult = m_pOPCGroup->GetState( 
		&m_actualRate, 
		&m_active, 
		&wszName, 
		&m_timebias,
		&m_deadband, 
		&m_LCID, 
		&hClient, 
		&m_server 
	);

	if (FAILED(hResult))
	{
		TRACE_METHOD_ERROR(GetState, hResult);
		return hResult;
	}

	CoTaskMemFree(wszName);

	*pIsActive = (m_active)?VARIANT_TRUE:VARIANT_FALSE;

	TraceByRefOutArg(pIsActive);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

STDMETHODIMP COPCGroup::put_IsActive(VARIANT_BOOL IsActive)
{
	TRACE_FUNCTION_ENTER("COPCGroup::put_IsActive");

	// check if group has been removed.
	if (m_pParent == NULL) return E_FAIL;

	// check arguments.
	if (!CheckByValInArg(IsActive)) return E_INVALIDARG; 

	// check if group is connected.
	if (!m_pOPCGroup)
	{
		TRACE_INTERFACE_ERROR(IOPCGroupStateMgt);
		return E_FAIL;
	}

	BOOL active = (IsActive == VARIANT_TRUE)?TRUE:FALSE;

	// update server state.
	DWORD dwRevisedRate = 0;

	TRACE_METHOD_CALL(SetState);

	HRESULT hResult = m_pOPCGroup->SetState( 
		NULL, 
		&dwRevisedRate, 
		&active, 
		NULL,
		NULL,
		NULL,
		NULL 
	);

	if (FAILED(hResult))
	{
		TRACE_METHOD_ERROR(SetState, hResult);
		return hResult;
	}

	// update cached value.
	m_active = active;

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// IsSubscribed Property

STDMETHODIMP COPCGroup::get_IsSubscribed(VARIANT_BOOL* pIsSubscribed)
{
	TRACE_FUNCTION_ENTER("COPCGroup::get_IsSubscribed");

	// check if group has been removed.
	if (m_pParent == NULL) return E_FAIL;

	// check arguments.
	if (!CheckByRefOutArg(pIsSubscribed)) return E_INVALIDARG; 

	*pIsSubscribed = m_subscribed;

	TraceByRefOutArg(pIsSubscribed);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

STDMETHODIMP COPCGroup::put_IsSubscribed(VARIANT_BOOL IsSubscribed)
{
	TRACE_FUNCTION_ENTER("COPCGroup::get_IsSubscribed");

	// check if group has been removed.
	if (m_pParent == NULL) return E_FAIL;

	// check arguments.
	if (!CheckByValInArg(IsSubscribed)) return E_INVALIDARG; 

	// check if group is connected.
	if (!m_pOPCGroup)
	{
		TRACE_INTERFACE_ERROR(IOPCGroupStateMgt);
		return E_FAIL;
	}

	// do nothing if state does not change.
	if (IsSubscribed == m_subscribed)
	{
		TRACE_FUNCTION_LEAVE();
		return S_OK;
	}

	HRESULT hResult = S_OK;

	// establish a subscription.
	if (IsSubscribed)
	{
		if(m_usingCP)
		{
			IConnectionPointContainer* ipCPC = NULL;
		
			TRACE_METHOD_CALL(QueryInterface);

			hResult = m_pOPCGroup->QueryInterface(IID_IConnectionPointContainer, (LPVOID*)&ipCPC);
		
			if (FAILED(hResult))
			{
				TRACE_INTERFACE_ERROR(IConnectionPointContainer);
			}
			else
			{
				IConnectionPoint* ipCP = NULL;
				
				TRACE_METHOD_CALL(FindConnectionPoint);
				hResult = ipCPC->FindConnectionPoint(IID_IOPCDataCallback, &ipCP);
				ipCPC->Release();

				if (FAILED(hResult))
				{
					TRACE_METHOD_ERROR(FindConnectionPoint, hResult);
					return hResult;
				}

				IUnknown* ipCallback = NULL;
				((COPCGroup*)this)->QueryInterface(IID_IUnknown, (LPVOID*)&ipCallback);

				TRACE_METHOD_CALL(Advise);
				hResult = ipCP->Advise(ipCallback, &m_DataConnection);
				
				ipCP->Release();
				ipCallback->Release();

				if (FAILED(hResult))
				{
					TRACE_METHOD_ERROR(Advise, hResult);
					return hResult;
				}
			}
		}
		else
		{
			m_pIDataObject = m_pOPCGroup;

			if (!m_pIDataObject)
			{
				TRACE_INTERFACE_ERROR(IDataObject);
				return hResult;
			}

			IAdviseSink* ipAdviseSink = NULL;

			hResult = ((COPCGroup*)this)->QueryInterface(IID_IAdviseSink, (LPVOID*)&ipAdviseSink);

			if (FAILED(hResult))
			{
				TRACE_INTERFACE_ERROR(IAdviseSink);
				return hResult;
			}

			// initialize data format structure.
			FORMATETC formatEtc;

			formatEtc.cfFormat = OPCSTMFORMATDATATIME;
			formatEtc.tymed    = TYMED_HGLOBAL;
			formatEtc.ptd      = NULL;
			formatEtc.dwAspect = DVASPECT_CONTENT;
			formatEtc.lindex   = -1;

			// establish data read advise sink.
			TRACE_METHOD_CALL(DAdvise);
			
			hResult = m_pIDataObject->DAdvise(
				&formatEtc,
				(DWORD)ADVF_PRIMEFIRST,
				ipAdviseSink,
				&m_DataConnection
			);

			if (FAILED(hResult))
			{
				TRACE_METHOD_ERROR(DAdvise, hResult);
			}
			else 
			{
				// establish data write advise sink.
				formatEtc.cfFormat = OPCSTMFORMATWRITECOMPLETE;

				TRACE_METHOD_CALL(DAdvise);

				hResult = m_pIDataObject->DAdvise(
					&formatEtc,
					(DWORD)ADVF_PRIMEFIRST,
					ipAdviseSink,
					&m_WriteConnection
				);
			}

			// DAdvise will hold its own reference
			ipAdviseSink->Release();
		}

		m_subscribed = VARIANT_TRUE;
	}

	// close a subscription.
	else
	{
		if (m_usingCP)
		{
			// OPC 2.0 ConnectionPoints
			IConnectionPointContainer* ipCPC = NULL;

			TRACE_METHOD_CALL(QueryInterface);

			hResult = m_pOPCGroup->QueryInterface(IID_IConnectionPointContainer, (LPVOID*)&ipCPC);
			
			if (FAILED(hResult))
			{
				TRACE_INTERFACE_ERROR(IConnectionPointContainer);
			}
			else
			{
				IConnectionPoint* ipCP = NULL;

				TRACE_METHOD_CALL(FindConnectionPoint);

				hResult = ipCPC->FindConnectionPoint(IID_IOPCDataCallback, &ipCP);
				
				if (SUCCEEDED(hResult))
				{
					TRACE_METHOD_CALL(Unadvise);
					ipCP->Unadvise(m_DataConnection);
					ipCP->Release();
				}

				ipCPC->Release();
			}
		}
		else
		{
			if(m_pIDataObject.p != NULL)   // if a valid interface pointer
			{
				TRACE_METHOD_CALL(DUnadvise);
				m_pIDataObject->DUnadvise(m_DataConnection);

				TRACE_METHOD_CALL(DUnadvise);
				m_pIDataObject->DUnadvise(m_WriteConnection);

				m_pIDataObject.Release();
			}

			// clear unused callback ids.
			m_DataConnection  = NULL;
			m_WriteConnection = NULL;
		}

		m_subscribed = VARIANT_FALSE;
	}

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// ClientHandle Property

STDMETHODIMP COPCGroup::get_ClientHandle(LONG* pClientHandle)
{
	TRACE_FUNCTION_ENTER("COPCGroup::get_ClientHandle");
	
	// check if group has been removed.
	if (m_pParent == NULL) return E_FAIL;

	// check arguments.
	if (!CheckByRefOutArg(pClientHandle)) return E_INVALIDARG; 

	*pClientHandle = m_client;

	TraceByRefOutArg(pClientHandle);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

STDMETHODIMP COPCGroup::put_ClientHandle(LONG ClientHandle)
{
	TRACE_FUNCTION_ENTER("COPCGroup::put_ClientHandle");

	// check if group has been removed.
	if (m_pParent == NULL) return E_FAIL;

	// check arguments.
	if (!CheckByValInArg(ClientHandle)) return E_INVALIDARG; 

	m_client = ClientHandle;
	
	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// ServerHandle Property

STDMETHODIMP COPCGroup::get_ServerHandle(LONG* pServerHandle)
{
	TRACE_FUNCTION_ENTER("COPCGroup::get_ServerHandle");

	// check if group has been removed.
	if (m_pParent == NULL) return E_FAIL;

	// check arguments.
	if (!CheckByRefOutArg(pServerHandle)) return E_INVALIDARG; 

	*pServerHandle = (LONG)this;
	
	TraceByRefOutArg(pServerHandle);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// LocaleID Property

STDMETHODIMP COPCGroup::get_LocaleID(LONG* pLocaleID)
{
	TRACE_FUNCTION_ENTER("COPCGroup::get_LocaleID");

	// check if group has been removed.
	if (m_pParent == NULL) return E_FAIL;

	// check arguments.
	if (!CheckByRefOutArg(pLocaleID)) return E_INVALIDARG; 

	*pLocaleID = 0;

	// check if group is connected.
	if (!m_pOPCGroup)
	{
		TRACE_INTERFACE_ERROR(IOPCGroupStateMgt);
		return E_FAIL;
	}

	// fetch locale from the server.
	LPWSTR    wszName = NULL;
	OPCHANDLE hClient = NULL;

	TRACE_METHOD_CALL(GetState);

	HRESULT hResult = m_pOPCGroup->GetState( 
		&m_actualRate, 
		&m_active, 
		&wszName, 
		&m_timebias,
		&m_deadband, 
		&m_LCID, 
		&hClient, 
		&m_server 
	);

	if (FAILED(hResult))
	{
		TRACE_METHOD_ERROR(GetState, hResult);
		return hResult;
	}

	CoTaskMemFree(wszName);

	*pLocaleID = m_LCID;

	TraceByRefOutArg(pLocaleID);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

STDMETHODIMP COPCGroup::put_LocaleID(LONG LocaleID)
{
	TRACE_FUNCTION_ENTER("COPCGroup::put_LocaleID");

	// check if group has been removed.
	if (m_pParent == NULL) return E_FAIL;

	// check arguments.
	if (!CheckByValInArg(LocaleID)) return E_INVALIDARG; 

	// check if group is connected.
	if (!m_pOPCGroup)
	{
		TRACE_INTERFACE_ERROR(IOPCGroupStateMgt);
		return E_FAIL;
	}

	// update server state.
	DWORD dwRevisedRate = 0;

	TRACE_METHOD_CALL(SetState);

	HRESULT hResult = m_pOPCGroup->SetState( 
		NULL, 
		&dwRevisedRate, 
		NULL, 
		NULL,
		NULL,
		(DWORD*)&LocaleID,
		NULL 
	);

	if (FAILED(hResult))
	{
		TRACE_METHOD_ERROR(SetState, hResult);
		return hResult;
	}

	m_LCID = LocaleID;

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// TimeBias Property

STDMETHODIMP COPCGroup::get_TimeBias(LONG* pTimeBias)
{
	TRACE_FUNCTION_ENTER("COPCGroup::get_TimeBias");
	
	// check if group has been removed.
	if (m_pParent == NULL) return E_FAIL;

	// check arguments.
	if (!CheckByRefOutArg(pTimeBias)) return E_INVALIDARG; 

	*pTimeBias = 0;

	// check if group is connected.
	if (!m_pOPCGroup)
	{
		TRACE_INTERFACE_ERROR(IOPCGroupStateMgt);
		return E_FAIL;
	}

	// fetch timebias from the server.
	LPWSTR    wszName = NULL;
	OPCHANDLE hClient = NULL;

	TRACE_METHOD_CALL(GetState);

	HRESULT hResult = m_pOPCGroup->GetState( 
		&m_actualRate, 
		&m_active, 
		&wszName, 
		&m_timebias,
		&m_deadband, 
		&m_LCID, 
		&hClient, 
		&m_server 
	);

	if (FAILED(hResult))
	{
		TRACE_METHOD_ERROR(GetState, hResult);
		return hResult;
	}

	CoTaskMemFree(wszName);

	*pTimeBias = m_timebias;

	TraceByRefOutArg(pTimeBias);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

STDMETHODIMP COPCGroup::put_TimeBias(LONG TimeBias)
{
	TRACE_FUNCTION_ENTER("COPCGroup::put_TimeBias");

	// check if group has been removed.
	if (m_pParent == NULL) return E_FAIL;

	// check arguments.
	if (!CheckByValInArg(TimeBias)) return E_INVALIDARG; 

	// check if group is connected.
	if (!m_pOPCGroup)
	{
		TRACE_INTERFACE_ERROR(IOPCGroupStateMgt);
		return E_FAIL;
	}

	// update server state.
	DWORD dwRevisedRate = 0;

	TRACE_METHOD_CALL(SetState);

	HRESULT hResult = m_pOPCGroup->SetState( 
		NULL, 
		&dwRevisedRate, 
		NULL, 
		&TimeBias,
		NULL,
		NULL,
		NULL 
	);

	if (FAILED(hResult))
	{
		TRACE_METHOD_ERROR(SetState, hResult);
		return hResult;
	}

	m_timebias = TimeBias;

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// DeadBand Property

STDMETHODIMP COPCGroup::get_DeadBand(FLOAT* pDeadBand)
{
	TRACE_FUNCTION_ENTER("COPCGroup::get_DeadBand");
	
	// check if group has been removed.
	if (m_pParent == NULL) return E_FAIL;

	// check arguments.
	if (!CheckByRefOutArg(pDeadBand)) return E_INVALIDARG; 

	*pDeadBand = 0;

	// check if group is connected.
	if (!m_pOPCGroup)
	{
		TRACE_INTERFACE_ERROR(IOPCGroupStateMgt);
		return E_FAIL;
	}

	// fetch current state from the server.
	LPWSTR    wszName = NULL;
	OPCHANDLE hClient = NULL;

	TRACE_METHOD_CALL(GetState);

	HRESULT hResult = m_pOPCGroup->GetState( 
		&m_actualRate, 
		&m_active, 
		&wszName, 
		&m_timebias,
		&m_deadband, 
		&m_LCID, 
		&hClient, 
		&m_server 
	);

	if (FAILED(hResult))
	{
		TRACE_METHOD_ERROR(GetState, hResult);
		return hResult;
	}

	CoTaskMemFree(wszName);

    *pDeadBand = m_deadband;

	TraceByRefOutArg(pDeadBand);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

STDMETHODIMP COPCGroup::put_DeadBand(FLOAT DeadBand)
{
	TRACE_FUNCTION_ENTER("COPCGroup::put_DeadBand");

	// check if group has been removed.
	if (m_pParent == NULL) return E_FAIL;

	// check arguments.
	if (!CheckByValInArg(DeadBand)) return E_INVALIDARG; 

	// check if group is connected.
	if (!m_pOPCGroup)
	{
		TRACE_INTERFACE_ERROR(IOPCGroupStateMgt);
		return E_FAIL;
	}

	// update server state.
	DWORD dwRevisedRate = 0;

	TRACE_METHOD_CALL(SetState);

	HRESULT hResult = m_pOPCGroup->SetState( 
		NULL, 
		&dwRevisedRate, 
		NULL, 
		NULL,
		&DeadBand,
		NULL,
		NULL 
	);

	if (FAILED(hResult))
	{
		TRACE_METHOD_ERROR(SetState, hResult);
		return hResult;
	}

	m_deadband = DeadBand;

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// UpdateRate Property

STDMETHODIMP COPCGroup::get_UpdateRate(LONG* pUpdateRate)
{
	TRACE_FUNCTION_ENTER("COPCGroup::get_UpdateRate");

	// check if group has been removed.
	if (m_pParent == NULL) return E_FAIL;

	// check arguments.
	if (!CheckByRefOutArg(pUpdateRate)) return E_INVALIDARG; 

	*pUpdateRate = 0;

	// check if group is connected.
	if (!m_pOPCGroup)
	{
		TRACE_INTERFACE_ERROR(IOPCGroupStateMgt);
		return E_FAIL;
	}

	// fetch current state from the server.
	LPWSTR    wszName = NULL;
	OPCHANDLE hClient = NULL;

	TRACE_METHOD_CALL(GetState);

	HRESULT hResult = m_pOPCGroup->GetState( 
		&m_actualRate, 
		&m_active, 
		&wszName, 
		&m_timebias,
		&m_deadband, 
		&m_LCID, 
		&hClient, 
		&m_server 
	);

	if (FAILED(hResult))
	{
		TRACE_METHOD_ERROR(GetState, hResult);
		return hResult;
	}

	CoTaskMemFree(wszName);

    *pUpdateRate = m_actualRate;

	TraceByRefOutArg(pUpdateRate);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

STDMETHODIMP COPCGroup::put_UpdateRate(LONG UpdateRate)
{
	TRACE_FUNCTION_ENTER("COPCGroup::put_UpdateRate");

	// check if group has been removed.
	if (m_pParent == NULL) return E_FAIL;

	// check arguments.
	if (!CheckByValInArg(UpdateRate)) return E_INVALIDARG; 

	// check if group is connected.
	if (!m_pOPCGroup)
	{
		TRACE_INTERFACE_ERROR(IOPCGroupStateMgt);
		return E_FAIL;
	}

	// update server state.
	DWORD dwRevisedRate = 0;

	TRACE_METHOD_CALL(SetState);

	HRESULT hResult = m_pOPCGroup->SetState( 
		(DWORD*)&UpdateRate, 
		&dwRevisedRate, 
		NULL, 
		NULL,
		NULL,
		NULL,
		NULL 
	);

	if (FAILED(hResult))
	{
		TRACE_METHOD_ERROR(SetState, hResult);
		return hResult;
	}

	m_rate       = UpdateRate;
	m_actualRate = dwRevisedRate;

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// OPCItems Property

STDMETHODIMP COPCGroup::get_OPCItems(OPCItems** ppItems)
{
	TRACE_FUNCTION_ENTER("COPCGroup::get_OPCItems");

	// check if group has been removed.
	if (m_pParent == NULL) return E_FAIL;

	// check arguments.
	if (!CheckByRefOutArg(ppItems)) return E_INVALIDARG; 

	// get interface pointer
	m_pItems->QueryInterface(IID_OPCItems, (void**)ppItems);
	_ASSERTE(*ppItems);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// SyncRead Method

STDMETHODIMP COPCGroup::SyncRead( 
	SHORT       Source,
	LONG        NumItems,
	SAFEARRAY** ppServerHandles,
	SAFEARRAY** ppValues,
	SAFEARRAY** ppErrors,
	VARIANT*    pQualities,  
	VARIANT*    pTimeStamps
)
{
	TRACE_FUNCTION_ENTER("COPCGroup::SyncRead");

	// check if group has been removed.
	if (m_pParent == NULL) return E_FAIL;

	// check arguments.
	if (!CheckByValInArg(Source))                                return E_INVALIDARG;
	if (!CheckByValInArg(NumItems))                              return E_INVALIDARG;
	if (!CheckByRefInArrayArg(ppServerHandles, VT_I4, NumItems)) return E_INVALIDARG;
	if (!CheckByRefOutArg(ppValues))                             return E_INVALIDARG;
	if (!CheckByRefOutArg(ppErrors))                             return E_INVALIDARG;
	if (!CheckByRefOutArg(pQualities))                           return E_INVALIDARG;
	if (!CheckByRefOutArg(pTimeStamps))                          return E_INVALIDARG;
	
	// get a handle that is guaranteed to be invalid.
	OPCHANDLE hInvalidHandle = GetInvalidHandle();

	// construct arrays of items and server handles.
	OPCHANDLE* phHandles = (OPCHANDLE*)malloc(NumItems*sizeof(OPCHANDLE));
	memset(phHandles, 0, NumItems*sizeof(OPCHANDLE));

	COPCItem** ppItems = (COPCItem**)malloc(NumItems*sizeof(COPCItem*));
	memset(ppItems, NULL, NumItems*sizeof(COPCItem*));

	for (LONG ii = 1; ii <= NumItems; ii++)
	{
		LONG hHandle = NULL;
		SafeArrayGetElement(*ppServerHandles, &ii, &hHandle);

		phHandles[ii-1] = hInvalidHandle;
		ppItems[ii-1]   = NULL;

		COPCItem* pItem = LookupItem(hHandle);
		
		if (pItem != NULL)
		{
			phHandles[ii-1] = pItem->GetServerHandle();
			ppItems[ii-1]   = pItem;
		}
	}

	// read values from the servers.
	OPCITEMSTATE* pItemStates = NULL;
	HRESULT*      pErrors     = NULL;

	TRACE_METHOD_CALL(SyncIO::Read)

	HRESULT hResult = m_pSyncIO->Read( 
		(OPCDATASOURCE)Source,
		(DWORD)NumItems,
		phHandles,
		&pItemStates,
        &pErrors
	);

	free(phHandles);
	
	if (FAILED(hResult))
	{
		free(ppItems);
		TRACE_METHOD_ERROR(SyncIO::Read, hResult)
		return hResult;
	}

	// check for invalid response from server.
	if (pItemStates == NULL || pErrors == NULL)
	{
		free(ppItems);
		TRACE_METHOD_ERROR(SyncIO::Read, E_POINTER)
		return E_POINTER;
	}

	// copy values and errors in returned arrays.
	SAFEARRAYBOUND cBounds;
	cBounds.lLbound   = 1;
	cBounds.cElements = NumItems;

	*ppValues = SafeArrayCreate(VT_VARIANT, 1, &cBounds);
	*ppErrors = SafeArrayCreate(VT_I4, 1, &cBounds);

	for (LONG ii = 1; ii <= NumItems; ii++)
	{
		// override the error code for invalid items.
		if (ppItems[ii-1] == NULL)
		{
			pErrors[ii-1] = OPC_E_INVALIDHANDLE;
		}

		SafeArrayPutElement(*ppErrors, &ii, &pErrors[ii-1]);

		// only return value if successful.
		if (SUCCEEDED(pErrors[ii-1]))
		{
			VariantToAutomation(&pItemStates[ii-1].vDataValue);
			SafeArrayPutElement(*ppValues, &ii, &pItemStates[ii-1].vDataValue);

			ppItems[ii-1]->Update(&pItemStates[ii-1]);
		}
		
		// free the returned memory.
		VariantClear(&pItemStates[ii-1].vDataValue);
	}

	// free item array.
	free(ppItems);

	// copy qualities in returned arrays.
	if (pQualities != NULL && pQualities->scode != DISP_E_PARAMNOTFOUND)
	{
		VariantClear(pQualities);

		SAFEARRAY* pArray = SafeArrayCreate(VT_I2, 1, &cBounds);
	
		for (LONG ii = 1; ii <= NumItems; ii++)
		{
			WORD wQuality = pItemStates[ii-1].wQuality;

			if (FAILED(pErrors[ii-1]))
			{
				wQuality = OPC_QUALITY_BAD;
			}

			SafeArrayPutElement(pArray, &ii, &wQuality);
		}

		pQualities->vt     = VT_ARRAY | VT_I2;
		pQualities->parray = pArray;
	}

	// copy timestamps in returned arrays.
	if (pTimeStamps != NULL && pTimeStamps->scode != DISP_E_PARAMNOTFOUND)
	{
		VariantClear(pTimeStamps);

		SAFEARRAY* pArray = SafeArrayCreate(VT_DATE, 1, &cBounds);
	
		for (LONG ii = 1; ii <= NumItems; ii++)
		{
			DATE dateTimeStamp = 0;

			if (SUCCEEDED(pErrors[ii-1]))
			{
				dateTimeStamp = GetUtcTime(pItemStates[ii-1].ftTimeStamp);
			}

			SafeArrayPutElement(pArray, &ii, &dateTimeStamp);
		}

		pTimeStamps->vt     = VT_ARRAY | VT_DATE;
		pTimeStamps->parray = pArray;
	}

	CoTaskMemFree(pItemStates);
    CoTaskMemFree(pErrors);

	TraceByRefOutArg(ppValues);
	TraceByRefOutArg(ppErrors);
	TraceByRefOutArg(pQualities);
	TraceByRefOutArg(pTimeStamps);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// SyncWrite Method

STDMETHODIMP COPCGroup::SyncWrite(
	LONG        NumItems,
	SAFEARRAY** ppServerHandles,
	SAFEARRAY** ppValues,
	SAFEARRAY** ppErrors
)
{
	TRACE_FUNCTION_ENTER("COPCGroup::SyncWrite");

	// check if group has been removed.
	if (m_pParent == NULL) return E_FAIL;

	// check arguments.
	if (!CheckByValInArg(NumItems))                              return E_INVALIDARG;
	if (!CheckByRefInArrayArg(ppServerHandles, VT_I4, NumItems)) return E_INVALIDARG;
	if (!CheckByRefInArrayArg(ppValues, VT_VARIANT, NumItems))   return E_INVALIDARG;
	if (!CheckByRefOutArg(ppErrors))                             return E_INVALIDARG;
	
	// check if interface is supported,
	if (!m_pSyncIO)
	{
		TRACE_INTERFACE_ERROR(IOPCSyncIO);
		return E_NOINTERFACE;
	}

	// get a handle that is guaranteed to be invalid.
	OPCHANDLE hInvalidHandle = GetInvalidHandle();

	// construct arrays of items, server handles and values.
	OPCHANDLE* phHandles = (OPCHANDLE*)malloc(NumItems*sizeof(OPCHANDLE));
	memset(phHandles, 0, NumItems*sizeof(OPCHANDLE));

	COPCItem** ppItems = (COPCItem**)malloc(NumItems*sizeof(COPCItem*));
	memset(ppItems, NULL, NumItems*sizeof(COPCItem*));

	for (LONG ii = 1; ii <= NumItems; ii++)
	{
		LONG hHandle = NULL;
		SafeArrayGetElement(*ppServerHandles, &ii, &hHandle);

		phHandles[ii-1] = hInvalidHandle;
		ppItems[ii-1]   = NULL;

		COPCItem* pItem = LookupItem(hHandle);
		
		if (pItem != NULL)
		{
			phHandles[ii-1] = pItem->GetServerHandle();
			ppItems[ii-1]   = pItem;
		}
	}

	// acquire a lock on the value array.
	VARIANT* pValues = NULL;

	if (FAILED(SafeArrayAccessData(*ppValues, (void**)&pValues)))
	{
		TRACE_ERROR("'%s, 'ppValues SafeArrayAccessData Failed'", TRACE_FUNCTION);
		return E_UNEXPECTED;
	}

	LONG lLBound = 0;

	if (FAILED(SafeArrayGetLBound(*ppValues, 1, &lLBound)))
	{		
		TRACE_ERROR("'%s, 'ppValues SafeArrayGetLBound Failed'", TRACE_FUNCTION);
		return E_UNEXPECTED;
	}

	// handle arrays passed from VB .NET that have an extra element prepended.
	if (lLBound == 0)
	{
		pValues++;
	}

	// call the server
	HRESULT* pErrors = NULL;

	TRACE_METHOD_CALL(SyncIO::Write);

	HRESULT hResult = m_pSyncIO->Write(
		(DWORD)NumItems,
		phHandles,
		pValues,
		&pErrors
	);
	
	free(phHandles);
	SafeArrayUnaccessData(*ppValues);
	pValues = NULL;

	if (FAILED(hResult))
	{
		free(ppItems);
		TRACE_METHOD_ERROR(SyncIO::Write, hResult);
		return hResult;
	}

	// check for invalid response from server.
	if (pErrors == NULL)
	{		
		free(ppItems);
		TRACE_METHOD_ERROR(SyncIO::Write, E_POINTER);
		return E_POINTER;
	}

	// copy values and errors in returned arrays.
	SAFEARRAYBOUND cBounds;
	cBounds.lLbound   = 1;
	cBounds.cElements = NumItems;

	*ppErrors = SafeArrayCreate(VT_I4, 1, &cBounds);

	for (LONG ii = 1; ii <= NumItems; ii++)
	{		
		// override the error code for invalid items.
		if (ppItems[ii-1] == NULL)
		{
			pErrors[ii-1] = OPC_E_INVALIDHANDLE;
		}

		SafeArrayPutElement(*ppErrors, &ii, &pErrors[ii-1]);
	}

	// free item array.
	free(ppItems);

	CoTaskMemFree(pErrors);

	TraceByRefOutArg(ppErrors);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// AsyncRead Method

STDMETHODIMP COPCGroup::AsyncRead( 
	LONG        NumItems,
	SAFEARRAY** ppServerHandles,
	SAFEARRAY** ppErrors,
	LONG        TransactionID,
	LONG*       pCancelID)
{
	TRACE_FUNCTION_ENTER("COPCGroup::AsyncRead");

	// check if group has been removed.
	if (m_pParent == NULL) return E_FAIL;

	// check arguments.
	if (!CheckByValInArg(NumItems))                              return E_INVALIDARG;
	if (!CheckByRefInArrayArg(ppServerHandles, VT_I4, NumItems)) return E_INVALIDARG;
	if (!CheckByRefOutArg(ppErrors))                             return E_INVALIDARG;
	if (!CheckByValInArg(TransactionID))                         return E_INVALIDARG;
	if (!CheckByRefOutArg(pCancelID))                            return E_INVALIDARG;
	
	// check if interface is supported.
	if (!m_pAsyncIO && !m_pAsyncIO2)
	{
		TRACE_INTERFACE_ERROR(IAsyncIO);
		return E_NOINTERFACE;
	}

	// get a handle that is guaranteed to be invalid.
	OPCHANDLE hInvalidHandle = GetInvalidHandle();

	// construct arrays of items and server handles.
	OPCHANDLE* phHandles = (OPCHANDLE*)malloc(NumItems*sizeof(OPCHANDLE));
	memset(phHandles, 0, NumItems*sizeof(OPCHANDLE));

	COPCItem** ppItems = (COPCItem**)malloc(NumItems*sizeof(COPCItem*));
	memset(ppItems, NULL, NumItems*sizeof(COPCItem*));

	for (LONG ii = 1; ii <= NumItems; ii++)
	{
		LONG hHandle = NULL;
		SafeArrayGetElement(*ppServerHandles, &ii, &hHandle);

		phHandles[ii-1] = hInvalidHandle;
		ppItems[ii-1]   = NULL;

		COPCItem* pItem = LookupItem(hHandle);
		
		if (pItem != NULL)
		{
			phHandles[ii-1] = pItem->GetServerHandle();
			ppItems[ii-1]   = pItem;
		}
	}

	// call the server
	HRESULT  hResult = S_OK;
	HRESULT* pErrors = NULL;

	TRACE_METHOD_CALL(AsyncIO::Read)

	if (m_usingCP)
	{
		hResult = m_pAsyncIO2->Read(
			(DWORD)NumItems,
			phHandles,
			TransactionID,
			(DWORD*)pCancelID,
			&pErrors
		);
	}
	else
	{
		// set this flag because callbacks can occur while waiting for a reply.
		m_asyncReading = TRUE;
		
		hResult = m_pAsyncIO->Read(
			m_DataConnection,
			OPC_DS_DEVICE,
			(DWORD)NumItems,
			phHandles,
			(DWORD*)pCancelID,
			&pErrors
		);
		
		m_asyncReading = FALSE;

		// need to save the cancel ID to figure out if the callback is a read or refresh
		if (SUCCEEDED(hResult))
		{
			BOOL bFound = FALSE;

			list<OPCHANDLE>::iterator iter = m_readIDs.begin();
			
			while (iter != m_readIDs.end() && !bFound)
			{
				if (*iter == *pCancelID)
				{
					m_readIDs.remove(*pCancelID);
					bFound = TRUE;
				}

				iter++;
			}

			if (!bFound)
			{
				m_readIDs.push_back(*pCancelID);
			}
		}
	}

	free(phHandles);
	
	if (FAILED(hResult))
	{
		free(ppItems);
		TRACE_METHOD_ERROR(AsyncIO::Read, hResult)
		return hResult;
	}

	// check for invalid response from server.
	if (pErrors == NULL)
	{
		free(ppItems);
		TRACE_METHOD_ERROR(AsyncIO::Read, E_POINTER)
		return E_POINTER;
	}

	// copy errors into returned arrays.
	SAFEARRAYBOUND cBounds;
	cBounds.lLbound   = 1;
	cBounds.cElements = NumItems;

	*ppErrors = SafeArrayCreate(VT_I4, 1, &cBounds);

	for (LONG ii = 1; ii <= NumItems; ii++)
	{		
		// override the error code for invalid items.
		if (ppItems[ii-1] == NULL)
		{
			pErrors[ii-1] = OPC_E_INVALIDHANDLE;
		}

		SafeArrayPutElement(*ppErrors, &ii, &pErrors[ii-1]);
	}

	// free item array.
	free(ppItems);

	CoTaskMemFree(pErrors);

	TraceByRefOutArg(ppErrors);
	TraceByRefOutArg(pCancelID);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// AsyncWrite Method

STDMETHODIMP COPCGroup::AsyncWrite(
	LONG		NumItems,
	SAFEARRAY** ppServerHandles,
	SAFEARRAY** ppValues,
	SAFEARRAY** ppErrors,
	LONG        TransactionID,
	LONG*       pCancelID
)
{
	TRACE_FUNCTION_ENTER("COPCGroup::AsyncWrite");

	// check if group has been removed.
	if (m_pParent == NULL) return E_FAIL;

	// check arguments.
	if (!CheckByValInArg(NumItems))                              return E_INVALIDARG;
	if (!CheckByRefInArrayArg(ppServerHandles, VT_I4, NumItems)) return E_INVALIDARG;
	if (!CheckByRefInArrayArg(ppValues, VT_VARIANT, NumItems))   return E_INVALIDARG;
	if (!CheckByRefOutArg(ppErrors))                             return E_INVALIDARG;
	if (!CheckByValInArg(TransactionID))                         return E_INVALIDARG;
	if (!CheckByRefOutArg(pCancelID))                            return E_INVALIDARG;
	
	// check if interface is supported.
	if (!m_pAsyncIO && !m_pAsyncIO2)
	{		
		TRACE_INTERFACE_ERROR(IOPCAsyncIO);
		return E_NOINTERFACE;
	}

	// get a handle that is guaranteed to be invalid.
	OPCHANDLE hInvalidHandle = GetInvalidHandle();

	// construct arrays of items, server handles and values.
	OPCHANDLE* phHandles = (OPCHANDLE*)malloc(NumItems*sizeof(OPCHANDLE));
	memset(phHandles, 0, NumItems*sizeof(OPCHANDLE));

	COPCItem** ppItems = (COPCItem**)malloc(NumItems*sizeof(COPCItem*));
	memset(ppItems, NULL, NumItems*sizeof(COPCItem*));

	for (LONG ii = 1; ii <= NumItems; ii++)
	{
		LONG hHandle = NULL;
		SafeArrayGetElement(*ppServerHandles, &ii, &hHandle);

		phHandles[ii-1] = hInvalidHandle;
		ppItems[ii-1]   = NULL;

		COPCItem* pItem = LookupItem(hHandle);
		
		if (pItem != NULL)
		{
			phHandles[ii-1] = pItem->GetServerHandle();
			ppItems[ii-1]   = pItem;
		}
	}

	// acquire a lock on the value array.
	VARIANT* pValues = NULL;

	if (FAILED(SafeArrayAccessData(*ppValues, (void**)&pValues)))
	{
		TRACE_ERROR("'%s, 'ppValues SafeArrayAccessData Failed'", TRACE_FUNCTION);
		return E_UNEXPECTED;
	}

	LONG lLBound = 0;

	if (FAILED(SafeArrayGetLBound(*ppValues, 1, &lLBound)))
	{		
		TRACE_ERROR("'%s, 'ppValues SafeArrayGetLBound Failed'", TRACE_FUNCTION);
		return E_UNEXPECTED;
	}

	// handle arrays passed from VB .NET that have an extra element prepended.
	if (lLBound == 0)
	{
		pValues++;
	}

	// call the server
	HRESULT  hResult = S_OK;
	HRESULT* pErrors = NULL;

	TRACE_METHOD_CALL(AsyncIO::Write)

	if (m_usingCP)
	{
		hResult = m_pAsyncIO2->Write(
			(DWORD)NumItems,
			phHandles,
			pValues,
			TransactionID,
			(DWORD*)pCancelID,
			&pErrors
		);
   }
	else
	{
		hResult = m_pAsyncIO->Write(
			m_WriteConnection,
			(DWORD)NumItems,
			phHandles,
			pValues,
			(DWORD*)pCancelID,
			&pErrors
		);
	}
	
	free(phHandles);
	SafeArrayUnaccessData(*ppValues);
	pValues = NULL;

	if (FAILED(hResult))
	{
		free(ppItems);
		TRACE_METHOD_ERROR(AsyncIO::Write, hResult);
		return hResult;
	}

	// check for invalid response from server.
	if (pErrors == NULL)
	{
		free(ppItems);
		TRACE_METHOD_ERROR(AsyncIO::Write, E_POINTER);
		return E_POINTER;
	}

	// copy values and errors in returned arrays.
	SAFEARRAYBOUND cBounds;
	cBounds.lLbound   = 1;
	cBounds.cElements = NumItems;

	*ppErrors = SafeArrayCreate(VT_I4, 1, &cBounds);

	for (LONG ii = 1; ii <= NumItems; ii++)
	{
		// override the error code for invalid items.
		if (ppItems[ii-1] == NULL)
		{
			pErrors[ii-1] = OPC_E_INVALIDHANDLE;
		}

		SafeArrayPutElement(*ppErrors, &ii, &pErrors[ii-1]);
	}

	// free item array.
	free(ppItems);

	CoTaskMemFree(pErrors);

	TraceByRefOutArg(ppErrors);
	TraceByRefOutArg(pCancelID);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// AsyncRefresh Method

STDMETHODIMP COPCGroup::AsyncRefresh(
	SHORT Source,
	LONG  TransactionID,
	LONG* pCancelID
)
{
	TRACE_FUNCTION_ENTER("COPCGroup::AsyncRefresh");

	// check if group has been removed.
	if (m_pParent == NULL) return E_FAIL;

	// check arguments.
	if (!CheckByValInArg(Source))        return E_INVALIDARG; 
	if (!CheckByValInArg(TransactionID)) return E_INVALIDARG; 
	if (!CheckByRefOutArg(pCancelID))    return E_INVALIDARG; 

	// call the server
	HRESULT  hResult = S_OK;

	TRACE_METHOD_CALL(AsyncIO::Refresh);

	if (m_usingCP)
	{
		hResult = m_pAsyncIO2->Refresh2((OPCDATASOURCE)Source, TransactionID, (DWORD*)pCancelID);
    }
	else
	{
		hResult = m_pAsyncIO->Refresh(m_DataConnection, (OPCDATASOURCE)Source, (DWORD*)pCancelID);
	}
	
	if (FAILED(hResult))
	{
		TRACE_METHOD_ERROR(AsyncIO::Refresh, hResult);
		return hResult;
	}

	TraceByRefOutArg(pCancelID);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// AsyncCancel Method

STDMETHODIMP COPCGroup::AsyncCancel(LONG CancelID)
{
	TRACE_FUNCTION_ENTER("COPCGroup::AsyncCancel");

	// check if group has been removed.
	if (m_pParent == NULL) return E_FAIL;

	// check arguments.
	if (!CheckByValInArg(CancelID)) return E_INVALIDARG; 

	// call the server
	HRESULT  hResult = S_OK;

	TRACE_METHOD_CALL(AsyncIO::Cancel);

	if (m_usingCP)
	{
		hResult = m_pAsyncIO2->Cancel2((DWORD)CancelID);
    }
	else
	{
		hResult = m_pAsyncIO->Cancel((DWORD)CancelID);
	}
	
	if (FAILED(hResult))
	{
		TRACE_METHOD_ERROR(AsyncIO::Cancel, hResult);
		return hResult;
	}

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}
