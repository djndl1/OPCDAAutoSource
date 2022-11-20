//==============================================================================
// TITLE: OPCItem.cpp
//
// CONTENTS:
//
//  This is the OPCItem object.
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
//

#include "StdAfx.h"
#include "OPCItem.h"
#include "OPCGroup.h"
#include "OPCUtils.h"
#include "OPCTrace.h"

//==============================================================================
// Constructor

COPCItem::COPCItem()
{
   MagicNumber = MAGIC_NUMBER;

   m_pParent        = NULL;
   m_active         = TRUE;
   m_accessRights   = 0;
   m_datatype       = VT_EMPTY;
   m_nativeDatatype = VT_EMPTY;
   m_client         = 0;
   m_server         = 0;
   m_quality        = 0;
   m_timestamp      = GetMinTime();
   
   m_itemID.Empty();
   m_accessPath.Empty();
   m_value.Clear();
}

//==============================================================================
// Destructor

COPCItem::~COPCItem()
{	
   MagicNumber = 0;

   m_pOPCItem.Release();
   m_pOPCSyncIO.Release();
}

//==============================================================================
// Initialize Method

HRESULT COPCItem::Initialize( 
	OPCHANDLE          hClient,
	OPCITEMDEF*        pItemDef,
	OPCITEMRESULT*     pItemResult,
	IOPCGroupStateMgt* pGroup,
	COPCGroup*         pParent 
)
{
	TRACE_FUNCTION_ENTER("COPCItem::Initialize");

	// save group interface pointer
	m_pOPCItem   = pGroup;
	m_pOPCSyncIO = pGroup;

	// initialize item.
	m_pParent        = pParent;
	m_itemID         = pItemDef->szItemID;
	m_accessPath     = pItemDef->szAccessPath;
	m_active         = pItemDef->bActive;
	m_accessRights   = pItemResult->dwAccessRights;
	m_datatype       = pItemDef->vtRequestedDataType;
	m_nativeDatatype = pItemResult->vtCanonicalDataType;
	m_client         = hClient;
	m_server         = pItemResult->hServer;

	// clear value.
	m_value.Clear();
	m_quality   = OPC_QUALITY_BAD;
	m_timestamp = GetMinTime();

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// Update Method

void COPCItem::Update(const OPCITEMSTATE* pItemState)
{
	// copy value.
	m_value.Copy(&pItemState->vDataValue);

	// convert value to automation compatible data type.
	VariantToAutomation(&m_value);

	// save quality and timestamp.
	m_quality   = pItemState->wQuality;
	m_timestamp = pItemState->ftTimeStamp;
}

//==============================================================================
// Update Method

void COPCItem::Update(
	const VARIANT* pValue,
    const DWORD    quality,
    const FILETIME timestamp
)
{
	// copy value.
	m_value.Copy(pValue);

	// convert value to automation compatible data type.
	VariantToAutomation(&m_value);

	// save quality and timestamp.
	m_quality   = quality;
	m_timestamp = timestamp;
}

//==============================================================================
// Parent Property

STDMETHODIMP COPCItem::get_Parent(OPCGroup** ppParent)
{
	TRACE_FUNCTION_ENTER("COPCItem::get_Parent");

	// check arguments.
	if (!CheckByRefOutArg(ppParent)) return E_INVALIDARG;

	HRESULT hResult = ((COPCGroup*)m_pParent)->QueryInterface(IID_IOPCGroup, (void**)ppParent);

	if (FAILED(hResult))
	{
	   	TRACE_INTERFACE_ERROR(IOPCGroup);
		return hResult;
	}

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// ClientHandle Property

STDMETHODIMP COPCItem::get_ClientHandle(LONG* pClientHandle)
{
	TRACE_FUNCTION_ENTER("COPCItem::get_ClientHandle");
	
	// check arguments.
	if (!CheckByRefOutArg(pClientHandle)) return E_INVALIDARG;

	*pClientHandle = m_client;

	TraceByRefOutArg(pClientHandle);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

STDMETHODIMP COPCItem::put_ClientHandle(LONG ClientHandle)
{
	TRACE_FUNCTION_ENTER("COPCItem::put_ClientHandle");

	// check arguments.
	if (!CheckByValInArg(ClientHandle)) return E_INVALIDARG;

	m_client = ClientHandle;

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// ServerHandle Property

STDMETHODIMP COPCItem::get_ServerHandle(LONG* pServerHandle)
{
	TRACE_FUNCTION_ENTER("COPCItem::get_ServerHandle");
	
	// check arguments.
	if (!CheckByRefOutArg(pServerHandle)) return E_INVALIDARG;

	*pServerHandle = (LONG)this;

	TraceByRefOutArg(pServerHandle);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// AccessPath Property

STDMETHODIMP COPCItem::get_AccessPath(BSTR* pAccessPath)
{
	TRACE_FUNCTION_ENTER("COPCItem::get_AccessPath");
	
	// check arguments.
	if (!CheckByRefOutArg(pAccessPath)) return E_INVALIDARG;

	*pAccessPath = m_accessPath.Copy();

	TraceByRefOutArg(pAccessPath);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// AccessRights Property

STDMETHODIMP COPCItem::get_AccessRights(LONG* pAccessRights)
{
	TRACE_FUNCTION_ENTER("COPCItem::get_AccessRights");
	
	// check arguments.
	if (!CheckByRefOutArg(pAccessRights)) return E_INVALIDARG;

    *pAccessRights = m_accessRights;

	TraceByRefOutArg(pAccessRights);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// ItemID Property

STDMETHODIMP COPCItem::get_ItemID(BSTR* pItemID)
{
	TRACE_FUNCTION_ENTER("COPCItem::get_ItemID");
	
	// check arguments.
	if (!CheckByRefOutArg(pItemID)) return E_INVALIDARG;

	*pItemID = m_itemID.Copy();

	TraceByRefOutArg(pItemID);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// IsActive Property

STDMETHODIMP COPCItem::get_IsActive(VARIANT_BOOL* pIsActive)
{
	TRACE_FUNCTION_ENTER("COPCItem::get_IsActive");
	
	// check arguments.
	if (!CheckByRefOutArg(pIsActive)) return E_INVALIDARG;

	*pIsActive = (m_active)?VARIANT_TRUE:VARIANT_FALSE;

	TraceByRefOutArg(pIsActive);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

STDMETHODIMP COPCItem::put_IsActive(VARIANT_BOOL IsActive)
{
	TRACE_FUNCTION_ENTER("COPCItem::put_IsActive");

	// check arguments.
	if (!CheckByValInArg(IsActive)) return E_INVALIDARG;

	// check if state is not changing.
	if (IsActive && m_active || !IsActive && !m_active)
	{
   		TRACE_FUNCTION_LEAVE();
		return S_OK;
	}

	// update the active state on the server.
	HRESULT* pErrors = NULL;

	TRACE_METHOD_CALL(SetActiveState);

	HRESULT hResult = m_pOPCItem->SetActiveState(1, &m_server, (VARIANT_FALSE == IsActive)?FALSE:TRUE, &pErrors);

	if (FAILED(hResult))
	{
		TRACE_METHOD_ERROR(SetActiveState, hResult);
		return hResult;
	}

	// check the item level error code.
	if (pErrors != NULL)
	{
		if (FAILED(pErrors[0]))
		{
	   		TRACE_ERROR("%s, 'SetActiveState failed - hResult=[0x%08X]'", TRACE_FUNCTION, pErrors[0]);
			hResult = pErrors[0];
		}

		CoTaskMemFree(pErrors);

		if (FAILED(hResult))
		{
			return hResult;
		}
	}

	// update local value.
	m_active = IsActive;

   	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// RequestedDataType Property

STDMETHODIMP COPCItem::get_RequestedDataType(SHORT* pRequestedDataType)
{
	TRACE_FUNCTION_ENTER("COPCItem::get_RequestedDataType");
	
	// check arguments.
	if (!CheckByRefOutArg(pRequestedDataType)) return E_INVALIDARG;

	*pRequestedDataType = m_datatype;

	TraceByRefOutArg(pRequestedDataType);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

STDMETHODIMP COPCItem::put_RequestedDataType(SHORT RequestedDataType)
{
	TRACE_FUNCTION_ENTER("COPCItem::put_RequestedDataType");

	// check arguments.
	if (!CheckByValInArg(RequestedDataType)) return E_INVALIDARG;

	// update the requested data type on the server.
	HRESULT* pErrors = NULL;

	TRACE_METHOD_CALL(SetDatatypes);

	HRESULT hResult = m_pOPCItem->SetDatatypes(1, &m_server, (VARTYPE*)&RequestedDataType, &pErrors);

	if (FAILED(hResult))
	{
		TRACE_METHOD_ERROR(SetDatatypes, hResult);
		return hResult;
	}

	// check the item level error code.
	if (pErrors != NULL)
	{
		if (FAILED(pErrors[0]))
		{
	   		TRACE_ERROR("%s, 'SetDatatypes failed - hResult=[0x%08X]'", TRACE_FUNCTION, pErrors[0]);
			hResult = pErrors[0];
		}

		CoTaskMemFree(pErrors);

		if (FAILED(hResult))
		{
			return hResult;
		}
	}

	// update local value.
	m_datatype = RequestedDataType;

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// Value Property

STDMETHODIMP COPCItem::get_Value(VARIANT* pCurrentValue)
{
	TRACE_FUNCTION_ENTER("COPCItem::get_Value");
	
	// check arguments.
	if (!CheckByRefOutArg(pCurrentValue)) return E_INVALIDARG;

	VariantClear(pCurrentValue);
	VariantCopy(pCurrentValue, &m_value);

	TraceByRefOutArg(pCurrentValue);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// Quality Property

STDMETHODIMP COPCItem::get_Quality(LONG* pQuality)
{
	TRACE_FUNCTION_ENTER("COPCItem::get_Quality");
	
	// check arguments.
	if (!CheckByRefOutArg(pQuality)) return E_INVALIDARG;

    *pQuality = m_quality;

	TraceByRefOutArg(pQuality);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// TimeStamp Property

STDMETHODIMP COPCItem::get_TimeStamp(DATE* pTimeStamp)
{
	TRACE_FUNCTION_ENTER("COPCItem::get_TimeStamp");
	
	// check arguments.
	if (!CheckByRefOutArg(pTimeStamp)) return E_INVALIDARG;

	*pTimeStamp = GetUtcTime(m_timestamp);

	TraceByRefOutArg(pTimeStamp);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// CanonicalDataType Property

STDMETHODIMP COPCItem::get_CanonicalDataType(SHORT* pCanonicalDataType)
{
	TRACE_FUNCTION_ENTER("COPCItem::get_CanonicalDataType");
	
	// check arguments.
	if (!CheckByRefOutArg(pCanonicalDataType)) return E_INVALIDARG;

	*pCanonicalDataType = m_nativeDatatype;

	TraceByRefOutArg(pCanonicalDataType);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// EUType Property

STDMETHODIMP COPCItem::get_EUType(SHORT* pEUType)
{
	TRACE_FUNCTION_ENTER("COPCItem::get_EUType");

	// check arguments.
	if (!CheckByRefOutArg(pEUType)) return E_INVALIDARG;

	TRACE_ERROR("%s, 'E_NOTIMPL'", TRACE_FUNCTION);
	return E_NOTIMPL;
}

//==============================================================================
// EUInfo Property

STDMETHODIMP COPCItem::get_EUInfo(VARIANT* pEUInfo)
{
	TRACE_FUNCTION_ENTER("COPCItem::get_EUInfo");

	// check arguments.
	if (!CheckByRefOutArg(pEUInfo)) return E_INVALIDARG;

	TRACE_ERROR("%s, 'E_NOTIMPL'", TRACE_FUNCTION);
	return E_NOTIMPL;
}

//==============================================================================
// Read Method

STDMETHODIMP COPCItem::Read( 
	SHORT    Source,
	VARIANT* pValue,
	VARIANT* pQuality,
	VARIANT* pTimeStamp
)
{
	TRACE_FUNCTION_ENTER("COPCItem::Read");
	
	// check arguments.
	if (!CheckByValInArg(Source))      return E_INVALIDARG;
	if (!CheckByRefOutArg(pValue))     return E_INVALIDARG;
	if (!CheckByRefOutArg(pQuality))   return E_INVALIDARG;
	if (!CheckByRefOutArg(pTimeStamp)) return E_INVALIDARG;

	// check if interface is supported.
	if (!m_pOPCSyncIO)
	{
	   	TRACE_INTERFACE_ERROR(IOPCSyncIO);
		return E_NOINTERFACE;
	}

	// read value from server.
	OPCITEMSTATE* pItemState = NULL;
	HRESULT*      pErrors    = NULL;

	TRACE_METHOD_CALL(SyncIO::Read);

	HRESULT hResult = m_pOPCSyncIO->Read((OPCDATASOURCE)Source, 1, &m_server, &pItemState, &pErrors);

	if (FAILED(hResult))
	{
		TRACE_METHOD_ERROR(SyncIO::Read, hResult);
		return hResult;
	}

	// check for invalid response from server.
	if (pItemState == NULL || pErrors == NULL)
	{
		TRACE_METHOD_ERROR(SyncIO::Read, E_POINTER);
		return E_POINTER;
	}

	// check the item level error code.
	if (FAILED(pErrors[0]))
	{
		hResult = pErrors[0];

		VariantClear(&pItemState->vDataValue);
		CoTaskMemFree(pItemState);
		CoTaskMemFree(pErrors);

   		TRACE_ERROR("%s, 'SyncIO Read failed - hResult=[0x%08X]'", TRACE_FUNCTION, hResult);
		return hResult;
	}

	// save item state.
    Update(pItemState);

	// return value.
	if (pValue->vt != VT_ERROR && pValue->scode != DISP_E_PARAMNOTFOUND)
	{
		VariantClear(pValue);
		VariantCopy(pValue, &m_value);
	}

	// return quality.
	if (pQuality->vt != VT_ERROR && pQuality->scode != DISP_E_PARAMNOTFOUND)
	{
		VariantClear(pQuality);
        pQuality->vt   = VT_I2;
        pQuality->iVal = (short)m_quality;
	}

   	// return quality.
	if (pTimeStamp->vt != VT_ERROR && pTimeStamp->scode != DISP_E_PARAMNOTFOUND)
	{
		VariantClear(pTimeStamp);
        pTimeStamp->vt   = VT_DATE;
		pTimeStamp->date = GetUtcTime(m_timestamp);
	}

	VariantClear(&pItemState->vDataValue);
	CoTaskMemFree(pItemState);
	CoTaskMemFree(pErrors);
	
	TraceByRefOutArg(pValue);
	TraceByRefOutArg(pQuality);
	TraceByRefOutArg(pTimeStamp);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// Write Method

STDMETHODIMP COPCItem::Write(VARIANT Value)
{
	TRACE_FUNCTION_ENTER("COPCItem::Write");
	
	if (!CheckByValInVariantArg(Value, VT_VARIANT, false)) return E_INVALIDARG;

	// check if interface is supported.
	if (!m_pOPCSyncIO)
	{
	   	TRACE_INTERFACE_ERROR(IOPCSyncIO);
		return E_NOINTERFACE;
	}

	HRESULT hResult = S_OK;

	// convert to requested data type.
	if (m_datatype != VT_EMPTY && m_datatype != m_value.vt && (m_datatype & VT_ARRAY) == 0)
	{
		hResult = VariantChangeType(&m_value, &Value, 0, m_datatype);

		if (FAILED(hResult))
		{
			VariantCopy(&m_value, &Value);
		}
	}
	else
	{
		VariantCopy(&m_value, &Value);
	}

	// write value to server.
	HRESULT* pErrors = NULL;

	TRACE_METHOD_CALL(SyncIO::Write);
  
	hResult = m_pOPCSyncIO->Write(1, &m_server, &m_value, &pErrors);

	if (FAILED(hResult))
	{
		TRACE_METHOD_ERROR(SyncIO::Write, hResult);
		return hResult;
	}

	// check for invalid response from server.
	if (pErrors == NULL)
	{
		TRACE_METHOD_ERROR(SyncIO::Write, E_POINTER);
		return E_POINTER;
	}

	// check the item level error code.
	if (FAILED(pErrors[0]))
	{
		hResult = pErrors[0];
		CoTaskMemFree(pErrors);

   		TRACE_ERROR("%s, 'SyncIO Write failed - hResult=[0x%08X]'", TRACE_FUNCTION, hResult);
		return hResult;
	}

	// free memory.
	CoTaskMemFree(pErrors);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}
