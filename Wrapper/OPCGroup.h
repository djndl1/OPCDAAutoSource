//==============================================================================
// TITLE: OPCGroup.h
//
// CONTENTS:
//
// A class that implements the DA Automation OPCGroup object.
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
// 2002/11/17 RSA   Merged bug fixes provided by members.

#ifndef __OPCGroup_H__      
#define __OPCGroup_H__

#include "OPCItem.h"
#include "OPCGroupEvent.h"
#include "OPCItems.h"

class COPCGroups;

//==============================================================================
// CLASS:   COPCGroup
// PURPOSE: Implements all interfaces supported by the OPCGroup object.

class ATL_NO_VTABLE COPCGroup :
   public IAdviseSink,
   public IOPCDataCallback,
   public CComDualImpl<IOPCGroup, &IID_IOPCGroup, &LIBID_OPCAutomation>,
   public IConnectionPointContainerImpl<COPCGroup>,
   public CProxyDIOPCGroupEvent<COPCGroup>,
   public CComObjectRootEx<CComMultiThreadModel>
{
public:

	//==========================================================================
	// ATL Declarations

	BEGIN_COM_MAP(COPCGroup)
		COM_INTERFACE_ENTRY(IAdviseSink)
		COM_INTERFACE_ENTRY(IOPCDataCallback)
		COM_INTERFACE_ENTRY2(IDispatch, IOPCGroup)
		COM_INTERFACE_ENTRY(IOPCGroup)
		COM_INTERFACE_ENTRY_IMPL(IConnectionPointContainer)
	END_COM_MAP()

	BEGIN_CONNECTION_POINT_MAP(COPCGroup)
		CONNECTION_POINT_ENTRY(DIID_DIOPCGroupEvent)
	END_CONNECTION_POINT_MAP()

	DECLARE_NOT_AGGREGATABLE(OPCGroup)
	
	//==========================================================================
	// Public Operators

	// Constructor
	COPCGroup();

	// Destructor
	~COPCGroup();

	//==========================================================================
	// OPCGroup

	STDMETHOD(get_Parent)(IOPCAutoServer** ppParent);
	STDMETHOD(get_Name)(BSTR* pName);
	STDMETHOD(put_Name)(BSTR Name);
	STDMETHOD(get_IsPublic)(VARIANT_BOOL* pIsPublic);
	STDMETHOD(get_IsActive)(VARIANT_BOOL* pIsActive);
	STDMETHOD(put_IsActive)(VARIANT_BOOL IsActive);
	STDMETHOD(get_IsSubscribed)(VARIANT_BOOL* pIsSubscribed);
	STDMETHOD(put_IsSubscribed)(VARIANT_BOOL IsSubscribed);
	STDMETHOD(get_ClientHandle)(LONG* pClientHandle);
	STDMETHOD(put_ClientHandle)(LONG ClientHandle);
	STDMETHOD(get_ServerHandle)(LONG* pServerHandle);
	STDMETHOD(get_LocaleID)(LONG* pLocaleID);
	STDMETHOD(put_LocaleID)(LONG LocaleID);
	STDMETHOD(get_TimeBias)(LONG* pTimeBias);
	STDMETHOD(put_TimeBias)(LONG TimeBias);
	STDMETHOD(get_DeadBand)(FLOAT* pDeadBand);
	STDMETHOD(put_DeadBand)(FLOAT DeadBand);
	STDMETHOD(get_UpdateRate)(LONG* pUpdateRate);
	STDMETHOD(put_UpdateRate)(LONG UpdateRate);
	STDMETHOD(get_OPCItems)(OPCItems** ppItems);
	STDMETHOD(SyncRead)(SHORT Source, LONG NumItems, SAFEARRAY** ppServerHandles, SAFEARRAY** ppValues,	SAFEARRAY** ppErrors, VARIANT* pQualities, VARIANT* pTimeStamps);
	STDMETHOD(SyncWrite)(LONG NumItems, SAFEARRAY** ppServerHandles, SAFEARRAY** ppValues, SAFEARRAY** ppErrors);
	STDMETHOD(AsyncRead)(LONG NumItems, SAFEARRAY** ppServerHandles, SAFEARRAY** ppErrors, LONG TransactionID, LONG* pCancelID);
	STDMETHOD(AsyncWrite)(LONG NumItems, SAFEARRAY** ppServerHandles, SAFEARRAY** ppValues, SAFEARRAY** ppErrors, LONG TransactionID, LONG* pCancelID);
	STDMETHOD(AsyncRefresh)(SHORT Source, LONG TransactionID, LONG* pCancelID);
	STDMETHOD(AsyncCancel)(LONG CancelID);

	//==========================================================================
	// IOPCDataCallback

	STDMETHODIMP OnDataChange(
		DWORD      dwTransid, 
		OPCHANDLE  hGroup, 
		HRESULT    hrMasterquality,
		HRESULT    hrMastererror,
		DWORD      dwCount, 
		OPCHANDLE* phClientItems, 
		VARIANT*   pvValues, 
		WORD*      pwQualities,
		FILETIME*  pftTimeStamps,
		HRESULT*   pErrors
	);

	STDMETHODIMP OnReadComplete(
		DWORD      dwTransid, 
		OPCHANDLE  hGroup, 
		HRESULT    hrMasterquality,
		HRESULT    hrMastererror,
		DWORD      dwCount, 
		OPCHANDLE* phClientItems, 
		VARIANT*   pvValues, 
		WORD*      pwQualities,
		FILETIME*  pftTimeStamps,
		HRESULT*   pErrors
	);

	STDMETHODIMP OnWriteComplete(
		DWORD      dwTransid, 
		OPCHANDLE  hGroup, 
		HRESULT    hrMastererror,
		DWORD      dwCount, 
		OPCHANDLE* phClientItems, 
		HRESULT*   pErrors
	);

	STDMETHODIMP OnCancelComplete(
		DWORD      dwTransid, 
		OPCHANDLE  hGroup
	);

	//==========================================================================
	// IAdviseSink

	STDMETHODIMP_(void) OnDataChange(LPFORMATETC, LPSTGMEDIUM);
	STDMETHODIMP_(void) OnWriteComplete(LPFORMATETC, LPSTGMEDIUM);
	STDMETHODIMP_(void) OnViewChange(DWORD, LONG);
	STDMETHODIMP_(void) OnRename(LPMONIKER);
	STDMETHODIMP_(void) OnSave(void);
	STDMETHODIMP_(void) OnClose(void);

	//==========================================================================
	// Public Methods

	HRESULT Initialize(COPCGroups* pParent, IUnknown* pUnknown);
	void Uninitialize();

	// the name of the group.
	const CComBSTR& GetName() const { return m_name; }

	// the server handle for the group.
	const LONG GetServerHandle() const { return m_server; }

private:

	//==========================================================================
	// Private Methods

	// finds the item object for the specified handle.
	COPCItem* LookupItem(OPCHANDLE ServerHandle) { return m_pItems->LookupItem(ServerHandle); }
	
	// finds a server handle that should always be invalid.
	OPCHANDLE GetInvalidHandle() { return m_pItems->GetInvalidHandle(); }

	//==========================================================================
	// Private Members

	// pointer to server object.
	COPCGroups* m_pParent;

	// state variables.
	CComBSTR  m_name;
	DWORD     m_rate;
	DWORD     m_actualRate;
	BOOL      m_active;
	LONG      m_timebias;
	FLOAT     m_deadband;
	DWORD     m_LCID;
	OPCHANDLE m_client;
	OPCHANDLE m_server;

	// async callback state variables.
	BOOL            m_usingCP;     
	VARIANT_BOOL    m_subscribed;
	DWORD           m_DataConnection;
	DWORD           m_WriteConnection;
	list<OPCHANDLE> m_readIDs;  
	BOOL            m_asyncReading;

	// group interface pointers.
	IOPCGroupStateMgtPtr m_pOPCGroup;
	IOPCSyncIOPtr        m_pSyncIO;
	IOPCAsyncIOPtr       m_pAsyncIO;
	IOPCAsyncIO2Ptr      m_pAsyncIO2;
	IDataObjectPtr       m_pIDataObject;

	// the implementation of the items object.
	COPCItems* m_pItems;
};

typedef CComObject<COPCGroup> COPCGroupObject;

#endif // __OPCGroup_H__