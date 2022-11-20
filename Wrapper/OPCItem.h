//==============================================================================
// TITLE: OPCItem.h
//
// CONTENTS:
//
// A class that implements the DA Automation OPCItem object.
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

#ifndef __OPCItem_H__   
#define __OPCItem_H__

class COPCGroup;

//==============================================================================
// MACRO:   MAGIC_NUMBER
// PURPOSE: A number used to test whether an COPCItem pointer is a valid address.

#define MAGIC_NUMBER 0x0A0B0C0D

//==============================================================================
// CLASS:   COPCItem
// PURPOSE: Implements all interfaces supported by the OPCItem object.

class ATL_NO_VTABLE COPCItem :
   public CComDualImpl<OPCItem, &IID_OPCItem, &LIBID_OPCAutomation>,
   public CComObjectRootEx<CComMultiThreadModel>
{
public:
	
	//==========================================================================
	// ATL Declarations

	BEGIN_COM_MAP(COPCItem)
		COM_INTERFACE_ENTRY(OPCItem)
	END_COM_MAP()

	DECLARE_NOT_AGGREGATABLE(OPCItem)

	//==========================================================================
	// Public Operators

	// Constructor
	COPCItem();

	// Destructor
	~COPCItem();
	
	//==========================================================================
	// Public Methods

	// intializes a new item object.
	HRESULT Initialize( 
		OPCHANDLE          hClient,
		OPCITEMDEF*        pItemDef,
		OPCITEMRESULT*     pResult,
		IOPCGroupStateMgt* pGroup,
		COPCGroup*         pParent 
	);

	// updates the item after a read.
	void Update(const OPCITEMSTATE* pItemState);

	// updates the item after a data update or a write.
	void Update(const VARIANT* pValue, const DWORD quality, const FILETIME timestamp);

	// the item id.
	const BSTR GetItemID() const {return m_itemID;}

	// the last known value for the item.
	const VARIANT* GetValue() const { return &m_value; }

	// the last known quality for the item value.
	DWORD GetQuality() const { return m_quality; }

	// the last known timestamp for the item value.
	const FILETIME& GetTimestamp() const {return m_timestamp;}

	// the item handle assigned by the remote server.
	const OPCHANDLE GetServerHandle() const { return m_server; }

	// the item handle assigned by the wrapper's caller.
	const OPCHANDLE GetClientHandle() const { return m_client; }

	// updates item handle assigned by the wrapper's caller.
	void SetClientHandle(OPCHANDLE clientHandle) { m_client = clientHandle; }

	// updates the active state (does not affect the remote server).
	void SetActiveState(BOOL active) { m_active = active; }
	
	// updates the data type (does not affect the remote server).
	void SetDataType(VARTYPE datatype) { m_datatype = datatype; }

	//==========================================================================
	// OPCItem

	STDMETHOD(get_Parent)(OPCGroup** ppParent);
	STDMETHOD(get_ClientHandle)(LONG* pClientHandle);
	STDMETHOD(put_ClientHandle)(LONG ClientHandle);
	STDMETHOD(get_ServerHandle)(LONG* pServerHandle);
	STDMETHOD(get_AccessPath)(BSTR* pAccessPath);
	STDMETHOD(get_AccessRights)(LONG* pAccessRights);
	STDMETHOD(get_ItemID)(BSTR* pItemID);
	STDMETHOD(get_IsActive)(VARIANT_BOOL* pIsActive);
	STDMETHOD(put_IsActive)(VARIANT_BOOL IsActive);
	STDMETHOD(get_RequestedDataType)(SHORT* pRequestedDataType);
	STDMETHOD(put_RequestedDataType)(SHORT RequestedDataType);
	STDMETHOD(get_Value)(VARIANT* pCurrentValue);
	STDMETHOD(get_Quality)(LONG* pQuality);
	STDMETHOD(get_TimeStamp)(DATE* pTimeStamp);
	STDMETHOD(get_CanonicalDataType)(SHORT* pCanonicalDataType);
	STDMETHOD(get_EUType)(SHORT* pEUType);
	STDMETHOD(get_EUInfo)(VARIANT* pEUInfo);
	STDMETHOD(Read)(SHORT Source, VARIANT* pValue, VARIANT* pQuality, VARIANT* pTimeStamp);
	STDMETHOD(Write)(VARIANT Value);
	
	//==========================================================================
	// Public Members

	DWORD MagicNumber;

private:

	//==========================================================================
	// Private Members


	// pointer to parent
	COPCGroup* m_pParent;

	// state variables.
	CComBSTR    m_itemID;
	CComBSTR    m_accessPath;
	BOOL        m_active;
	DWORD       m_accessRights;
	VARTYPE     m_datatype;
	VARTYPE     m_nativeDatatype;
	CComVariant m_value;
	DWORD       m_quality;
	FILETIME    m_timestamp;
	OPCHANDLE   m_client;
	OPCHANDLE   m_server;

	// the remote server.
	IOPCItemMgtPtr m_pOPCItem;
	IOPCSyncIOPtr  m_pOPCSyncIO;
};

//==============================================================================
// TYPE:    COPCItemObject
// PURPOSE: Creates a concrete class for COPCItem.

typedef CComObject<COPCItem> COPCItemObject;

#endif // __OPCItem_H__  