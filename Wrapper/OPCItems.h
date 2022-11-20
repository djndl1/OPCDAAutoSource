//==============================================================================
// TITLE: OPCItems.h
//
// CONTENTS:
//
// A class that implements the DA Automation OPCItems collection.
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
// 2003/11/17 RSA   Created from OPCGroup.h.

#ifndef __OPCItems_H__ 
#define __OPCItems_H__

#include "OPCItem.h"
#include "OPCItems.h"
#include "OPCGroupEvent.h"

class COPCGroup;
class COPCGroups;

//==============================================================================
// CLASS:   COPCItems
// PURPOSE: Implements all interfaces supported by the OPCItems object.

class ATL_NO_VTABLE COPCItems :
   public CComDualImpl<OPCItems, &IID_OPCItems, &LIBID_OPCAutomation>,
   public CComObjectRootEx<CComMultiThreadModel>
{
public:
	
	//==========================================================================
	// ATL Declarations

	BEGIN_COM_MAP(COPCItems)
		COM_INTERFACE_ENTRY2(IDispatch, OPCItems)
		COM_INTERFACE_ENTRY(OPCItems)
	END_COM_MAP()

	DECLARE_NOT_AGGREGATABLE(OPCItem)

	//==========================================================================
	// Public Operators

	// Constructor
	COPCItems();

	// Destructor
	~COPCItems();

	//==========================================================================
	// Public Members

	// saves a reference to the containing group.
	void Initialize(COPCGroup* pParent);

	// saves a reference to the remote server.
	void Connect(IUnknown* ipUnknown);

	// releases all references to the remote server.
	void Disconnect();

	// removes all items in the connect.
	void RemoveAllItems();

	// finds the item object with the specified handle.
	COPCItem* LookupItem(OPCHANDLE ServerHandle);

	// returns a server handle that should be invalid.
	OPCHANDLE GetInvalidHandle();

	//==========================================================================
	// OPCItems

	STDMETHOD(get_Parent)(OPCGroup** ppParent);
	STDMETHOD(get_DefaultRequestedDataType)(SHORT* pDefaultRequestedDataType);
	STDMETHOD(put_DefaultRequestedDataType)(SHORT DefaultRequestedDataType);
	STDMETHOD(get_DefaultAccessPath)(BSTR* pDefaultAccessPath);
	STDMETHOD(put_DefaultAccessPath)(BSTR DefaultAccessPath);
	STDMETHOD(get_DefaultIsActive)(VARIANT_BOOL* pDefaultIsActive);
	STDMETHOD(put_DefaultIsActive)(VARIANT_BOOL DefaultIsActive);
	STDMETHOD(get_Count)(LONG* pCount);
	STDMETHOD(get__NewEnum)(IUnknown** ppUnk);
	STDMETHOD(Item)(VARIANT ItemSpecifier, OPCItem** ppItem);
	STDMETHOD(GetOPCItem)(LONG ServerHandle, OPCItem** ppItem);
	STDMETHOD(AddItem)(BSTR ItemID, LONG ClientHandle, OPCItem** ppItem);
	STDMETHOD(AddItems)(LONG NumItems, SAFEARRAY** ppItemIDs, SAFEARRAY** ppClientHandles, SAFEARRAY** ppServerHandles, SAFEARRAY ** ppErrors, VARIANT RequestedDataTypes, VARIANT AccessPaths);
	STDMETHOD(Remove)(LONG NumItems, SAFEARRAY** ppServerHandles, SAFEARRAY** ppErrors);
	STDMETHOD(Validate)(LONG NumItems, SAFEARRAY** ppItemIDs, SAFEARRAY** ppErrors,	VARIANT RequestedDataTypes, VARIANT AccessPaths);
	STDMETHOD(SetActive)(LONG NumItems,	SAFEARRAY** ppServerHandles, VARIANT_BOOL ActiveState, SAFEARRAY** ppErrors);
	STDMETHOD(SetClientHandles)(LONG NumItems, SAFEARRAY** ppServerHandles,	SAFEARRAY** ppClientHandles, SAFEARRAY** ppErrors);
	STDMETHOD(SetDataTypes)(LONG NumItems, SAFEARRAY** ppServerHandles,	SAFEARRAY** ppRequestedDataTypes, SAFEARRAY** ppErrors);


private:
	
	//==========================================================================
	// Private Members

	// initializes an array of item definition structures.
	OPCITEMDEF* CreateDefinitions(
		LONG       lCount,
		SAFEARRAY* psaItemIDs,
		SAFEARRAY* psaReqTypes,
		SAFEARRAY* psaAccessPaths
	);

	// frees an array of item definition and result structures.
	void FreeDefinitions(
		LONG           lCount,
		OPCITEMDEF*    pItemDefs,
		OPCITEMRESULT* pItemResults,
		HRESULT*       pErrors
	);

	// pointer to parent
	COPCGroup* m_pParent;
	
	// the collection of items.
	list<COPCItem*> m_items; 

	// state variables.
	VARTYPE      m_defaultDataType;
	CComBSTR     m_defaultAccessPath;
	VARIANT_BOOL m_defaultActive;

	// the remote server.
	IOPCGroupStateMgtPtr m_pOPCGroup;
	IOPCItemMgtPtr       m_pOPCItem;
};

//==============================================================================
// TYPE:    COPCItemsObject
// PURPOSE: Creates a concrete class for COPCItems.

typedef CComObject<COPCItems> COPCItemsObject;

#endif // __OPCItems_H__