//==============================================================================
// TITLE: OPCGroups.h
//
// CONTENTS:
//
// A class that implements the DA Automation OPCGroups collection.
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
// 2003/11/11 RSA   Created from OPCAutoServer.h

#ifndef __OPCGroups_H__    
#define __OPCGroups_H__

#include "OPCGroup.h"

class COPCAutoServer;

//==============================================================================
// CLASS:   COPCGroup
// PURPOSE: Implements all interfaces supported by the OPCGroups object.

class ATL_NO_VTABLE COPCGroups :
   public CComDualImpl<IOPCGroups, &IID_IOPCGroups, &LIBID_OPCAutomation>,
   public IConnectionPointContainerImpl<COPCGroups>,
   public CProxyDIOPCGroupsEvent<COPCGroups>,
   public CComObjectRootEx<CComMultiThreadModel>
{
public:

	//==========================================================================
	// ATL Declarations

	BEGIN_COM_MAP(COPCGroups)
		COM_INTERFACE_ENTRY2(IDispatch, IOPCGroups)
		COM_INTERFACE_ENTRY(IOPCGroups)
		COM_INTERFACE_ENTRY_IMPL(IConnectionPointContainer)
	END_COM_MAP()

	BEGIN_CONNECTION_POINT_MAP(COPCGroups)
		CONNECTION_POINT_ENTRY(DIID_DIOPCGroupsEvent)
	END_CONNECTION_POINT_MAP()

	DECLARE_NOT_AGGREGATABLE(COPCGroups)

	//==========================================================================
	// Public Operators

	// Constructor
	COPCGroups();

	// Destructor
	~COPCGroups();
	
	//==========================================================================
	// Public Methods

	// saves a reference to the containing server.
	void Initialize(COPCAutoServer* pParent);

	// saves references to the remote server.
	void Connect(IUnknown* ipUnknown);

	// releases all references to the remote server.
	void Disconnect();
	
	//==========================================================================
	// OPCGroups

	STDMETHOD(get_Parent)(IOPCAutoServer** ppParent);
	STDMETHOD(get_DefaultGroupIsActive)(VARIANT_BOOL* pDefaultGroupIsActive);
	STDMETHOD(put_DefaultGroupIsActive)(VARIANT_BOOL DefaultGroupIsActive);
	STDMETHOD(get_DefaultGroupUpdateRate)(LONG* pDefaultGroupUpdateRate);
	STDMETHOD(put_DefaultGroupUpdateRate)(LONG DefaultGroupUpdateRate);
	STDMETHOD(get_DefaultGroupDeadband)(float* pDefaultGroupDeadband);
	STDMETHOD(put_DefaultGroupDeadband)(float DefaultGroupDeadband);
	STDMETHOD(get_DefaultGroupLocaleID)(LONG* pDefaultGroupLocaleID);
	STDMETHOD(put_DefaultGroupLocaleID)(LONG DefaultGroupLocaleID);
	STDMETHOD(get_DefaultGroupTimeBias)(LONG* pDefaultGroupTimeBias);
	STDMETHOD(put_DefaultGroupTimeBias)(LONG DefaultGroupTimeBias);
	STDMETHOD(get_Count)(LONG* pCount);
	STDMETHOD(get__NewEnum)(IUnknown** ppUnk);
	STDMETHOD(Item)(VARIANT ItemSpecifier, OPCGroup** ppGroup);
	STDMETHOD(Add)(VARIANT Name, OPCGroup** ppGroup);
	STDMETHOD(GetOPCGroup)(VARIANT ItemSpecifier, OPCGroup** ppGroup);
	STDMETHOD(RemoveAll)(void);
	STDMETHOD(Remove)(VARIANT ItemSpecifier);
	STDMETHOD(ConnectPublicGroup)(BSTR Name, OPCGroup** ppGroup);
	STDMETHOD(RemovePublicGroup)(VARIANT ItemSpecifier);

private:

	//==========================================================================
	// Private Members

	// the server containing the groups.
	COPCAutoServer* m_pParent;

	// state variables.
	VARIANT_BOOL   m_defaultActive;
	LONG           m_defaultUpdate;
	float          m_defaultDeadband;
	LONG           m_defaultLocale;
	LONG           m_defaultTimeBias;

	// the set of groups.
	list<COPCGroup*> m_groups;

	// the remote server.
	IOPCServerPtr m_pOPCServer;
};

//==============================================================================
// TYPE:    COPCGroupsObject
// PURPOSE: Creates a concrete class for COPCGroups.

typedef CComObject<COPCGroups> COPCGroupsObject;

#endif // __OPCGroups_H__    







