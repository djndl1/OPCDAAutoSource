//==============================================================================
// TITLE: OPCAutoServer.h
//
// CONTENTS:
//
// A class that implements the DA Automation OPCServer object.
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

#ifndef __OPCAutoServer_H__      
#define __OPCAutoServer_H__

#include "OPCGroupEvent.h"

class COPCGroups;

#pragma warning( disable : 4996 ) // ATL::CComModule::UpdateRegistryClass deprecated.

//==============================================================================
// CLASS:   COPCAutoServer
// PURPOSE: Implements all interfaces supported by the OPCServer object.

class ATL_NO_VTABLE COPCAutoServer :
   public IOPCShutdown,
   public CComDualImpl<IOPCAutoServer, &IID_IOPCAutoServer, &LIBID_OPCAutomation>,
   public IConnectionPointContainerImpl<COPCAutoServer>,
   public CProxyDIOPCServerEvent<COPCAutoServer>,
   public CComObjectRootEx<CComMultiThreadModel>,
   public CComCoClass<COPCAutoServer, &CLSID_OPCServer>
{
public:

	//==========================================================================
	// ATL Declarations

	BEGIN_COM_MAP(COPCAutoServer)
		COM_INTERFACE_ENTRY(IOPCShutdown)
		COM_INTERFACE_ENTRY2(IDispatch, IOPCAutoServer)
		COM_INTERFACE_ENTRY(IOPCAutoServer)
		COM_INTERFACE_ENTRY_IMPL(IConnectionPointContainer)
	END_COM_MAP()

	BEGIN_CONNECTION_POINT_MAP(COPCAutoServer)
		CONNECTION_POINT_ENTRY(DIID_DIOPCServerEvent)
	END_CONNECTION_POINT_MAP()

	DECLARE_NOT_AGGREGATABLE(COPCAutoServer)
	DECLARE_REGISTRY(COPCAutoServer, _T("OPC.Automation.1"), _T("OPC.Automation"), IDS_PROJNAME, THREADFLAGS_BOTH)

	//==========================================================================
	// Public Operators

	// Constructor
	COPCAutoServer();
	
	// Destructor
	~COPCAutoServer();
	
	// FinalConstruct
	virtual HRESULT FinalConstruct();

	// FinalRelease
	virtual void FinalRelease();

	//==========================================================================
	// IOPCShutdown
	
	STDMETHOD(ShutdownRequest)(LPCWSTR szReason);
	
	//==========================================================================
	// IOPCAutoServer

    STDMETHOD(get_StartTime)(DATE* pStartTime);
    STDMETHOD(get_CurrentTime)(DATE* pCurrentTime);
    STDMETHOD(get_LastUpdateTime)(DATE* pLastUpdateTime);
    STDMETHOD(get_MajorVersion)(short* pMajorVersion);
    STDMETHOD(get_MinorVersion)(short* pMinorVersion);
    STDMETHOD(get_BuildNumber)(short* pBuildNumber);
    STDMETHOD(get_VendorInfo)(BSTR* pVendorInfo);
    STDMETHOD(get_ServerState)(LONG* pServerState);
    STDMETHOD(get_ServerName)(BSTR* pServerName);
    STDMETHOD(get_ServerNode)(BSTR* pServerNode);
    STDMETHOD(get_ClientName)(BSTR* pClientName);
    STDMETHOD(put_ClientName)(BSTR ClientName);
    STDMETHOD(get_LocaleID)(LONG* pLocaleID);
    STDMETHOD(put_LocaleID)(LONG LocaleID);
    STDMETHOD(get_Bandwidth)(LONG* pBandwidth);
    STDMETHOD(get_OPCGroups)(OPCGroups** ppGroups);
    STDMETHOD(get_PublicGroupNames)(VARIANT* pPublicGroups);
    STDMETHOD(GetOPCServers)(VARIANT Node, VARIANT* pOPCServers);
    STDMETHOD(Connect)(BSTR ProgID, VARIANT Node);
    STDMETHOD(Disconnect)(void);
	STDMETHOD(CreateBrowser)(OPCBrowser** ppBrowser);
    STDMETHOD(GetErrorString)(LONG ErrorCode, BSTR* ErrorString);
    STDMETHOD(QueryAvailableLocaleIDs)(VARIANT* LocaleIDs);
    STDMETHOD(QueryAvailableProperties)(BSTR ItemID, LONG* pCount, SAFEARRAY** PropertyIDs, SAFEARRAY** Descriptions, SAFEARRAY** ppDataTypes);
    STDMETHOD(GetItemProperties)(BSTR ItemID, LONG Count, SAFEARRAY** PropertyIDs, SAFEARRAY** PropertyValues, SAFEARRAY** Errors);
    STDMETHOD(LookupItemIDs)(BSTR ItemID, LONG Count, SAFEARRAY** PropertyIDs, SAFEARRAY ** NewItemIDs, SAFEARRAY ** Errors);

	//==========================================================================
	// Internal Methods
	
	STDMETHOD(Attach)(IUnknown* ipServer, BSTR ProgID, VARIANT NodeName);

private:
	
	//==========================================================================
	// Private Operators

	// the container for the the group.
	COPCGroups* m_pGroups;

	// state variables.
	CComBSTR m_name;
	CComBSTR m_node;
	CComBSTR m_client;
	LONG     m_localeID;
	DWORD    m_ShutdownConnection;

	// remote server interface pointers.
	IOPCServerPtr         m_pOPCServer;
	IOPCItemPropertiesPtr m_pOPCItemProperties;
	IOPCCommonPtr         m_pOPCCommon;
};

//==============================================================================
// TYPE:    COPCAutoServerObject
// PURPOSE: Creates a concrete class instance for COPCAutoServer.

typedef CComObject<COPCAutoServer> COPCAutoServerObject;

#endif // __OPCAutoServer_H__






