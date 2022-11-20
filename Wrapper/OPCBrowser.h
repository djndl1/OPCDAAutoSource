//==============================================================================
// TITLE: OPCBrowser.h
//
// CONTENTS:
//
// A class that implements the DA Automation OPCBrowser object.
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

#ifndef __OPCBrowser_H__               
#define __OPCBrowser_H__

//==============================================================================
// CLASS:   COPCBrowser
// PURPOSE: Implements all interfaces supported by the OPCBrowser object.

class ATL_NO_VTABLE COPCBrowser :
   public CComDualImpl<OPCBrowser, &IID_OPCBrowser, &LIBID_OPCAutomation>,
   public CComObjectRootEx<CComMultiThreadModel>
{
public:
	
	//==========================================================================
	// ATL Declarations

	BEGIN_COM_MAP(COPCBrowser)
		COM_INTERFACE_ENTRY(OPCBrowser)
	END_COM_MAP()
	
	DECLARE_NOT_AGGREGATABLE(OPCBrowser)
	
	//==========================================================================
	// Public Operators

	// Constructor
	COPCBrowser();

	// Destructor
	~COPCBrowser();
	
	//==========================================================================
	// Public Members
	
	// Initialize
	bool Initialize(IUnknown* ipUnknown) 
	{ 
		m_pOPCBrowser = ipUnknown; 
		return (!m_pOPCBrowser)?false:true;
	}

	//==========================================================================
	// OPCBrowser

	STDMETHOD(get_Organization)(LONG* pOrganization);
	STDMETHOD(get_Filter)(BSTR* pFilter);
	STDMETHOD(put_Filter)(BSTR Filter);
	STDMETHOD(get_DataType)(SHORT* pDataType);
	STDMETHOD(put_DataType)(SHORT DataType);
	STDMETHOD(get_AccessRights)(LONG* pAccessRights);
	STDMETHOD(put_AccessRights)(LONG AccessRights);
	STDMETHOD(get_CurrentPosition)(BSTR* pCurrentPosition);
	STDMETHOD(get_Count)(LONG* pCount);
	STDMETHOD(get__NewEnum)(IUnknown** ppUnk);
	STDMETHOD(Item)(VARIANT ItemSpecifier, BSTR* pItem);
	STDMETHOD(ShowBranches)();
	STDMETHOD(ShowLeafs)(VARIANT Flat);
	STDMETHOD(MoveUp)();
	STDMETHOD(MoveToRoot)();
	STDMETHOD(MoveDown)(BSTR Branch);
	STDMETHOD(MoveTo)(SAFEARRAY** ppBranches);
	STDMETHOD(GetItemID)(BSTR Leaf, BSTR* pItemID);
	STDMETHOD(GetAccessPaths)(BSTR ItemID, VARIANT* pAccessPaths);
	
private:
	
	//==========================================================================
	// Private Methods

	void ClearNames();

	//==========================================================================
	// Private Members

	// remote server.
	IOPCBrowseServerAddressSpacePtr m_pOPCBrowser;

	// state variables.
	CComBSTR m_filter;
	VARTYPE  m_dataType;
	LONG     m_accessRights;
	CComBSTR m_currentPosition;
	
	// the set of node names at the current location.
	list<CComBSTR*> m_names;
};

//==============================================================================
// TYPE:    COPCBrowserObject
// PURPOSE: Creates a concrete class for COPCBrowser.

typedef CComObject<COPCBrowser> COPCBrowserObject;

#endif // __OPCBrowser_H__