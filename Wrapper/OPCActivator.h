//==============================================================================
// TITLE: OPCActivator.h
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

#ifndef __OPCActivator_H__      
#define __OPCActivator_H__

#include "OPCAutoServer.h"

#pragma warning( disable : 4996 ) // ATL::CComModule::UpdateRegistryClass deprecated.

//==============================================================================
// CLASS:   COPCActivator
// PURPOSE: Implements all interfaces supported by the OPCServer object.

class ATL_NO_VTABLE COPCActivator :
   public CComDualImpl<IOPCActivator, &IID_IOPCActivator, &LIBID_OPCAutomation>,
   public CComObjectRootEx<CComMultiThreadModel>,
   public CComCoClass<COPCActivator, &CLSID_OPCActivator>
{
public:

	//==========================================================================
	// ATL Declarations

	BEGIN_COM_MAP(COPCActivator)
		COM_INTERFACE_ENTRY2(IDispatch, IOPCActivator)
		COM_INTERFACE_ENTRY(IOPCActivator)
	END_COM_MAP()

	DECLARE_NOT_AGGREGATABLE(COPCActivator)
	DECLARE_REGISTRY(COPCActivator, _T("OPC.Automation.Activator.1"), _T("OPC.Automation.Activator"), IDS_PROJNAME, THREADFLAGS_BOTH)

	//==========================================================================
	// Public Operators

	// Constructor
	COPCActivator() {}
	
	// Destructor
	~COPCActivator() {}
	
	//==========================================================================
	// IOPCActivator
	
	// Attach
	STDMETHOD(Attach)(IUnknown* ipServer, BSTR ProgID, VARIANT NodeName, IOPCAutoServer** ppWrapper)
	{
		if (ipServer == NULL || ppWrapper == NULL)
		{
			return E_INVALIDARG;
		}

		*ppWrapper = NULL;

		COPCAutoServerObject* pWrapper = NULL;
		COPCAutoServerObject::CreateInstance((COPCAutoServerObject**)&pWrapper);

		if (pWrapper == NULL)
		{
			return E_UNEXPECTED;
		}

		pWrapper->AddRef();

		HRESULT hResult = pWrapper->Attach(ipServer, ProgID, NodeName);

		if (FAILED(hResult))
		{
			pWrapper->Release();
			return hResult;
		}

		*ppWrapper = pWrapper;
		return S_OK;
	}
};

//==============================================================================
// TYPE:    COPCActivatorObject
// PURPOSE: Creates a concrete class instance for COPCActivator.

typedef CComObject<COPCActivator> COPCActivatorObject;

#endif // __OPCActivator_H__






