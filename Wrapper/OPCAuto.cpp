//==============================================================================
// TITLE: OPCAuto.cpp
//
// CONTENTS:
//
//  This file contains functions required for an in-process COM server.
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
//

#include "StdAfx.h"
#include "initguid.h"
#include "OPCAutoServer.h"
#include "OPCActivator.h"
#include "OPCTrace.h"

#ifdef _MERGE_PROXYSTUB
extern "C" HINSTANCE hProxyDll;
#endif

CComModule _Module;

BEGIN_OBJECT_MAP(ObjectMap)
   OBJECT_ENTRY(CLSID_OPCServer, COPCAutoServerObject)
   OBJECT_ENTRY(CLSID_OPCActivator, COPCActivatorObject)
END_OBJECT_MAP()

UINT OPCSTMFORMATDATATIME = 0;
UINT OPCSTMFORMATWRITECOMPLETE = 0;

/////////////////////////////////////////////////////////////////////////////
// DLL Entry Point

extern "C"
BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
	lpReserved;

	#ifdef _MERGE_PROXYSTUB
	
	if (!PrxDllMain(hInstance, dwReason, lpReserved))
	{
		return FALSE;
	}

	#endif

	if (dwReason == DLL_PROCESS_ATTACH)
	{
		OPCSTMFORMATDATATIME = RegisterClipboardFormat(_T("OPCSTMFORMATDATATIME"));
		OPCSTMFORMATWRITECOMPLETE = RegisterClipboardFormat(_T("OPCSTMFORMATWRITECOMPLETE"));
		
		_Module.Init(ObjectMap, hInstance);
		DisableThreadLibraryCalls(hInstance);

		GetTrace().Initialize(hInstance);
	}
	else if (dwReason == DLL_PROCESS_DETACH)
	{
		GetTrace().Uninitialize();

		_Module.Term();
	}

	return TRUE;
}

main()   // This is only to satisfy the linker when using MT lib
{ 
	return 0;
}

/////////////////////////////////////////////////////////////////////////////
// Used to determine whether the DLL can be unloaded by OLE

STDAPI DllCanUnloadNow(void)
{
	#ifdef _MERGE_PROXYSTUB
	
	if (PrxDllCanUnloadNow() != S_OK)
	{
		return S_FALSE;
	}
	
	#endif
	
	return (_Module.GetLockCount()==0) ? S_OK : S_FALSE;
}

/////////////////////////////////////////////////////////////////////////////
// Returns a class factory to create an object of the requested type

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
	#ifdef _MERGE_PROXYSTUB
	
	if (PrxDllGetClassObject(rclsid, riid, ppv) == S_OK)
	{
		return S_OK;
	}

	#endif

	return _Module.GetClassObject(rclsid, riid, ppv);
}

/////////////////////////////////////////////////////////////////////////////
// DllRegisterServer - Adds entries to the system registry

STDAPI DllRegisterServer(void)
{
	#ifdef _MERGE_PROXYSTUB

	HRESULT hRes = PrxDllRegisterServer();
	
	if (FAILED(hRes))
	{
		return hRes;
	}

	#endif

	// registers object, typelib and all interfaces in typelib
	return _Module.RegisterServer(TRUE);
}

/////////////////////////////////////////////////////////////////////////////
// DllUnregisterServer - Removes entries from the system registry

STDAPI DllUnregisterServer(void)
{
	#ifdef _MERGE_PROXYSTUB

	PrxDllUnregisterServer();
	
	#endif
	
	_Module.UnregisterServer();

	GetTrace().DeleteRegistryKeys();

	return S_OK;
}