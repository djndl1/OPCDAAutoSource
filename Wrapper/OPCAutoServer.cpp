//==============================================================================
// TITLE: OPCAutoServer.cpp
//
// CONTENTS:
// 
// The implementation of the OPCServer object,
//
// (c) Copyright 2003-2004 The OPC Foundation
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

#include "StdAfx.h"
#include "OPCAutoServer.h"
#include "OPCUtils.h"
#include "OPCTrace.h"
#include "OPCBrowser.h"
#include "OPCGroups.h"

//==============================================================================
// Local Declarations

#define BLOCK_SIZE 100

//==============================================================================
// CheckForLegacyWrapper
//   Checks for a local connect, if the server has an InProc configuration
//   and if this InProc Server is the OPCDisp (the old automation wrapper).

static BOOL CheckForLegacyWrapper(REFCLSID clsid)
{
	USES_CONVERSION;

	BOOL bLegacyWrapper = FALSE;

	LPOLESTR wszClsidString = NULL;
	StringFromCLSID(clsid, &wszClsidString);

	HKEY hKey = NULL;

	if (RegOpenKey(HKEY_CLASSES_ROOT, _T("CLSID"), &hKey) == ERROR_SUCCESS)
	{
		HKEY hClsidKey = NULL;

		if (RegOpenKey(hKey, OLE2T(wszClsidString), &hClsidKey) == ERROR_SUCCESS)
		{
			TCHAR tsInprocName[MAX_PATH+1];
			memset(tsInprocName, 0, sizeof(tsInprocName));
			LONG lSize = sizeof(tsInprocName);

			if (RegQueryValue(hClsidKey, _T("InprocServer32"), tsInprocName, &lSize) == ERROR_SUCCESS)
			{
				int iIndex = OpcReverseFind(tsInprocName, _T("\\"));

				if (iIndex != -1)
				{
					LPTSTR szName = OpcSubStr(tsInprocName, iIndex);

					if (_tcsicmp(_T("opcdisp.dll"), szName) == 0)
					{
						bLegacyWrapper = TRUE;
					}

					CoTaskMemFree(szName);
				}
			}

			RegCloseKey(hClsidKey);
		}
			
		RegCloseKey(hKey);
	}

	CoTaskMemFree(wszClsidString);

	return bLegacyWrapper;
}

//==============================================================================
// Constructor

COPCAutoServer::COPCAutoServer()
{
	m_localeID           = 0;
	m_ShutdownConnection = 0;

	m_pGroups = NULL;
	COPCGroupsObject::CreateInstance((COPCGroupsObject**)&m_pGroups);
	_ASSERTE(m_pGroups != NULL);

	m_pGroups->Initialize(this);
	m_pGroups->AddRef();
}

//==============================================================================
// Destructor

COPCAutoServer::~COPCAutoServer()
{
}

//==============================================================================
// FinalConstruct

HRESULT COPCAutoServer::FinalConstruct()
{
	TRACE_SPECIAL("COPCAutoServer::FinalConstruct");
	return S_OK;
}

//==============================================================================
// FinalRelease

void COPCAutoServer::FinalRelease()
{
	TRACE_SPECIAL("COPCAutoServer::FinalRelease");
	
	Disconnect();
	m_pGroups->Release();
}

//==============================================================================
// ShutdownRequest Callback

STDMETHODIMP COPCAutoServer::ShutdownRequest(LPCWSTR szReason)
{
	TRACE_FUNCTION_ENTER("COPCAutoServer::ShutdownRequest");
	
	CheckByValInArg(szReason);
	
	CComBSTR bstrReason(szReason);
	Fire_ServerShutDown(bstrReason);
	
   	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// StartTime Property

STDMETHODIMP COPCAutoServer::get_StartTime(DATE* pStartTime)
{
	TRACE_FUNCTION_ENTER("COPCAutoServer::get_StartTime");
	
	// check arguments.
	if (!CheckByRefOutArg(pStartTime)) return E_INVALIDARG; 

	// check if server is connected.
	if (!m_pOPCServer)
	{
		TRACE_INTERFACE_ERROR(IOPCServer);
		return E_FAIL;
	}

	// initialize return parameters.
	*pStartTime = 0; 

	// fetch server status.
	TRACE_METHOD_CALL(GetStatus);

	OPCSERVERSTATUS* pStatus = NULL;

	HRESULT hResult = m_pOPCServer->GetStatus(&pStatus);

	if (FAILED(hResult))
	{
	   	TRACE_METHOD_ERROR(GetStatus, hResult);
		return hResult;
	}

	// convert local time.
	*pStartTime = GetUtcTime(pStatus->ftStartTime);

	// free memory.
	CoTaskMemFree(pStatus->szVendorInfo);
	CoTaskMemFree(pStatus);
	
	// log the return parameters.
	TraceByRefOutArg(pStartTime);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// CurrentTime Property

STDMETHODIMP COPCAutoServer::get_CurrentTime(DATE* pCurrentTime)
{
	TRACE_FUNCTION_ENTER("COPCAutoServer::get_CurrentTime");
	
	// check arguments.
	if (!CheckByRefOutArg(pCurrentTime)) return E_INVALIDARG; 

	// check if server is connected.
	if (!m_pOPCServer)
	{
		TRACE_INTERFACE_ERROR(IOPCServer);
		return E_FAIL;
	}

	// initialize return parameters.
	*pCurrentTime = 0; 

	// fetch server status.
	TRACE_METHOD_CALL(GetStatus);

	OPCSERVERSTATUS* pStatus = NULL;

	HRESULT hResult = m_pOPCServer->GetStatus(&pStatus);

	if (FAILED(hResult))
	{
	   	TRACE_METHOD_ERROR(GetStatus, hResult);
		return hResult;
	}

	// convert local time.
	*pCurrentTime = GetUtcTime(pStatus->ftCurrentTime);

	// free memory.
	CoTaskMemFree(pStatus->szVendorInfo);
	CoTaskMemFree(pStatus);

	// log the return parameters.
	TraceByRefOutArg(pCurrentTime);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// LastUpdateTime Property

STDMETHODIMP COPCAutoServer::get_LastUpdateTime(DATE* pLastUpdateTime)
{
	TRACE_FUNCTION_ENTER("COPCAutoServer::get_LastUpdateTime");

	// check arguments.
	if (!CheckByRefOutArg(pLastUpdateTime)) return E_INVALIDARG; 

	// check if server is connected.
	if (!m_pOPCServer)
	{	   	
		TRACE_INTERFACE_ERROR(IOPCServer);
		return E_FAIL;
	}

	// initialize return parameters.
	*pLastUpdateTime = 0; 

	// fetch server status.
	TRACE_METHOD_CALL(GetStatus);

	OPCSERVERSTATUS* pStatus = NULL;

	HRESULT hResult = m_pOPCServer->GetStatus(&pStatus);

	if (FAILED(hResult))
	{
	   	TRACE_METHOD_ERROR(GetStatus, hResult);
		return hResult;
	}

	// convert local time.
	*pLastUpdateTime = GetUtcTime(pStatus->ftLastUpdateTime);

	// free memory.
	CoTaskMemFree(pStatus->szVendorInfo);
	CoTaskMemFree(pStatus);

	// log the return parameters.
	TraceByRefOutArg(pLastUpdateTime);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// MajorVersion Property

STDMETHODIMP COPCAutoServer::get_MajorVersion(short* pMajorVersion)
{
	TRACE_FUNCTION_ENTER("COPCAutoServer::get_MajorVersion");

	// check arguments.
	if (!CheckByRefOutArg(pMajorVersion)) return E_INVALIDARG; 

	// check if server is connected.
	if (!m_pOPCServer)
	{
		TRACE_INTERFACE_ERROR(IOPCServer);
		return E_FAIL;
	}

	// initialize return parameters.
	*pMajorVersion = 0; 

	// fetch server status.
	TRACE_METHOD_CALL(GetStatus);

	OPCSERVERSTATUS* pStatus = NULL;

	HRESULT hResult = m_pOPCServer->GetStatus(&pStatus);

	if (FAILED(hResult))
	{
	   	TRACE_METHOD_ERROR(GetStatus, hResult);
		return hResult;
	}

	*pMajorVersion = pStatus->wMajorVersion;

	// free memory.
	CoTaskMemFree(pStatus->szVendorInfo);
	CoTaskMemFree(pStatus);

	// log the return parameters.
	TraceByRefOutArg(pMajorVersion);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// MinorVersion Property

STDMETHODIMP COPCAutoServer::get_MinorVersion(short* pMinorVersion)
{
	TRACE_FUNCTION_ENTER("COPCAutoServer::get_MinorVersion");

	// check arguments.
	if (!CheckByRefOutArg(pMinorVersion)) return E_INVALIDARG; 

	// check if server is connected.
	if (!m_pOPCServer)
	{
		TRACE_INTERFACE_ERROR(IOPCServer);
		return E_FAIL;
	}

	// initialize return parameters.
	*pMinorVersion = 0; 

	// fetch server status.
	TRACE_METHOD_CALL(GetStatus);

	OPCSERVERSTATUS* pStatus = NULL;

	HRESULT hResult = m_pOPCServer->GetStatus(&pStatus);

	if (FAILED(hResult))
	{
	   	TRACE_METHOD_ERROR(GetStatus, hResult);
		return hResult;
	}

	*pMinorVersion = pStatus->wMinorVersion;

	// free memory.
	CoTaskMemFree(pStatus->szVendorInfo);
	CoTaskMemFree(pStatus);

	// log the return parameters.
	TraceByRefOutArg(pMinorVersion);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// BuildNumber Property

STDMETHODIMP COPCAutoServer::get_BuildNumber(short* pBuildNumber)
{
	TRACE_FUNCTION_ENTER("COPCAutoServer::get_BuildNumber");

	// check arguments.
	if (!CheckByRefOutArg(pBuildNumber)) return E_INVALIDARG; 

	// check if server is connected.
	if (!m_pOPCServer)
	{
		TRACE_INTERFACE_ERROR(IOPCServer);
		return E_FAIL;
	}

	// initialize return parameters.
	*pBuildNumber = 0; 

	// fetch server status.
	TRACE_METHOD_CALL(GetStatus);

	OPCSERVERSTATUS* pStatus = NULL;

	HRESULT hResult = m_pOPCServer->GetStatus(&pStatus);

	if (FAILED(hResult))
	{
	   	TRACE_METHOD_ERROR(GetStatus, hResult);
		return hResult;
	}

	*pBuildNumber = pStatus->wBuildNumber;

	// free memory.
	CoTaskMemFree(pStatus->szVendorInfo);
	CoTaskMemFree(pStatus);

	// log the return parameters.
	TraceByRefOutArg(pBuildNumber);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// VendorInfo Property

STDMETHODIMP COPCAutoServer::get_VendorInfo(BSTR* pVendorInfo)
{
	USES_CONVERSION;

	TRACE_FUNCTION_ENTER("COPCAutoServer::get_VendorInfo");

	// check arguments.
	if (!CheckByRefOutArg(pVendorInfo)) return E_INVALIDARG; 

	// check if server is connected.
	if (!m_pOPCServer)
	{
		TRACE_INTERFACE_ERROR(IOPCServer);
		return E_FAIL;
	}

	// initialize return parameters.
	*pVendorInfo = NULL; 

	// fetch server status.
	TRACE_METHOD_CALL(GetStatus);

	OPCSERVERSTATUS* pStatus = NULL;

	HRESULT hResult = m_pOPCServer->GetStatus(&pStatus);

	if (FAILED(hResult))
	{
	   	TRACE_METHOD_ERROR(GetStatus, hResult);
		return hResult;
	}

	// create copy of string.
	*pVendorInfo = SysAllocString(pStatus->szVendorInfo);

	// free memory.
	CoTaskMemFree(pStatus->szVendorInfo);
	CoTaskMemFree(pStatus);

	// log the return parameters.
	TraceByRefOutArg(pVendorInfo);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// ServerState Property

STDMETHODIMP COPCAutoServer::get_ServerState(LONG* pServerState)
{   
	TRACE_FUNCTION_ENTER("COPCAutoServer::get_ServerState");

	// check arguments.
	if (!CheckByRefOutArg(pServerState)) return E_INVALIDARG; 
	
	// initialize return parameters.
	*pServerState = OPCDisconnected; 

	// check if server is connected.
	if (m_pOPCServer)
	{		
		// fetch server status.
		TRACE_METHOD_CALL(GetStatus);

		OPCSERVERSTATUS* pStatus = NULL;

		HRESULT hResult = m_pOPCServer->GetStatus(&pStatus);

		if (FAILED(hResult))
		{
	   		TRACE_METHOD_ERROR(GetStatus, hResult);
			return hResult;
		}    

		// convert to the DA automation server state values.
		switch (pStatus->dwServerState)
		{
			case OPC_STATUS_RUNNING:   { *pServerState = OPCRunning;   break; }
			case OPC_STATUS_FAILED:    { *pServerState = OPCFailed;    break; }
			case OPC_STATUS_NOCONFIG:  { *pServerState = OPCNoconfig;  break; }
			case OPC_STATUS_SUSPENDED: { *pServerState = OPCSuspended; break; }
			case OPC_STATUS_TEST:      { *pServerState = OPCTest;      break; }
			default:                   { *pServerState = OPCFailed;    break; }

		}

		// free memory.
		CoTaskMemFree(pStatus->szVendorInfo);
		CoTaskMemFree(pStatus);
	}

	// log the return parameters.
	TraceByRefOutArg(pServerState);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// Bandwidth Property

STDMETHODIMP COPCAutoServer::get_Bandwidth(LONG* pBandwidth)
{
	TRACE_FUNCTION_ENTER("COPCAutoServer::get_Bandwidth");

	// check arguments.
	if (!CheckByRefOutArg(pBandwidth)) return E_INVALIDARG; 

	// check if server is connected.
	if (!m_pOPCServer)
	{
		TRACE_INTERFACE_ERROR(IOPCServer);
		return E_FAIL;
	}

	// initialize return parameters.
	*pBandwidth = 0; 

	// fetch server status.
	TRACE_METHOD_CALL(GetStatus);

	OPCSERVERSTATUS* pStatus = NULL;

	HRESULT hResult = m_pOPCServer->GetStatus(&pStatus);

	if (FAILED(hResult))
	{
	   	TRACE_METHOD_ERROR(GetStatus, hResult);
		return hResult;
	}

	*pBandwidth = (LONG)pStatus->dwBandWidth;

	// free memory.
	CoTaskMemFree(pStatus->szVendorInfo);
	CoTaskMemFree(pStatus);

	// log the return parameters.
	TraceByRefOutArg(pBandwidth);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// ServerName Property

STDMETHODIMP COPCAutoServer::get_ServerName(BSTR* pServerName)
{
	USES_CONVERSION;

	TRACE_FUNCTION_ENTER("COPCAutoServer::get_ServerName");

	// check arguments.
	if (!CheckByRefOutArg(pServerName)) return E_INVALIDARG; 

	// copy server name.
	*pServerName = m_name.Copy();

	// log the return parameters.
	TraceByRefOutArg(pServerName);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// ServerNode Property

STDMETHODIMP COPCAutoServer::get_ServerNode(BSTR* pServerNode)
{
	USES_CONVERSION;

	TRACE_FUNCTION_ENTER("COPCAutoServer::get_ServerNode");

	// check arguments.
	if (!CheckByRefOutArg(pServerNode)) return E_INVALIDARG; 

	// copy server node name.
	*pServerNode = m_node.Copy();

	// log the return parameters.
	TraceByRefOutArg(pServerNode);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// ClientName Property

STDMETHODIMP COPCAutoServer::get_ClientName(BSTR* pClientName)
{
	USES_CONVERSION;

	TRACE_FUNCTION_ENTER("COPCAutoServer::get_ClientName");

	// check arguments.
	if (!CheckByRefOutArg(pClientName)) return E_INVALIDARG; 

	// copy client name.
	*pClientName = m_client.Copy();

	// log the return parameters.
	TraceByRefOutArg(pClientName);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

STDMETHODIMP COPCAutoServer::put_ClientName(BSTR ClientName)
{
	USES_CONVERSION;

	TRACE_FUNCTION_ENTER("COPCAutoServer::put_ClientName");

	// check arguments.
	if (!CheckByValInArg(ClientName)) return E_INVALIDARG; 

	// update client name,
	m_client = ClientName;

	// check if server is connected.
	if (m_pOPCCommon != NULL)
	{
		// set client name on the server.
		TRACE_METHOD_CALL(SetClientName);

		HRESULT hResult = m_pOPCCommon->SetClientName(m_client);

		if (FAILED(hResult))
		{
	   		TRACE_METHOD_ERROR(SetClientName, hResult);
			return hResult;
		}
	}

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// LocaleID Property

STDMETHODIMP COPCAutoServer::get_LocaleID(LONG* pLocaleID)
{	
	TRACE_FUNCTION_ENTER("COPCAutoServer::get_LocaleID");

	// check arguments.
	if (!CheckByRefOutArg(pLocaleID)) return E_INVALIDARG; 

	*pLocaleID = m_localeID;

	// check if server is connected.
	if (m_pOPCCommon != NULL)
	{
		// update cached locale id with value in server.
		TRACE_METHOD_CALL(GetLocaleID);

		HRESULT hResult = m_pOPCCommon->GetLocaleID((LCID*)&m_localeID);
		
		if (SUCCEEDED(hResult))
		{
			*pLocaleID = m_localeID;
		}
		else
		{
	   		TRACE_METHOD_ERROR(GetLocaleID, hResult);
		}
	}

	// log the return parameters.
	TraceByRefOutArg(pLocaleID);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

STDMETHODIMP COPCAutoServer::put_LocaleID(LONG LocaleID)
{
	TRACE_FUNCTION_NAME("OPCAutoServer::put_LocaleID");

	// check arguments.
	if (!CheckByValInArg(LocaleID)) return E_INVALIDARG; 

	m_localeID = LocaleID;
	
	// check if server is connected.
	if (m_pOPCCommon != NULL)
	{
		// update server's locale id.
		TRACE_METHOD_CALL(SetLocaleID);

		HRESULT hResult = m_pOPCCommon->SetLocaleID(m_localeID);
		
		if (FAILED(hResult))
		{
	   		TRACE_METHOD_ERROR(SetLocaleID, hResult);
			return hResult;
		}
	}

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// OPCGroups Property

STDMETHODIMP COPCAutoServer::get_OPCGroups(OPCGroups** ppGroups)
{
	TRACE_FUNCTION_ENTER("COPCAutoServer::get_OPCGroups");

	// check arguments.
	if (!CheckByRefOutArg(ppGroups)) return E_INVALIDARG; 

	// initialize return parameters.
	*ppGroups = NULL;
    
	// get Interface pointer
	HRESULT hResult = m_pGroups->QueryInterface(IID_IOPCGroups, (void**)ppGroups);

	if (FAILED(hResult))
	{		
		TRACE_INTERFACE_ERROR(IOPCGroups);
		return hResult;
	}

	_ASSERTE(*ppGroups);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// PublicGroupNames Property

STDMETHODIMP COPCAutoServer::get_PublicGroupNames(VARIANT* pPublicGroups)
{	
	TRACE_FUNCTION_ENTER("COPCAutoServer::get_PublicGroupNames");

	// check arguments.
	if (!CheckByRefOutArg(pPublicGroups)) return E_INVALIDARG; 

	TRACE_ERROR("%s, 'E_NOTIMPL'", TRACE_FUNCTION);
	return E_NOTIMPL;
}

//==============================================================================
// CreateEnumerator

static IOPCServerList2* CreateEnumerator(const CLSID& clsid, LPCTSTR szNodeName)
{
	USES_CONVERSION;

	TRACE_FUNCTION_ENTER("CCOPCAutoServer::CreateEnumerator");

	// connect to server on local host by default.
	DWORD         clsctx = CLSCTX_LOCAL_SERVER;
	COSERVERINFO* sinptr = NULL;

	// initialize the server info only for remote nodes.
	COSERVERINFO  sin;

	if (szNodeName != NULL && szNodeName[0] != NULL)
	{
		sinptr = &sin;
		clsctx = CLSCTX_REMOTE_SERVER;

		sin.pwszName    = T2OLE(szNodeName);
		sin.pAuthInfo   = NULL;
		sin.dwReserved1 = NULL;
		sin.dwReserved2 = NULL;
	}

	TRACE_MISC("Creating a server enumerator on Node '%s'", szNodeName);
	
	// set up mqi
	MULTI_QI mqi;

	mqi.pIID = &IID_IOPCServerList2;
	mqi.hr   = S_OK;
	mqi.pItf = NULL;

	// create COM server.
	TRACE_METHOD_CALL(CoCreateInstanceEx);

	HRESULT hResult = CoCreateInstanceEx(clsid, NULL, clsctx, sinptr, 1, &mqi);

	if (FAILED(hResult))
	{
	   	TRACE_METHOD_ERROR(CoCreateInstanceEx, hResult);
		return NULL;
	}

	// check if querty interface failed.
	if (FAILED(mqi.hr))
	{
	   	TRACE_METHOD_ERROR(mqi.hr, hResult);
		return NULL;
	}
	
	TRACE_FUNCTION_LEAVE();
	return (IOPCServerList2*)mqi.pItf;
}

//==============================================================================
// GetOPCServers Method

STDMETHODIMP COPCAutoServer::GetOPCServers(VARIANT Node, VARIANT* pOPCServers)
{
	USES_CONVERSION;

	TRACE_FUNCTION_ENTER("COPCAutoServer::GetOPCServers");

	// check arguments.
	if (!CheckByValInVariantArg(Node, VT_BSTR, true)) return E_INVALIDARG;
	if (!CheckByRefOutArg(pOPCServers))               return E_INVALIDARG;

	// declare list containing prog ids. 
	list<CComBSTR*> cProgIDs;

	// get node name from variant.
	LPTSTR tsNodeName = NULL;

	if ((Node.vt == VT_BSTR) && (Node.bstrVal != NULL) && (Node.bstrVal[0] != 0))
	{
		tsNodeName = OLE2T(Node.bstrVal);
	}

	HRESULT hResult = E_FAIL;

	// connect to OPC server enumerator.
	IOPCServerList2* ipEnumerator = CreateEnumerator(CLSID_OpcServerList, tsNodeName);

	if (ipEnumerator != NULL)
	{		
		// enumerate all DA 1.0a and DA 2.0 servers.
		IOPCEnumGUID* ipEnumGUID = NULL;

		CATID cCatids[2];

		cCatids[0] = CATID_OPCDAServer10;
		cCatids[1] = CATID_OPCDAServer20;
		
		TRACE_METHOD_CALL(EnumClassesOfCategories);

		hResult = ipEnumerator->EnumClassesOfCategories(2, cCatids, 0, NULL, &ipEnumGUID);

		if (SUCCEEDED(hResult))
		{
			CLSID pClsids[BLOCK_SIZE];
			memset(pClsids, 0, sizeof(pClsids));

			do
			{
				// fetch next clsid.
				DWORD dwCount = 0;

				hResult = ipEnumGUID->Next(BLOCK_SIZE, pClsids, &dwCount);

				// check if end of list encountered.
				if (FAILED(hResult) || dwCount == 0)
				{
					hResult = S_OK;
					break;
				}

				for (DWORD ii = 0; ii < dwCount; ii++)
				{
					// fetch server prog id from the server.
					LPWSTR szProgID       = NULL;
					LPWSTR szUserType     = NULL;
					LPWSTR szVerIndProgID = NULL;

					hResult = ipEnumerator->GetClassDetails(pClsids[ii], &szProgID, &szUserType, &szVerIndProgID);

					// add prog id to list of servers.
					if (SUCCEEDED(hResult))
					{
						cProgIDs.push_back(new CComBSTR(szProgID));

						CoTaskMemFree(szProgID);
						CoTaskMemFree(szUserType);
						CoTaskMemFree(szVerIndProgID);
					}
				}
			}
			while (SUCCEEDED(hResult));

			// release guid enumerator.
			ipEnumGUID->Release();
		}

		// release server enumerator.
		ipEnumerator->Release();
		ipEnumerator = NULL;
	}

	// search the registery if OpcEnum did not work.
	if (FAILED(hResult))
	{			
		HKEY hKey = NULL;

		TRACE_METHOD_CALL(RegConnectRegistry);

		DWORD dwResult = RegConnectRegistry(tsNodeName, HKEY_CLASSES_ROOT, &hKey);
		
		if (dwResult != ERROR_SUCCESS)
		{
	   		TRACE_METHOD_ERROR(RegConnectRegistry, GetLastError());
			return E_FAIL;
		}

		TCHAR tsKey[MAX_PATH+1];
		memset(tsKey, 0, sizeof(tsKey));
		
		for (LONG ii = 0; RegEnumKey(hKey, ii, tsKey, MAX_PATH) == ERROR_SUCCESS; ii++)
		{
			HKEY hProgIDKey = NULL;

			if (RegOpenKey(hKey, tsKey, &hProgIDKey) == ERROR_SUCCESS)
			{				
				HKEY hOPCKey = NULL;	

				if (RegOpenKeyEx(hProgIDKey, _T("OPC"), 0, KEY_READ, &hOPCKey) == ERROR_SUCCESS)
				{					
					cProgIDs.push_back(new CComBSTR(tsKey)); 
					RegCloseKey(hOPCKey);
				}

				RegCloseKey(hProgIDKey);
			}
		}

		RegCloseKey(hKey);
	}

	// copy list of server prog ids into a SAFEARRAY.
	SAFEARRAYBOUND cBounds;
	cBounds.lLbound   = 1;
	cBounds.cElements = cProgIDs.size();

	SAFEARRAY* pArray = SafeArrayCreate(VT_BSTR, 1, &cBounds);

	if (pArray == NULL)
	{
	   	TRACE_ERROR("%s, 'pArray == NULL'", TRACE_FUNCTION);
		return E_OUTOFMEMORY;
	}

	list<CComBSTR*>::iterator iter = cProgIDs.begin();

	for (LONG ii = 1; ii <= (LONG)cBounds.cElements && iter != cProgIDs.end(); ii++)
	{
		CComBSTR* pbstrProgID = *iter;
		SafeArrayPutElement(pArray, &ii, (void*)*pbstrProgID);
		delete pbstrProgID;
		iter++;
	}

	// return the SAFEARRAY to the caller.
	pOPCServers->vt     = VT_BSTR | VT_ARRAY;
	pOPCServers->parray = pArray;

	TraceByRefOutArg(pOPCServers);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// Connect Method

STDMETHODIMP COPCAutoServer::Connect(BSTR ProgID, VARIANT Node)
{
	USES_CONVERSION;

	TRACE_FUNCTION_ENTER("COPCAutoServer::Connect");

	// disconnect current server.
	Disconnect();

	// check arguments.
	if (!CheckByValInArg(ProgID))                     return E_INVALIDARG;
	if (!CheckByValInVariantArg(Node, VT_BSTR, true)) return E_INVALIDARG;

	// get node name from variant.
	LPTSTR tsNodeName = NULL;

	if ((Node.vt == VT_BSTR) && (Node.bstrVal != NULL) && (Node.bstrVal[0] != 0))
	{
		// check for well-known node names.
		if (wcsicmp(Node.bstrVal, L"localhost") != 0 && wcsicmp(Node.bstrVal, L"127.0.0.1") != 0)
		{
			tsNodeName = OLE2T(Node.bstrVal);
		}
	}
	
	// check for non-null node names that reference the local computer.
	if (tsNodeName != NULL)
	{
		// get own computer name
		TCHAR szLocalHost[MAX_COMPUTERNAME_LENGTH+1]; 	
		DWORD dwHostSize = sizeof(szLocalHost);

		if (GetComputerName(szLocalHost, &dwHostSize) == TRUE)
		{
			if (_tcsicmp(szLocalHost, tsNodeName) == 0)
			{
				tsNodeName = NULL;
			}
		}
	}

	// get clsid.
	CLSID cClsid = GUID_NULL;

	// connect to OPC server enumerator.
	IOPCServerList2* ipEnumerator = CreateEnumerator(CLSID_OpcServerList, tsNodeName);

	if (ipEnumerator!= NULL)
	{		
		TRACE_METHOD_CALL(CLSIDFromProgID);

		HRESULT hResult = ipEnumerator->CLSIDFromProgID(ProgID, &cClsid);

		// release server enumerator.
		ipEnumerator->Release();
		ipEnumerator = NULL;

		if (FAILED(hResult))
		{		
			TRACE_METHOD_ERROR(CLSIDFromProgID, hResult);
			return hResult;
		}
	}

	// lookup CLSID in registry if OpcEnum not available.
	else
	{
		HKEY hKey = NULL;

		TRACE_METHOD_CALL(RegConnectRegistry);

		DWORD dwResult = RegConnectRegistry(tsNodeName, HKEY_CLASSES_ROOT, &hKey);
		
		if (dwResult != ERROR_SUCCESS)
		{			
			TRACE_METHOD_ERROR(RegConnectRegistry, GetLastError());
			return E_FAIL;
		}

		TCHAR tsKey[MAX_PATH+1];
		memset(tsKey, 0, sizeof(tsKey));
		
		HKEY hProgIDKey = NULL;

		if (RegOpenKey(hKey, OLE2T(ProgID), &hProgIDKey) == ERROR_SUCCESS)
		{			
			HKEY hClsidKey = NULL;	

			if (RegOpenKeyEx(hProgIDKey, _T("CLSID"), 0, KEY_READ, &hClsidKey) == ERROR_SUCCESS)
			{	
				TCHAR tsClsid[MAX_PATH+1];
				memset(tsClsid, 0, sizeof(tsClsid));
				DWORD dwLength = sizeof(tsClsid);

				if (RegQueryValueEx(hClsidKey, NULL, NULL, NULL, (BYTE*)tsClsid, &dwLength) == ERROR_SUCCESS)
				{
					if (FAILED(CLSIDFromString(T2OLE(tsClsid), &cClsid)))
					{
						cClsid = GUID_NULL;
					}
				}
				
				RegCloseKey(hClsidKey);
			}

			RegCloseKey(hProgIDKey);
		}

		RegCloseKey(hKey);

		if (cClsid == GUID_NULL)
		{
			TRACE_ERROR("%s, 'Lookup CLSID for %s Failed'", TRACE_FUNCTION, OLE2T(ProgID));
			return E_FAIL;
		}
	}

	// connect to server on local host by default.
	DWORD         clsctx = CLSCTX_ALL;
	COSERVERINFO* sinptr = NULL;

	// initialize the server info only for remote nodes.
	COSERVERINFO  sin;

	if (tsNodeName != NULL)
	{
		sinptr = &sin;
		clsctx = CLSCTX_REMOTE_SERVER;

		sin.pwszName    = T2OLE(tsNodeName);
		sin.pAuthInfo   = NULL;
		sin.dwReserved1 = NULL;
		sin.dwReserved2 = NULL;
	}

	// prevent inproc server from being loaded if it happens to be the old automation wrapper.
	else if (CheckForLegacyWrapper(cClsid))
	{
		clsctx = CLSCTX_LOCAL_SERVER;
	}

	// set up mqi
	MULTI_QI mqi;

	mqi.pIID = &IID_IOPCServer;
	mqi.hr   = S_OK;
	mqi.pItf = NULL;

	// create COM server.
	TRACE_METHOD_CALL(CoCreateInstanceEx);

	HRESULT hResult = CoCreateInstanceEx(cClsid, NULL, clsctx, sinptr, 1, &mqi);

	if (FAILED(hResult))
	{
		TRACE_METHOD_ERROR(CoCreateInstanceEx, hResult);
		return hResult;
	}

	// check if query interface failed.
	if (FAILED(mqi.hr))
	{
		TRACE_METHOD_ERROR(mqi.hr, hResult);
		return hResult;
	}

	// save server name and node.
	m_name = ProgID;
	m_node = tsNodeName;

	// save reference to server.
	m_pOPCServer         = mqi.pItf;
	m_pOPCCommon         = mqi.pItf;
	m_pOPCItemProperties = mqi.pItf;

	// update groups object.
	m_pGroups->Connect(mqi.pItf);

	// update the cached locale - ignore errors.
	LONG localeID = 0;
	get_LocaleID(&localeID);
	
	mqi.pItf->Release();

	// create a shutdown connection point callback.	
	TRACE_METHOD_CALL(QueryInterface);

	IConnectionPointContainer* ipCPC = NULL;
	
	hResult = m_pOPCServer->QueryInterface(IID_IConnectionPointContainer, (LPVOID*)&ipCPC);
	
	if (SUCCEEDED(hResult))
	{
		IConnectionPoint* ipCP = NULL;
		
		TRACE_METHOD_CALL(FindConnectionPoint);

		hResult = ipCPC->FindConnectionPoint(IID_IOPCShutdown, &ipCP);

		ipCPC->Release();

		if (SUCCEEDED(hResult))
		{
			IUnknown* ipCallback = NULL;
			((COPCAutoServerObject*)this)->QueryInterface(IID_IUnknown, (LPVOID*)&ipCallback);
			
			TRACE_METHOD_CALL(Advise);

			hResult = ipCP->Advise(ipCallback, &m_ShutdownConnection);

			ipCP->Release();
			ipCallback->Release();

			if (FAILED(hResult))
			{
				TRACE_METHOD_ERROR(Advise, hResult);
				return hResult;
			}
		}
	}

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// Disconnect Method

STDMETHODIMP COPCAutoServer::Disconnect(void)
{	
	TRACE_FUNCTION_ENTER("COPCAutoServer::Disconnect");

	// do nothing is not connected.
	if (!m_pOPCServer)
	{
		TRACE_FUNCTION_LEAVE();
		return S_OK;
	}

	// destroy a shutdown connection point callback.
	IConnectionPointContainer* ipCPC = NULL;
	
	TRACE_METHOD_CALL(QueryInterface);
	HRESULT hResult = m_pOPCServer->QueryInterface(IID_IConnectionPointContainer, (LPVOID*)&ipCPC);
	
	if (SUCCEEDED(hResult))
	{
		IConnectionPoint* ipCP = NULL;

		TRACE_METHOD_CALL(FindConnectionPoint);
		hResult = ipCPC->FindConnectionPoint(IID_IOPCShutdown, &ipCP);

		if (SUCCEEDED(hResult))
		{			
			TRACE_METHOD_CALL(Unadvise);
			ipCP->Unadvise(m_ShutdownConnection);
			ipCP->Release();
		}
		
		ipCPC->Release();
	}

	// remove all groups.
	m_pGroups->Disconnect();

	// release servers
	m_pOPCServer.Release();
	m_pOPCCommon.Release();
	m_pOPCItemProperties.Release();

	// initialize member variables.
	m_name.Empty();
	m_node.Empty();
	
	m_ShutdownConnection = NULL;

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// Attach Method

STDMETHODIMP COPCAutoServer::Attach(IUnknown* ipUnknown, BSTR ProgID, VARIANT NodeName)
{
	USES_CONVERSION;

	TRACE_FUNCTION_ENTER("COPCAutoServer::Attach");

	// check arguments.
	if (!CheckByValInArg(ProgID))                         return E_INVALIDARG;
	if (!CheckByValInVariantArg(NodeName, VT_BSTR, true)) return E_INVALIDARG;

	// disconnect current server.
	Disconnect();

	// check interface.
	TRACE_METHOD_CALL(QueryInterface);

	IOPCServer* ipServer = NULL;

	HRESULT hResult = ipUnknown->QueryInterface(IID_IOPCServer, (void**)&ipServer);

	if (FAILED(hResult))
	{
		TRACE_METHOD_ERROR(QueryInterface, hResult);
		return hResult;
	}

	// save reference to server.
	m_pOPCServer         = ipServer;
	m_pOPCCommon         = ipServer;
	m_pOPCItemProperties = ipServer;

	// update groups object.
	m_pGroups->Connect(ipServer);

	// save server name and node.
	m_name = ProgID;

	if ((NodeName.vt == VT_BSTR) && (NodeName.bstrVal != NULL) && (NodeName.bstrVal[0] != 0))
	{
		m_node = NodeName.bstrVal;
	}

	// update the cached locale - ignore errors.
	LONG localeID = 0;
	get_LocaleID(&localeID);
	
	ipServer->Release();

	// create a shutdown connection point callback.	
	TRACE_METHOD_CALL(QueryInterface);

	IConnectionPointContainer* ipCPC = NULL;
	
	hResult = m_pOPCServer->QueryInterface(IID_IConnectionPointContainer, (LPVOID*)&ipCPC);
	
	if (SUCCEEDED(hResult))
	{
		IConnectionPoint* ipCP = NULL;
		
		TRACE_METHOD_CALL(FindConnectionPoint);

		hResult = ipCPC->FindConnectionPoint(IID_IOPCShutdown, &ipCP);

		ipCPC->Release();

		if (SUCCEEDED(hResult))
		{
			IUnknown* ipCallback = NULL;
			((COPCAutoServerObject*)this)->QueryInterface(IID_IUnknown, (LPVOID*)&ipCallback);
			
			TRACE_METHOD_CALL(Advise);

			hResult = ipCP->Advise(ipCallback, &m_ShutdownConnection);

			ipCallback->Release();
			ipCP->Release();

			if (FAILED(hResult))
			{
				TRACE_METHOD_ERROR(Advise, hResult);
				return hResult;
			}
		}
	}

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// CreateBrowser Method

STDMETHODIMP COPCAutoServer::CreateBrowser(OPCBrowser** ppBrowser)
{
	TRACE_FUNCTION_ENTER("COPCAutoServer::CreateBrowser");
	
	// check arguments.
	if (!CheckByRefOutArg(ppBrowser)) return E_INVALIDARG;

	// check if server is connected.
	if (!m_pOPCServer)
	{
		TRACE_INTERFACE_ERROR(IOPCServer);
		return E_FAIL;
	}

	// create a new browser object,
	COPCBrowserObject* pBrowser = NULL;
	COPCBrowserObject::CreateInstance(&pBrowser);

	if (!pBrowser->Initialize(m_pOPCServer))
	{
		delete pBrowser;
	   	TRACE_ERROR("%s, '!pBrowser->Initialize'", TRACE_FUNCTION);
		return E_FAIL;
	}

	// get interface pointer
	pBrowser->QueryInterface(IID_OPCBrowser, (void**)ppBrowser);
	_ASSERTE(*ppBrowser);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// GetErrorString Method

STDMETHODIMP COPCAutoServer::GetErrorString(LONG ErrorCode, BSTR* ErrorString)
{
	USES_CONVERSION;

	TRACE_FUNCTION_ENTER("COPCAutoServer::GetErrorString");
	
	// check arguments.
	if (!CheckByValInArg(ErrorCode))    return E_INVALIDARG;
	if (!CheckByRefOutArg(ErrorString)) return E_INVALIDARG;

	*ErrorString = NULL;


	LPTSTR tsBuffer  = NULL;
	DWORD  dwBufSize = 0;

	// try to resolve the error code and local ID locally to make no unnecessary server calls.
	HRESULT hResult = ::GetErrorString(ErrorCode, m_localeID, false,  &tsBuffer);
	
	if (SUCCEEDED(hResult))
	{
		*ErrorString = SysAllocString(T2OLE(tsBuffer));
		CoTaskMemFree(tsBuffer);

	   	TRACE_MISC(_T("%s -> Message : [0x%08X] '%s'"), TRACE_FUNCTION, ErrorCode, OLE2T(*ErrorString));
		TRACE_FUNCTION_LEAVE();
		return S_OK;
	}

	// check if the server is connected.
	if (!m_pOPCServer)
	{
		TRACE_INTERFACE_ERROR(IOPCServer);
		return E_NOINTERFACE;
	}

	// call to server to get an OPC error string.
    LPWSTR wszString = NULL;
    
	TRACE_METHOD_CALL(GetErrorString);
	hResult = m_pOPCServer->GetErrorString(ErrorCode, m_localeID, &wszString);

	if (SUCCEEDED(hResult))
	{
		*ErrorString = SysAllocString(wszString);
		CoTaskMemFree(wszString);

	   	TRACE_MISC("%s -> from OPC Server: [0x%08X] '%s'", TRACE_FUNCTION, ErrorCode, OLE2T(*ErrorString));
		TRACE_FUNCTION_LEAVE();
		return S_OK;
	}

	// try to resolve the error code with any local ID (default language)
	hResult = ::GetErrorString(ErrorCode, m_localeID, true, &tsBuffer);

	if (SUCCEEDED(hResult))
	{
		*ErrorString = SysAllocString(T2OLE(tsBuffer));
		CoTaskMemFree(tsBuffer);

	   	TRACE_MISC("%s -> Message: [0x%08X] '%s'", TRACE_FUNCTION, ErrorCode, OLE2T(*ErrorString));
		TRACE_FUNCTION_LEAVE();
		return S_OK;
	}

	// should never happen.
	TRACE_MISC("%s -> Couldn't resolve ErrorCode:[0x%08X]", TRACE_FUNCTION, ErrorCode);
    return E_FAIL;
}

//==============================================================================
// QueryAvailableLocaleIDs Method

STDMETHODIMP COPCAutoServer::QueryAvailableLocaleIDs(VARIANT* LocaleIDs)
{
	TRACE_FUNCTION_ENTER("COPCAutoServer::QueryAvailableLocaleIDs");
	
	// check arguments.
	if (!CheckByRefOutArg(LocaleIDs)) return E_INVALIDARG;

	// initialize return parameters.
	VariantInit(LocaleIDs);

	// check if server supports IOPCCommon.
	if (!m_pOPCCommon)
	{
		TRACE_INTERFACE_ERROR(IOPCCommon);
		return E_FAIL;
	}

	// fetch locales from the server.
	DWORD dwCount  = 0;
	LCID* pLocales = NULL;

	TRACE_METHOD_CALL(QueryAvailableLocaleIDs);

	HRESULT hResult = m_pOPCCommon->QueryAvailableLocaleIDs(&dwCount, &pLocales);
	
	if (FAILED(hResult))
	{
		TRACE_METHOD_ERROR(QueryAvailableLocaleIDs, hResult);
		return hResult;
	}

	// handle servers that return pLocales == NULL and S_OK
	if (pLocales == NULL)
	{
		TRACE_METHOD_ERROR(QueryAvailableLocaleIDs, E_POINTER);
		return E_POINTER;
	}

	// copy locales to a safearray.
	SAFEARRAYBOUND cBounds;
	cBounds.lLbound   = 1;
	cBounds.cElements = dwCount;

	SAFEARRAY* pArray = SafeArrayCreate(VT_I4, 1, &cBounds);
	
	if (pArray != NULL)
	{	
		for (DWORD ii = 1; ii <= dwCount; ii++ )
		{
			SafeArrayPutElement(pArray, (LONG*)&ii, &pLocales[ii-1]);
		}

		LocaleIDs->vt     = VT_ARRAY | VT_I4;
		LocaleIDs->parray = pArray;
	}

	CoTaskMemFree(pLocales);

	TraceByRefOutArg(LocaleIDs);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// QueryAvailableProperties Method

STDMETHODIMP COPCAutoServer::QueryAvailableProperties(
	BSTR        ItemID,
	LONG*       pCount,
	SAFEARRAY** ppPropertyIDs,
	SAFEARRAY** ppDescriptions,
	SAFEARRAY** ppDataTypes
)
{	
	USES_CONVERSION;

	TRACE_FUNCTION_ENTER("COPCAutoServer::QueryAvailableProperties");
	
	// check arguments.
	if (!CheckByValInArg(ItemID))          return E_INVALIDARG;
	if (!CheckByRefOutArg(pCount))         return E_INVALIDARG;
	if (!CheckByRefOutArg(ppPropertyIDs))  return E_INVALIDARG;
	if (!CheckByRefOutArg(ppDescriptions)) return E_INVALIDARG;
	if (!CheckByRefOutArg(ppDataTypes))    return E_INVALIDARG;

	// initialize return parameters.
	*pCount         = 0;
	*ppPropertyIDs  = NULL;
	*ppDescriptions = NULL;
	*ppDataTypes    = NULL;

	// check if server supports IOPCItemProperties.
	if (!m_pOPCItemProperties)
	{
		TRACE_INTERFACE_ERROR(IOPCItemProperties);
		return E_FAIL;
	}

	// fetch properties from the server.
	DWORD*   pPropertyIDs  = NULL;
	LPWSTR*  pDescriptions = NULL;
	VARTYPE* pDataTypes    = NULL;

	TRACE_METHOD_CALL(QueryAvailableProperties);

	HRESULT hResult = m_pOPCItemProperties->QueryAvailableProperties(
		ItemID,
		(DWORD*)pCount,
		&pPropertyIDs,
		&pDescriptions,
		&pDataTypes
	);

	if (FAILED(hResult))
	{
		TRACE_METHOD_ERROR(QueryAvailableProperties, hResult);
		return hResult;
	}

	// check for invalid response from server.
	if (pPropertyIDs == NULL || pDescriptions == NULL || pDataTypes == NULL)
	{
		TRACE_METHOD_ERROR(QueryAvailableProperties, E_POINTER);
		return E_POINTER;
	}

	// convert returned arrays to SAFEARRAYs.
	SAFEARRAYBOUND cBounds;
	cBounds.lLbound   = 1;
	cBounds.cElements = *pCount;

	*ppPropertyIDs  = SafeArrayCreate(VT_I4, 1, &cBounds);
	*ppDescriptions = SafeArrayCreate(VT_BSTR, 1, &cBounds);
	*ppDataTypes    = SafeArrayCreate(VT_I2, 1, &cBounds);

	for (LONG ii = 1; ii <= *pCount; ii++)
	{
		SafeArrayPutElement(*ppPropertyIDs, (LONG*)&ii, &pPropertyIDs[ii-1]);

		BSTR szDescription = SysAllocString(pDescriptions[ii-1]);
		SafeArrayPutElement(*ppDescriptions, (LONG*)&ii, szDescription);
		SysFreeString(szDescription);
		CoTaskMemFree(pDescriptions[ii-1]);

		SafeArrayPutElement(*ppDataTypes, (LONG*)&ii, &pDataTypes[ii-1]);
	}

	// free memory returned from the server.
	CoTaskMemFree(pPropertyIDs);
	CoTaskMemFree(pDescriptions);
	CoTaskMemFree(pDataTypes);

	// write results to a log.
	TraceByRefOutArg(pCount);
	TraceByRefOutArg(ppPropertyIDs);
	TraceByRefOutArg(ppDescriptions);
	TraceByRefOutArg(ppDataTypes);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// GetItemProperties Method

STDMETHODIMP COPCAutoServer::GetItemProperties(
	BSTR        ItemID,
	LONG        Count,
	SAFEARRAY** ppPropertyIDs,
	SAFEARRAY** ppPropertyValues,
	SAFEARRAY** ppErrors
)
{
	USES_CONVERSION;

	TRACE_FUNCTION_ENTER("COPCAutoServer::GetItemProperties");

	// check arguments.
	if (!CheckByValInArg(ItemID))                            return E_INVALIDARG;
	if (!CheckByValInArg(Count))                             return E_INVALIDARG;
	if (!CheckByRefInArrayArg(ppPropertyIDs, VT_I4, Count))  return E_INVALIDARG;
	if (!CheckByRefOutArg(ppPropertyValues))                 return E_INVALIDARG;
	if (!CheckByRefOutArg(ppErrors))                         return E_INVALIDARG;

	// initialize return parameters.
	*ppPropertyValues = NULL;
	*ppErrors         = NULL;

	// check if server supports IOPCItemProperties.
	if (!m_pOPCItemProperties)
	{
		TRACE_INTERFACE_ERROR(IOPCItemProperties);
		return E_FAIL;
	}

	// copy property ids into an array.
	DWORD* pPropertyIDs = (DWORD*)malloc(Count*sizeof(DWORD));

	for (LONG ii = 1; ii <= Count; ii++)
	{		
		SafeArrayGetElement(*ppPropertyIDs, &ii, &pPropertyIDs[ii-1]);
	}

	// get property values from server.
	VARIANT* pValues = NULL;
	HRESULT* pErrors = NULL;

	TRACE_METHOD_CALL(GetItemProperties);

	HRESULT hResult = m_pOPCItemProperties->GetItemProperties(
		ItemID,
		(DWORD)Count,
		pPropertyIDs,
		&pValues,
		&pErrors
	);

	free(pPropertyIDs);

	if (FAILED(hResult))
	{
		TRACE_METHOD_ERROR(GetItemProperties, hResult);
		return hResult;
	}

	// check for invalid response from server.
	if (pValues == NULL || pErrors == NULL)
	{
		TRACE_METHOD_ERROR(GetItemProperties, E_POINTER);
		return E_POINTER;
	}

	// convert returned arrays to SAFEARRAYs.
	SAFEARRAYBOUND cBounds;
	cBounds.lLbound   = 1;
	cBounds.cElements = Count;

	*ppPropertyValues = SafeArrayCreate(VT_VARIANT, 1, &cBounds);
	*ppErrors         = SafeArrayCreate(VT_I4, 1, &cBounds);

	for (ii = 1; ii <= Count; ii++)
	{
		VariantToAutomation(&pValues[ii-1]);
		SafeArrayPutElement(*ppPropertyValues, (LONG*)&ii, &pValues[ii-1]);
		VariantClear(&pValues[ii-1]);

		SafeArrayPutElement(*ppErrors, (LONG*)&ii, &pErrors[ii-1]);
	}

	// free memory returned from the server.
	CoTaskMemFree(pValues);
	CoTaskMemFree(pErrors);

	// write results to a log.
	TraceByRefOutArg(ppPropertyValues);
	TraceByRefOutArg(ppErrors);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// LookupItemIDs Method

STDMETHODIMP COPCAutoServer::LookupItemIDs(
	BSTR        ItemID,
	LONG        Count,
	SAFEARRAY** ppPropertyIDs,
	SAFEARRAY** ppNewItemIDs,
	SAFEARRAY** ppErrors
)
{
	USES_CONVERSION;

	TRACE_FUNCTION_ENTER("COPCAutoServer::LookupItemIDs");

	// check arguments.
	if (!CheckByValInArg(ItemID))                            return E_INVALIDARG;
	if (!CheckByValInArg(Count))                             return E_INVALIDARG;
	if (!CheckByRefInArrayArg(ppPropertyIDs, VT_I4, Count))  return E_INVALIDARG;
	if (!CheckByRefOutArg(ppNewItemIDs))                     return E_INVALIDARG;
	if (!CheckByRefOutArg(ppErrors))                         return E_INVALIDARG;

	// initialize return parameters.
	*ppNewItemIDs = NULL;
	*ppErrors     = NULL;

	// check if server supports IOPCItemProperties.
	if (!m_pOPCItemProperties)
	{
		TRACE_INTERFACE_ERROR(IOPCItemProperties);
		return E_FAIL;
	}

	// copy property ids into an array.
	DWORD* pPropertyIDs = (DWORD*)malloc(Count*sizeof(DWORD));

	for (LONG ii = 1; ii <= Count; ii++)
	{		
		SafeArrayGetElement(*ppPropertyIDs, &ii, &pPropertyIDs[ii-1]);
	}

	// get property values from server.
	LPWSTR*  pNewItemIDs = NULL;
	HRESULT* pErrors     = NULL;

	TRACE_METHOD_CALL(LookupItemIDs);

	HRESULT hResult = m_pOPCItemProperties->LookupItemIDs(
		ItemID,
		(DWORD)Count,
		pPropertyIDs,
		&pNewItemIDs,
		&pErrors
	);

	free(pPropertyIDs);

	if (FAILED(hResult))
	{
		TRACE_METHOD_ERROR(GetItemProperties, hResult);
		return hResult;
	}

	// check for invalid response from server.
	if (pNewItemIDs == NULL || pErrors == NULL)
	{
		TRACE_METHOD_ERROR(GetItemProperties, E_POINTER);
		return E_POINTER;
	}

	// convert returned arrays to SAFEARRAYs.
	SAFEARRAYBOUND cBounds;
	cBounds.lLbound   = 1;
	cBounds.cElements = Count;

	*ppNewItemIDs = SafeArrayCreate(VT_BSTR, 1, &cBounds);
	*ppErrors     = SafeArrayCreate(VT_I4, 1, &cBounds);

	for (ii = 1; ii <= Count; ii++)
	{
		BSTR szItemID = SysAllocString(pNewItemIDs[ii-1]);
	
		SafeArrayPutElement(*ppNewItemIDs, (LONG*)&ii, szItemID);
		SafeArrayPutElement(*ppErrors, (LONG*)&ii, &pErrors[ii-1]);

		SysFreeString(szItemID);
		CoTaskMemFree(pNewItemIDs[ii-1]);
	}

	// free memory returned from the server.
	CoTaskMemFree(pNewItemIDs);
	CoTaskMemFree(pErrors);

	// write results to a log.
	TraceByRefOutArg(pNewItemIDs);
	TraceByRefOutArg(ppErrors);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}
