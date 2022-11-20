//==============================================================================
// TITLE: StdAfx.h
//
// CONTENTS:
//
//  Precompiled header file.
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

#if !defined(AFX_STDAFX_H__09D080DA_F1B3_11D0_8D0A_000000000000__INCLUDED_)
#define AFX_STDAFX_H__09D080DA_F1B3_11D0_8D0A_000000000000__INCLUDED_

#if _MSC_VER >= 1000
#pragma once
#endif // _MSC_VER >= 1000

#define STRICT
#pragma warning( disable : 4786 )  // "identifier truncated" warning

#ifndef _ATL_STATIC_REGISTRY
#define _ATL_STATIC_REGISTRY
#endif

#define _WIN32_WINNT 0x0400
#define _ATL_DEBUG_QI

#include <atlbase.h>
//You may derive a class from CComModule and use it if you want to override
//something, but do not change the name of _Module
extern CComModule _Module;
#include <atlcom.h>

#ifdef _ATL_STATIC_REGISTRY
#include <statreg.h>
#endif

#include <List>
using namespace std; // for STL

#include <errno.h>

#include "resource.h"

#include "OpcComn.h"
#include "OpcDa.h"
#include "OpcEnum.h"
#include "opcerror.h"

#include "OPCAuto.h"

typedef CComQIPtr<IOPCServer, &IID_IOPCServer> IOPCServerPtr;
typedef CComQIPtr<IOPCItemProperties, &IID_IOPCItemProperties> IOPCItemPropertiesPtr;
typedef CComQIPtr<IOPCCommon, &IID_IOPCCommon> IOPCCommonPtr;
typedef CComQIPtr<IOPCBrowseServerAddressSpace, &IID_IOPCBrowseServerAddressSpace> IOPCBrowseServerAddressSpacePtr;
typedef CComQIPtr<IOPCGroupStateMgt, &IID_IOPCGroupStateMgt> IOPCGroupStateMgtPtr;
typedef CComQIPtr<IOPCSyncIO, &IID_IOPCSyncIO> IOPCSyncIOPtr;
typedef CComQIPtr<IOPCAsyncIO, &IID_IOPCAsyncIO> IOPCAsyncIOPtr;
typedef CComQIPtr<IOPCAsyncIO2, &IID_IOPCAsyncIO2> IOPCAsyncIO2Ptr;
typedef CComQIPtr<IOPCItemMgt, &IID_IOPCItemMgt> IOPCItemMgtPtr;
typedef CComQIPtr<IDataObject, &IID_IDataObject> IDataObjectPtr;
typedef CComQIPtr<IOPCGroupStateMgt, &IID_IOPCGroupStateMgt> IOPCGroupStateMgtPtr;

//{{AFX_INSERT_LOCATION}}
// Microsoft Developer Studio will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_STDAFX_H__09D080DA_F1B3_11D0_8D0A_000000000000__INCLUDED)
