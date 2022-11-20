//==============================================================================
// TITLE: OPCGroupEvent.h
//
// CONTENTS:
//
// Defines the events exposed by the DA Automation interface objects.
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

#ifndef __OPCGroupEvent_H__      
#define __OPCGroupEvent_H__

//==============================================================================
// CLASS:   CProxyDIOPCGroupEvent
// PURPOSE: Implements methods defined for the DIOPCGroupEvent event.

template <class T>
class CProxyDIOPCGroupEvent : public IConnectionPointImpl<T, &DIID_DIOPCGroupEvent, CComDynamicUnkArray>
{
public:

	//==========================================================================
	// Fire_DataChange

	void Fire_DataChange(
		long       TransactionID,
		long       NumItems,
		SAFEARRAY* ClientHandles,
		SAFEARRAY* ItemValues,
		SAFEARRAY* Qualities,
		SAFEARRAY* TimeStamps
	)
	{
		VARIANTARG* pvars = new VARIANTARG[6];
		for (int i = 0; i < 6; i++) VariantInit(&pvars[i]);
		
		// acquire synchronization lock on object.
		T* pT = (T*)this;
		pT->Lock();
		
		IUnknown** pp = m_vec.begin();
		
		while (pp < m_vec.end())
		{
			if (*pp != NULL)
			{
				pvars[5].vt     = VT_I4;
				pvars[5].lVal   = TransactionID;
				pvars[4].vt     = VT_I4;
				pvars[4].lVal   = NumItems;
				pvars[3].vt     = VT_I4 | VT_ARRAY;
				pvars[3].parray = ClientHandles;
				pvars[2].vt     = VT_VARIANT | VT_ARRAY;
				pvars[2].parray = ItemValues;
				pvars[1].vt     = VT_I4 | VT_ARRAY;
				pvars[1].parray = Qualities;
				pvars[0].vt     = VT_DATE | VT_ARRAY;
				pvars[0].parray = TimeStamps;

				DISPPARAMS disp = { pvars, NULL, 6, 0 };
				IDispatch* pDispatch = reinterpret_cast<IDispatch*>(*pp);
				pDispatch->Invoke(0x1, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &disp, NULL, NULL, NULL);
			}

			pp++;
		}

		// release synchronization lock on object.
		pT->Unlock();
		
		delete[] pvars;
	}
	
	//==========================================================================
	// Fire_AsyncReadComplete

	void Fire_AsyncReadComplete(
		long       TransactionID,
		long       NumItems,
		SAFEARRAY* ClientHandles,
		SAFEARRAY* ItemValues,
		SAFEARRAY* Qualities,
		SAFEARRAY* TimeStamps,
		SAFEARRAY* Errors
	)
	{
		VARIANTARG* pvars = new VARIANTARG[7];
		for (int i = 0; i < 7; i++)	VariantInit(&pvars[i]);
		
		// acquire synchronization lock on object.
		T* pT = (T*)this;
		pT->Lock();
		
		IUnknown** pp = m_vec.begin();
		
		while (pp < m_vec.end())
		{
			if (*pp != NULL)
			{
				pvars[6].vt     = VT_I4;
				pvars[6].lVal   = TransactionID;
				pvars[5].vt     = VT_I4;
				pvars[5].lVal   = NumItems;
				pvars[4].vt     = VT_I4 | VT_ARRAY;
				pvars[4].parray = ClientHandles;
				pvars[3].vt     = VT_VARIANT | VT_ARRAY;
				pvars[3].parray = ItemValues;
				pvars[2].vt     = VT_I4 | VT_ARRAY;
				pvars[2].parray = Qualities;
				pvars[1].vt     = VT_DATE | VT_ARRAY;
				pvars[1].parray = TimeStamps;
				pvars[0].vt     = VT_I4 | VT_ARRAY;
				pvars[0].parray = Errors;

				DISPPARAMS disp = { pvars, NULL, 7, 0 };
				IDispatch* pDispatch = reinterpret_cast<IDispatch*>(*pp);
				pDispatch->Invoke(0x2, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &disp, NULL, NULL, NULL);
			}

			pp++;
		}

		// release synchronization lock on object.
		pT->Unlock();
		
		delete[] pvars;
	}
	
	//==========================================================================
	// Fire_AsyncWriteComplete

	void Fire_AsyncWriteComplete(
		long       TransactionID,
		long       NumItems,
		SAFEARRAY* ClientHandles,
		SAFEARRAY* Errors
	)
	{
		VARIANTARG* pvars = new VARIANTARG[4];
		for (int i = 0; i < 4; i++)	VariantInit(&pvars[i]);
		
		// acquire synchronization lock on object.
		T* pT = (T*)this;
		pT->Lock();
		
		IUnknown** pp = m_vec.begin();
		
		while (pp < m_vec.end())
		{
			if (*pp != NULL)
			{
				pvars[3].vt     = VT_I4;
				pvars[3].lVal   = TransactionID;
				pvars[2].vt     = VT_I4;
				pvars[2].lVal   = NumItems;
				pvars[1].vt     = VT_I4 | VT_ARRAY;
				pvars[1].parray = ClientHandles;
				pvars[0].vt     = VT_I4 | VT_ARRAY;
				pvars[0].parray = Errors;

				DISPPARAMS disp = { pvars, NULL, 4, 0 };
				IDispatch* pDispatch = reinterpret_cast<IDispatch*>(*pp);
				pDispatch->Invoke(0x3, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &disp, NULL, NULL, NULL);
			}
			pp++;
		}
		
		// release synchronization lock on object.
		pT->Unlock();
		
		delete[] pvars;
	}
	
	//==========================================================================
	// Fire_AsyncCancelComplete

	void Fire_AsyncCancelComplete(long TransactionID)
	{
		VARIANTARG* pvars = new VARIANTARG[1];
		for (int i = 0; i < 1; i++)	VariantInit(&pvars[i]);
		
		// acquire synchronization lock on object.
		T* pT = (T*)this;
		pT->Lock();

		IUnknown** pp = m_vec.begin();
		while (pp < m_vec.end())
		{
			if (*pp != NULL)
			{
				pvars[0].vt   = VT_I4;
				pvars[0].lVal = TransactionID;

				DISPPARAMS disp = { pvars, NULL, 1, 0 };
				IDispatch* pDispatch = reinterpret_cast<IDispatch*>(*pp);
				pDispatch->Invoke(0x4, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &disp, NULL, NULL, NULL);
			}

			pp++;
		}

		// release synchronization lock on object.
		pT->Unlock();
		
		delete[] pvars;
	}

};

//==============================================================================
// CLASS:   CProxyDIOPCServerEvent
// PURPOSE: Implements methods defined for the DIOPCServerEvent event.

template <class T>
class CProxyDIOPCServerEvent : public IConnectionPointImpl<T, &DIID_DIOPCServerEvent, CComDynamicUnkArray>
{
public:

	//==========================================================================
	// Fire_AsyncCancelComplete

	void Fire_ServerShutDown( BSTR Reason)
	{
		VARIANTARG* pvars = new VARIANTARG[1];
		for (int i = 0; i < 1; i++)	VariantInit(&pvars[i]);
		
		// acquire synchronization lock on object.
		T* pT = (T*)this;
		pT->Lock();

		IUnknown** pp = m_vec.begin();
		
		while (pp < m_vec.end())
		{
			if (*pp != NULL)
			{
				pvars[0].vt      = VT_BSTR;
				pvars[0].bstrVal = Reason;

				DISPPARAMS disp = { pvars, NULL, 1, 0 };
				IDispatch* pDispatch = reinterpret_cast<IDispatch*>(*pp);
				pDispatch->Invoke(0x1, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &disp, NULL, NULL, NULL);
			}

			pp++;
		}
		
		// release synchronization lock on object.
		pT->Unlock();

		delete[] pvars;
	}
};

//==============================================================================
// CLASS:   CProxyDIOPCGroupsEvent
// PURPOSE: Implements methods defined for DIOPCGroupsEvent event.

template <class T>
class CProxyDIOPCGroupsEvent : public IConnectionPointImpl<T, &DIID_DIOPCGroupsEvent, CComDynamicUnkArray>
{
public:
 
	//==========================================================================
	// Fire_GlobalDataChange

	void Fire_GlobalDataChange(
		long       TransactionID,
		long       GroupHandle,
		long       NumItems,
		SAFEARRAY* ClientHandles,
		SAFEARRAY* ItemValues,
		SAFEARRAY* Qualities,
		SAFEARRAY* TimeStamps
	)
	{
		VARIANTARG* pvars = new VARIANTARG[7];
		for (int i = 0; i < 7; i++) VariantInit(&pvars[i]);
		
		// acquire synchronization lock on object.
		T* pT = (T*)this;
		pT->Lock();
		
		IUnknown** pp = m_vec.begin();
		
		while (pp < m_vec.end())
		{
			if (*pp != NULL)
			{
				pvars[6].vt     = VT_I4;
				pvars[6].lVal   = TransactionID;
				pvars[5].vt     = VT_I4;
				pvars[5].lVal   = GroupHandle;
				pvars[4].vt     = VT_I4;
				pvars[4].lVal   = NumItems;
				pvars[3].vt     = VT_I4 | VT_ARRAY;
				pvars[3].parray = ClientHandles;
				pvars[2].vt     = VT_VARIANT | VT_ARRAY;
				pvars[2].parray = ItemValues;
				pvars[1].vt     = VT_I4 | VT_ARRAY;
				pvars[1].parray = Qualities;
				pvars[0].vt     = VT_DATE | VT_ARRAY;
				pvars[0].parray = TimeStamps;

				DISPPARAMS disp = { pvars, NULL, 7, 0 };
				IDispatch* pDispatch = reinterpret_cast<IDispatch*>(*pp);
				pDispatch->Invoke(0x1, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &disp, NULL, NULL, NULL);
			}

			pp++;
		}
		
		// release synchronization lock on object.
		pT->Unlock();
		
		delete[] pvars;
	}
};

#endif //__OPCGroupEvent_H__   