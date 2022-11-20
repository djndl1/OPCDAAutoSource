//==============================================================================
// TITLE: OPCGroupCallback.cpp
//
// CONTENTS:
//
//  These are the Group Object's data advise callback functions.
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
// 2003/10/27 RSA   Merged bug fixes received from members.
//

#include "StdAfx.h"
#include "OPCGroups.h"
#include "OPCUtils.h"
#include "OPCTrace.h"

extern UINT OPCSTMFORMATDATATIME;
extern UINT OPCSTMFORMATWRITECOMPLETE;

//==============================================================================
// IOPCDataCallback::OnDataChange Method

STDMETHODIMP COPCGroup::OnDataChange(
    DWORD      dwTransid, 
    OPCHANDLE  hGroup, 
    HRESULT    hrMasterquality,
    HRESULT    hrMastererror,
    DWORD      dwCount, 
    OPCHANDLE* phClientItems, 
    VARIANT*   pvValues, 
    WORD*      pwQualities,
    FILETIME*  pftTimeStamps,
    HRESULT*   pErrors
)
{
	TRACE_FUNCTION_ENTER("IOPCDataCallback::OnDataChange");

	// validate group handle.
	if (hGroup != (OPCHANDLE)this)
	{		
		TRACE_ERROR("%s 'Callback with invalid Group Handle=0x%08X'", TRACE_FUNCTION, hGroup);
		return E_INVALIDARG;
	}

	// check for data.
	if (dwCount == 0)
	{
		TRACE_OUT("%s - No data received", TRACE_FUNCTION);
		return S_OK;
	}

	// check for invalid pointers.
	if (phClientItems == NULL || pvValues == NULL || pwQualities == NULL || pftTimeStamps == NULL || pErrors == NULL)
	{	
		TRACE_ERROR("%s 'Callback passed a NULL pointer.'", TRACE_FUNCTION);	
		return E_POINTER;
	}

	// create a list of valid items.
	list<COPCItem*> items;

	for (LONG ii = 0; ii < (LONG)dwCount; ii++)
	{
		COPCItem* pItem = LookupItem(phClientItems[ii]);
			
		if (pItem != NULL)
		{
			pItem->Update(&(pvValues[ii]), pwQualities[ii], pftTimeStamps[ii]);
			items.push_back(pItem);
		}
	}

	// print an error if invalid items encountered.
	if (items.size() != dwCount)
	{
		TRACE_ERROR("%s 'Callback passed %d invalid item handles.'", TRACE_FUNCTION, dwCount-items.size());	

		if (items.size() == 0)
		{
			return S_OK;
		}
	}

	// create arrays for callback.
	SAFEARRAYBOUND cBounds;
	cBounds.lLbound   = 1;
	cBounds.cElements = items.size();

	SAFEARRAY* psaHandles    = SafeArrayCreate(VT_I4, 1, &cBounds);
	SAFEARRAY* psaValues     = SafeArrayCreate(VT_VARIANT, 1, &cBounds);
	SAFEARRAY* psaQualities  = SafeArrayCreate(VT_I4, 1, &cBounds);
	SAFEARRAY* psaTimestamps = SafeArrayCreate(VT_DATE, 1, &cBounds);

	list<COPCItem*>::iterator iter = items.begin();

	for (ii = 1; ii <= (LONG)items.size() && iter != items.end(); ii++)
	{		
		COPCItem* pItem = *iter;

		// add client handle.
		OPCHANDLE hClient = pItem->GetClientHandle();
		SafeArrayPutElement(psaHandles, &ii, &hClient);
		
		// add value.
		const VARIANT* pValue = pItem->GetValue();
		SafeArrayPutElement(psaValues, &ii, (void*)pValue);

		// add quality.
		LONG lQuality = pItem->GetQuality();
		SafeArrayPutElement(psaQualities, &ii, &lQuality);
		
		// add timestamp.
		DATE dateTimestamp = GetUtcTime(pItem->GetTimestamp());
		SafeArrayPutElement(psaTimestamps, &ii, &dateTimestamp);

		iter++;
	}

	TraceByRefOutArg(&psaHandles);
	TraceByRefOutArg(&psaValues);
	TraceByRefOutArg(&psaQualities);
	TraceByRefOutArg(&psaTimestamps);

	// send the group data callback.
	Fire_DataChange(dwTransid, dwCount, psaHandles, psaValues, psaQualities, psaTimestamps);

	// send the global data callback.
	m_pParent->Fire_GlobalDataChange(dwTransid, m_client, dwCount, psaHandles, psaValues, psaQualities, psaTimestamps);

	SafeArrayDestroy(psaHandles);
	SafeArrayDestroy(psaValues);
	SafeArrayDestroy(psaQualities);
	SafeArrayDestroy(psaTimestamps);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// IOPCDataCallback::OnReadComplete Method

STDMETHODIMP COPCGroup::OnReadComplete(	
	DWORD      dwTransid, 
	OPCHANDLE  hGroup, 
	HRESULT    hrMasterquality,
	HRESULT    hrMastererror,
	DWORD      dwCount, 
	OPCHANDLE* phClientItems, 
	VARIANT*   pvValues, 
	WORD*      pwQualities,
	FILETIME*  pftTimeStamps,
	HRESULT*   pErrors
)
{
	TRACE_FUNCTION_ENTER("IOPCDataCallback::OnReadComplete");

	// validate group handle.
	if (hGroup != (OPCHANDLE)this)
	{		
		TRACE_ERROR("%s 'Callback with invalid Group Handle=0x%08X'", TRACE_FUNCTION, hGroup);
		return E_INVALIDARG;
	}

	// check for data.
	if (dwCount == 0)
	{
		TRACE_OUT("%s - No data received", TRACE_FUNCTION);
		return S_OK;
	}

	// check for invalid pointers.
	if (phClientItems == NULL || pvValues == NULL || pwQualities == NULL || pftTimeStamps == NULL || pErrors == NULL)
	{	
		TRACE_ERROR("%s 'Callback passed a NULL pointer.'", TRACE_FUNCTION);	
		return E_POINTER;
	}

	// create arrays for callback.
	SAFEARRAYBOUND cBounds;
	cBounds.lLbound   = 1;
	cBounds.cElements = dwCount;

	SAFEARRAY* psaHandles    = SafeArrayCreate(VT_I4, 1, &cBounds);
	SAFEARRAY* psaValues     = SafeArrayCreate(VT_VARIANT, 1, &cBounds);
	SAFEARRAY* psaQualities  = SafeArrayCreate(VT_I4, 1, &cBounds);
	SAFEARRAY* psaTimestamps = SafeArrayCreate(VT_DATE, 1, &cBounds);
	SAFEARRAY* psaErrors     = SafeArrayCreate(VT_I4, 1, &cBounds);

	for (LONG ii = 1; ii <= (LONG)dwCount; ii++)
	{		
		COPCItem* pItem = LookupItem(phClientItems[ii-1]);

		OPCHANDLE hClient       = NULL;
		LONG      lQuality      = OPC_QUALITY_BAD;
		DATE      dateTimestamp = 0;
		HRESULT   hError        = OPC_E_INVALIDHANDLE;

		VARIANT varValue;    
		VariantInit(&varValue);
		
		if (pItem != NULL)
		{
			hClient = pItem->GetClientHandle();
			SafeArrayPutElement(psaHandles, &ii, &hClient);

			if (SUCCEEDED(pErrors[ii-1]))
			{
				pItem->Update(&(pvValues[ii-1]), pwQualities[ii-1], pftTimeStamps[ii-1]);
				
				const VARIANT* pValue = pItem->GetValue();
				SafeArrayPutElement(psaValues, &ii, (void*)pValue);

				lQuality = pItem->GetQuality();
				SafeArrayPutElement(psaQualities, &ii, &lQuality);
				
				dateTimestamp = GetUtcTime(pItem->GetTimestamp());
				SafeArrayPutElement(psaTimestamps, &ii, &dateTimestamp);
			}
			else
			{
				SafeArrayPutElement(psaValues, &ii, &varValue);
				SafeArrayPutElement(psaQualities, &ii, &lQuality);
				SafeArrayPutElement(psaTimestamps, &ii, &dateTimestamp);
			}

			hError = pErrors[ii-1];
		}
		else
		{
			SafeArrayPutElement(psaHandles, &ii, &hClient);
			SafeArrayPutElement(psaValues, &ii, &varValue);
			SafeArrayPutElement(psaQualities, &ii, &lQuality);			
			SafeArrayPutElement(psaTimestamps, &ii, &dateTimestamp);
		}
		
		SafeArrayPutElement(psaErrors, &ii, &hError);
	}
	
	TraceByRefOutArg(&psaHandles);
	TraceByRefOutArg(&psaValues);
	TraceByRefOutArg(&psaQualities);
	TraceByRefOutArg(&psaTimestamps);
	TraceByRefOutArg(&psaErrors);

	// send the group read complete callback.
	Fire_AsyncReadComplete(dwTransid, dwCount, psaHandles, psaValues, psaQualities, psaTimestamps, psaErrors);

	SafeArrayDestroy(psaHandles);
	SafeArrayDestroy(psaValues);
	SafeArrayDestroy(psaQualities);
	SafeArrayDestroy(psaTimestamps);
	SafeArrayDestroy(psaErrors);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// IOPCDataCallback::OnWriteComplete Method

STDMETHODIMP COPCGroup::OnWriteComplete(
	DWORD      dwTransid, 
	OPCHANDLE  hGroup, 
	HRESULT    hrMastererr, 
	DWORD      dwCount, 
	OPCHANDLE* pClientHandles, 
	HRESULT*   pErrors
)
{
	TRACE_FUNCTION_ENTER("IOPCDataCallback::OnWriteComplete");

	// validate group handle.
	if (hGroup != (OPCHANDLE)this)
	{		
		TRACE_ERROR("%s 'Callback with invalid Group Handle=0x%08X'", TRACE_FUNCTION, hGroup);
		return E_INVALIDARG;
	}

	// check for data.
	if (dwCount == 0)
	{
		TRACE_OUT("%s - No data received", TRACE_FUNCTION);
		return S_OK;
	}

	// check for invalid pointers.
	if (pClientHandles == NULL || pErrors == NULL)
	{	
		TRACE_ERROR("%s 'Callback passed a NULL pointer.'", TRACE_FUNCTION);	
		return E_POINTER;
	}

	// create arrays for callback.
	SAFEARRAYBOUND cBounds;
	cBounds.lLbound   = 1;
	cBounds.cElements = dwCount;

	SAFEARRAY* psaHandles = SafeArrayCreate(VT_I4, 1, &cBounds);
	SAFEARRAY* psaErrors  = SafeArrayCreate(VT_I4, 1, &cBounds);

	for (LONG ii = 1; ii <= (LONG)dwCount; ii++)
	{		
		COPCItem* pItem = LookupItem(pClientHandles[ii-1]);

		if (pItem != NULL && SUCCEEDED(pErrors[ii-1]))
		{
			OPCHANDLE hClient = pItem->GetClientHandle();
			SafeArrayPutElement(psaHandles, &ii, &hClient);

			SafeArrayPutElement(psaErrors, &ii, &(pErrors[ii-1]));
		}
		else
		{
			OPCHANDLE hClient = NULL;
			SafeArrayPutElement(psaHandles, &ii, &hClient);
			
			HRESULT hError = (pItem != NULL)?pErrors[ii-1]:OPC_E_INVALIDHANDLE;
			SafeArrayPutElement(psaErrors, &ii, &hError);
		}
	}
	
	// write to log file.
	TraceByRefOutArg(&psaHandles);
	TraceByRefOutArg(&psaErrors);

	// send the group write complete callback.
	Fire_AsyncWriteComplete(dwTransid, dwCount, psaHandles, psaErrors);

	SafeArrayDestroy(psaHandles);
	SafeArrayDestroy(psaErrors);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// IOPCDataCallback::OnCancelComplete Method

STDMETHODIMP COPCGroup::OnCancelComplete(DWORD dwTransid, OPCHANDLE hGroup)
{
	TRACE_FUNCTION_ENTER("IOPCDataCallback::OnCancelComplete");

	// validate group handle.
	if (hGroup != (OPCHANDLE)this)
	{		
		TRACE_ERROR("%s 'Callback with invalid Group Handle=0x%08X'", TRACE_FUNCTION, hGroup);
		return E_INVALIDARG;
	}
	
	// write to log file.
	TraceByRefOutArg(&dwTransid);

	// send the group write complete callback.
	Fire_AsyncCancelComplete(dwTransid);

	TRACE_FUNCTION_LEAVE();
	return S_OK;
}

//==============================================================================
// IAdviseSink::OnDataChange Method

STDMETHODIMP_(void) COPCGroup::OnDataChange(LPFORMATETC pFE, LPSTGMEDIUM pSTM)
{
	TRACE_FUNCTION_ENTER("IAdviseSink::OnDataChange");
	
	// check arguments.
	if (pFE == NULL || pSTM == NULL)
	{		
		TRACE_ERROR("%s, 'pFE == NULL || pSTM == NULL'", TRACE_FUNCTION);
		return;
	}

	// verify the format follows the OPC specification.
	if (TYMED_HGLOBAL != pFE->tymed)
	{
		TRACE_ERROR("%s, 'TYMED_HGLOBAL != pFE->tymed'", TRACE_FUNCTION);
		return;
	}

	if (pSTM->hGlobal == NULL)
	{
		TRACE_ERROR("%s, 'pSTM->hGlobal == NULL'", TRACE_FUNCTION);
		return;
	}

	// check for write complete notifications.
	if (OPCSTMFORMATDATATIME != pFE->cfFormat)
	{
		OnWriteComplete(pFE, pSTM);
		TRACE_FUNCTION_LEAVE();
		return;
	}

	// handle data notification.
	const BYTE* pBuffer = (BYTE*)GlobalLock(pSTM->hGlobal);
	
	if (pBuffer == NULL)
	{
		TRACE_ERROR("%s, 'pBuffer == NULL'", TRACE_FUNCTION);
		return;
	}

	const OPCGROUPHEADER* pHeader = (OPCGROUPHEADER*)pBuffer;

	// validate group handle.
	if (pHeader->hClientGroup != (OPCHANDLE)this)
	{
		GlobalUnlock(pSTM->hGlobal);
		TRACE_ERROR("%s 'Callback with invalid Group Handle=0x%08X'", TRACE_FUNCTION, pHeader->hClientGroup);
		return;
	}

	// check if no data received.
	if (pHeader->dwItemCount == 0)
	{
		GlobalUnlock(pSTM->hGlobal);
		TRACE_OUT("%s - No data received", TRACE_FUNCTION);
		return;
	}

	int iOffset = sizeof(OPCGROUPHEADER);

	// create a list of valid items.
	list<COPCItem*> items;

	for (LONG ii = 0; ii < (LONG)pHeader->dwItemCount; ii++, iOffset += sizeof(OPCITEMHEADER1))
	{
		const OPCITEMHEADER1* pItemHeader = (OPCITEMHEADER1*)&pBuffer[iOffset];

		VARIANT* pValue = (VARIANT*)&pBuffer[pItemHeader->dwValueOffset];
		
		// strings and arrays are packed in the stream requiring unpacking
		if( pValue->vt == VT_BSTR )
		{
			pValue->bstrVal = (BSTR)((BYTE*)pValue + sizeof(VARIANT) + sizeof(DWORD));
		}
		else if( (pValue->vt & VT_ARRAY) == VT_ARRAY )
		{
			pValue->parray = (SAFEARRAY*)((BYTE*)pValue + sizeof(VARIANT));
			pValue->parray->pvData = ((BYTE*)pValue->parray + sizeof(SAFEARRAY));
		}

		// lookup item from client handle.
		COPCItem* pItem = LookupItem(pItemHeader->hClient);
			
		if (pItem != NULL)
		{
			pItem->Update(pValue, (DWORD)pItemHeader->wQuality, pItemHeader->ftTimeStampItem);
		}

		// add to list of valid items.
		items.push_back(pItem);
	}

	DWORD dwItemCount     = pHeader->dwItemCount;
	DWORD dwTransactionID = pHeader->dwTransactionID;

	GlobalUnlock(pSTM->hGlobal);
	pHeader = NULL;

	// print an error if invalid items encountered.
	if (items.size() != dwItemCount)
	{
		TRACE_ERROR("%s 'Callback passed %d invalid item handles.'", TRACE_FUNCTION, pHeader->dwItemCount-items.size());	

		if (items.size() == 0)
		{
			return;
		}
	}

	// build SAFEARRAYs of handles, values, qualities, and timestamps
	SAFEARRAYBOUND cBounds;
	cBounds.lLbound = 1;
	cBounds.cElements = dwItemCount;
	
	SAFEARRAY* psaHandles    = SafeArrayCreate(VT_I4,      1, &cBounds);
	SAFEARRAY* psaValues     = SafeArrayCreate(VT_VARIANT, 1, &cBounds);
	SAFEARRAY* psaQualities  = SafeArrayCreate(VT_I4,      1, &cBounds);
	SAFEARRAY* psaTimestamps = SafeArrayCreate(VT_DATE,    1, &cBounds);
	SAFEARRAY* psaErrors     = SafeArrayCreate(VT_I4,      1, &cBounds);

	list<COPCItem*>::iterator iter = items.begin();

	for (ii = 1; ii <= (LONG)items.size() && iter != items.end(); ii++)
	{		
		COPCItem* pItem = *iter;

		OPCHANDLE      hClient       = NULL;
		const VARIANT* pValue        = NULL;
		LONG           lQuality      = 0;
		DATE           dateTimestamp = 0;
		HRESULT        hError        = OPC_E_INVALIDHANDLE;

		if (pItem != NULL)
		{
			hClient       = pItem->GetClientHandle();
			pValue        = pItem->GetValue();
			lQuality      = pItem->GetQuality();
			dateTimestamp = GetUtcTime(pItem->GetTimestamp());
			hError        = S_OK;
		}

		SafeArrayPutElement(psaHandles, &ii, &hClient);
		SafeArrayPutElement(psaValues, &ii, (void*)pValue);
		SafeArrayPutElement(psaQualities, &ii, &lQuality);
		SafeArrayPutElement(psaTimestamps, &ii, &dateTimestamp);
		SafeArrayPutElement(psaErrors, &ii, &hError);

		iter++;
	}

	bool bDataUpdate = true;

	// check for read complete notification.
	if (dwTransactionID != 0)
	{
		// check if the callback arrived before the AsyncRead returns.
		if  (m_asyncReading)   
		{
			m_readIDs.push_back(dwTransactionID);
			bDataUpdate = false;
		}

		// find transaction id in list of outstanding async reads.
		else
		{
			list<OPCHANDLE>::iterator iter = m_readIDs.begin();

			while (iter != m_readIDs.end())
			{
				if (*iter == dwTransactionID)
				{
					m_readIDs.remove(dwTransactionID);
					bDataUpdate = false;
					break;
				}

				iter++;
			}

			// treat as a data change notification if unrecognized transaction id.
		}
	}

	// write to log file.
	TraceByRefOutArg(&psaHandles);
	TraceByRefOutArg(&psaValues);
	TraceByRefOutArg(&psaQualities);
	TraceByRefOutArg(&psaTimestamps);
	TraceByRefOutArg(&psaErrors);

	// send data change notifications.
	if (bDataUpdate)
	{
		Fire_DataChange(dwTransactionID, items.size(), psaHandles, psaValues, psaQualities, psaTimestamps);

		// send the global group event notification.
		m_pParent->Fire_GlobalDataChange(dwTransactionID, m_client, items.size(), psaHandles, psaValues, psaQualities, psaTimestamps);
	}

	// send the on read complete notification.
	else
	{
		Fire_AsyncReadComplete(dwTransactionID, items.size(), psaHandles, psaValues, psaQualities, psaTimestamps, psaErrors);
	}

	SafeArrayDestroy(psaHandles);
	SafeArrayDestroy(psaValues);
	SafeArrayDestroy(psaQualities);
	SafeArrayDestroy(psaTimestamps);
	SafeArrayDestroy(psaErrors);

	TRACE_FUNCTION_LEAVE();
	return;
}

//==============================================================================
// COPCGroup::OnWriteComplete Method

STDMETHODIMP_(void) COPCGroup::OnWriteComplete(LPFORMATETC pFE, LPSTGMEDIUM pSTM)
{
	TRACE_FUNCTION_ENTER("COPCGroup::OnWriteComplete");
	
	// check arguments.
	if (pFE == NULL || pSTM == NULL)
	{		
		TRACE_ERROR("%s, 'pFE == NULL || pSTM == NULL'", TRACE_FUNCTION);
		return;
	}

	// verify the format follows the OPC specification.
	if (TYMED_HGLOBAL != pFE->tymed)
	{
		TRACE_ERROR("%s, 'TYMED_HGLOBAL != pFE->tymed'", TRACE_FUNCTION);
		return;
	}

	if (pSTM->hGlobal == NULL)
	{
		TRACE_ERROR("%s, 'pSTM->hGlobal == NULL'", TRACE_FUNCTION);
		return;
	}

	// check for write complete notifications.
	if (OPCSTMFORMATWRITECOMPLETE != pFE->cfFormat)
	{
		TRACE_ERROR("%s, 'OPCSTMFORMATWRITECOMPLETE != pFE->cfFormat'", TRACE_FUNCTION);
		return;
	}

	// handle data notification.
	const BYTE* pBuffer = (BYTE*)GlobalLock(pSTM->hGlobal);
	
	if (pBuffer == NULL)
	{
		TRACE_ERROR("%s, 'pBuffer == NULL'", TRACE_FUNCTION);
		return;
	}

	const OPCGROUPHEADERWRITE* pHeader = (OPCGROUPHEADERWRITE*)pBuffer;

	// validate group handle.
	if (pHeader->hClientGroup != (OPCHANDLE)this)
	{
		GlobalUnlock(pSTM->hGlobal);
		TRACE_ERROR("%s 'Callback with invalid Group Handle=0x%08X'", TRACE_FUNCTION, pHeader->hClientGroup);
		return;
	}

	// check if no data received.
	if (pHeader->dwItemCount == 0)
	{
		GlobalUnlock(pSTM->hGlobal);
		TRACE_OUT("%s - No data received", TRACE_FUNCTION);
		return;
	}

	// build SAFEARRAYs of handles and errors.
	SAFEARRAYBOUND cBounds;
	cBounds.lLbound = 1;
	cBounds.cElements = pHeader->dwItemCount;
	
	SAFEARRAY* psaHandles = SafeArrayCreate(VT_I4, 1, &cBounds);
	SAFEARRAY* psaErrors  = SafeArrayCreate(VT_I4, 1, &cBounds);

	// get the results of the write operation.
	int iOffset = sizeof(OPCGROUPHEADERWRITE);

	for (LONG ii = 1; ii <= (LONG)pHeader->dwItemCount; ii++, iOffset += sizeof(OPCITEMHEADERWRITE))
	{
		OPCITEMHEADERWRITE* pItemHeader = (OPCITEMHEADERWRITE*)&pBuffer[iOffset];

		OPCHANDLE hClient = NULL;
		HRESULT   hError  = OPC_E_INVALIDHANDLE;

		// lookup item from client handle.
		COPCItem* pItem = LookupItem(pItemHeader->hClient);
			
		if (pItem != NULL)
		{
			hClient = pItem->GetClientHandle();
			hError  = pItemHeader->dwError;
		}

		SafeArrayPutElement(psaHandles, &ii, &hClient);
		SafeArrayPutElement(psaErrors,  &ii, &hError);
	}

	// write to log file.
	TraceByRefOutArg(&psaHandles);
	TraceByRefOutArg(&psaErrors);

	// send the write complete notification.
	Fire_AsyncWriteComplete(pHeader->dwTransactionID, pHeader->dwItemCount, psaHandles, psaErrors);

	SafeArrayDestroy(psaHandles);
	SafeArrayDestroy(psaErrors);

	GlobalUnlock(pSTM->hGlobal);

	TRACE_FUNCTION_LEAVE();
	return;
}

//==============================================================================
// IAdviseSink::OnViewChange Method

STDMETHODIMP_(void) COPCGroup::OnViewChange(DWORD dwAspect, LONG lindex)
{
}

//==============================================================================
// IAdviseSink::OnRename Method

STDMETHODIMP_(void) COPCGroup::OnRename(LPMONIKER pmk)
{
}

//==============================================================================
// IAdviseSink::OnSave Method

STDMETHODIMP_(void) COPCGroup::OnSave()
{
}

//==============================================================================
// IAdviseSink::OnClose Method

STDMETHODIMP_(void) COPCGroup::OnClose()
{
}
