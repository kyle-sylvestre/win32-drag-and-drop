//
//	DATAOBJECT.CPP
//
//	Implementation of the IDataObject COM interface
//
//	By J Brown 2004 
//
//	www.catch22.net
//

#include <windows.h>

// defined in enumformat.cpp
HRESULT CreateEnumFormatEtc(UINT nNumFormats, FORMATETC *pFormatEtc, IEnumFORMATETC **ppEnumFormatEtc);

class CDataObject : public IDataObject
{
public:
	//
    // IUnknown members
	//
    HRESULT __stdcall QueryInterface (REFIID iid, void ** ppvObject);
    ULONG   __stdcall AddRef (void);
    ULONG   __stdcall Release (void);
		
    //
	// IDataObject members
	//
    HRESULT __stdcall GetData				(FORMATETC *pFormatEtc,  STGMEDIUM *pMedium);
    HRESULT __stdcall GetDataHere			(FORMATETC *pFormatEtc,  STGMEDIUM *pMedium);
    HRESULT __stdcall QueryGetData			(FORMATETC *pFormatEtc);
	HRESULT __stdcall GetCanonicalFormatEtc (FORMATETC *pFormatEct,  FORMATETC *pFormatEtcOut);
    HRESULT __stdcall SetData				(FORMATETC *pFormatEtc,  STGMEDIUM *pMedium,  BOOL fRelease);
	HRESULT __stdcall EnumFormatEtc			(DWORD      dwDirection, IEnumFORMATETC **ppEnumFormatEtc);
	HRESULT __stdcall DAdvise				(FORMATETC *pFormatEtc,  DWORD advf, IAdviseSink *pAdvSink, DWORD *pdwConnection);
	HRESULT __stdcall DUnadvise				(DWORD      dwConnection);
	HRESULT __stdcall EnumDAdvise			(IEnumSTATDATA **ppEnumAdvise);
	
	//
    // Constructor / Destructor
	//
    CDataObject(FORMATETC *fmt, STGMEDIUM *stgmed, int count);
    ~CDataObject();
	
private:

	int LookupFormatEtc(FORMATETC *pFormatEtc);

    //
	// any private members and functions
	//
    LONG	   m_lRefCount;

	FORMATETC *m_pFormatEtc;
	STGMEDIUM *m_pStgMedium;
	LONG	   m_nNumFormats;

};

//
//	Constructor
//
CDataObject::CDataObject(FORMATETC *fmtetc, STGMEDIUM *stgmed, int count) 
{
	m_lRefCount  = 1;
	m_nNumFormats = count;
	
	m_pFormatEtc  = new FORMATETC[count];
	m_pStgMedium  = new STGMEDIUM[count];

	for(int i = 0; i < count; i++)
	{
		m_pFormatEtc[i] = fmtetc[i];
		m_pStgMedium[i] = stgmed[i];
	}
}

//
//	Destructor
//
CDataObject::~CDataObject()
{
	// cleanup
	if(m_pFormatEtc) delete[] m_pFormatEtc;
	if(m_pStgMedium) delete[] m_pStgMedium;
}

//
//	IUnknown::AddRef
//
ULONG __stdcall CDataObject::AddRef(void)
{
    // increment object reference count
    return InterlockedIncrement(&m_lRefCount);
}

//
//	IUnknown::Release
//
ULONG __stdcall CDataObject::Release(void)
{
    // decrement object reference count
	LONG count = InterlockedDecrement(&m_lRefCount);
		
	if(count == 0)
	{
		delete this;
		return 0;
	}
	else
	{
		return count;
	}
}

//
//	IUnknown::QueryInterface
//
HRESULT __stdcall CDataObject::QueryInterface(REFIID iid, void **ppvObject)
{
    // check to see what interface has been requested
    if(iid == IID_IDataObject || iid == IID_IUnknown)
    {
        AddRef();
        *ppvObject = this;
        return S_OK;
    }
    else
    {
        *ppvObject = 0;
        return E_NOINTERFACE;
    }
}

HGLOBAL DupMem(HGLOBAL hMem)
{
	// lock the source memory object
	SIZE_T  len    = GlobalSize(hMem);
	PVOID   source = GlobalLock(hMem);
	
	// create a fixed "global" block - i.e. just
	// a regular lump of our process heap
	PVOID   dest   = GlobalAlloc(GMEM_FIXED, len);

	memcpy(dest, source, len);

	GlobalUnlock(hMem);

	return dest;
}

int CDataObject::LookupFormatEtc(FORMATETC *pFormatEtc)
{
	for(int i = 0; i < m_nNumFormats; i++)
	{
		if((pFormatEtc->tymed    &  m_pFormatEtc[i].tymed)   &&
			pFormatEtc->cfFormat == m_pFormatEtc[i].cfFormat && 
			pFormatEtc->dwAspect == m_pFormatEtc[i].dwAspect)
		{
			return i;
		}
	}

	return -1;

}

//
//	IDataObject::GetData
//
HRESULT __stdcall CDataObject::GetData (FORMATETC *pFormatEtc, STGMEDIUM *pMedium)
{
	int idx;

	//
	// try to match the requested FORMATETC with one of our supported formats
	//
	if((idx = LookupFormatEtc(pFormatEtc)) == -1)
	{
		return DV_E_FORMATETC;
	}

	//
	// found a match! transfer the data into the supplied storage-medium
	//
	pMedium->tymed			 = m_pFormatEtc[idx].tymed;
	pMedium->pUnkForRelease  = 0;
	
	switch(m_pFormatEtc[idx].tymed)
	{
	case TYMED_HGLOBAL:

		pMedium->hGlobal = DupMem(m_pStgMedium[idx].hGlobal);
		//return S_OK;
		break;

	default:
		return DV_E_FORMATETC;
	}

	return S_OK;
}

//
//	IDataObject::GetDataHere
//
HRESULT __stdcall CDataObject::GetDataHere (FORMATETC *pFormatEtc, STGMEDIUM *pMedium)
{
	// GetDataHere is only required for IStream and IStorage mediums
	// It is an error to call GetDataHere for things like HGLOBAL and other clipboard formats
	//
	//	OleFlushClipboard 
	//
	return DATA_E_FORMATETC;
}

//
//	IDataObject::QueryGetData
//
//	Called to see if the IDataObject supports the specified format of data
//
HRESULT __stdcall CDataObject::QueryGetData (FORMATETC *pFormatEtc)
{
	return (LookupFormatEtc(pFormatEtc) == -1) ? DV_E_FORMATETC : S_OK;
}

//
//	IDataObject::GetCanonicalFormatEtc
//
HRESULT __stdcall CDataObject::GetCanonicalFormatEtc (FORMATETC *pFormatEct, FORMATETC *pFormatEtcOut)
{
	// Apparently we have to set this field to NULL even though we don't do anything else
	pFormatEtcOut->ptd = NULL;
	return E_NOTIMPL;
}

//
//	IDataObject::SetData
//
HRESULT __stdcall CDataObject::SetData (FORMATETC *pFormatEtc, STGMEDIUM *pMedium,  BOOL fRelease)
{
	return E_NOTIMPL;
}


//
//	IDataObject::EnumFormatEtc
//
HRESULT __stdcall CDataObject::EnumFormatEtc (DWORD dwDirection, IEnumFORMATETC **ppEnumFormatEtc)
{
	if(dwDirection == DATADIR_GET)
	{
		// for Win2k+ you can use the SHCreateStdEnumFmtEtc API call, however
		// to support all Windows platforms we need to implement IEnumFormatEtc ourselves.
		return CreateEnumFormatEtc(m_nNumFormats, m_pFormatEtc, ppEnumFormatEtc);
	}
	else
	{
		// the direction specified is not support for drag+drop
		return E_NOTIMPL;
	}
}

//
//	IDataObject::DAdvise
//
HRESULT __stdcall CDataObject::DAdvise (FORMATETC *pFormatEtc, DWORD advf, IAdviseSink *pAdvSink, DWORD *pdwConnection)
{
	return OLE_E_ADVISENOTSUPPORTED;
}

//
//	IDataObject::DUnadvise
//
HRESULT __stdcall CDataObject::DUnadvise (DWORD dwConnection)
{
	return OLE_E_ADVISENOTSUPPORTED;
}

//
//	IDataObject::EnumDAdvise
//
HRESULT __stdcall CDataObject::EnumDAdvise (IEnumSTATDATA **ppEnumAdvise)
{
	return OLE_E_ADVISENOTSUPPORTED;
}

//
//	Helper function
//
HRESULT CreateDataObject (FORMATETC *fmtetc, STGMEDIUM *stgmeds, UINT count, IDataObject **ppDataObject)
{
	if(ppDataObject == 0)
		return E_INVALIDARG;

	*ppDataObject = new CDataObject(fmtetc, stgmeds, count);

	return (*ppDataObject) ? S_OK : E_OUTOFMEMORY;
}

//
//	ENUMFORMAT.CPP
//
//	By J Brown 2004 
//
//	www.catch22.net
//
//	Implementation of the IEnumFORMATETC interface
//
//	For Win2K and above look at the SHCreateStdEnumFmtEtc API call!!
//
//	Apparently the order of formats in an IEnumFORMATETC object must be
//  the same as those that were stored in the clipboard
//
//

#include <windows.h>

class CEnumFormatEtc : public IEnumFORMATETC
{
public:

	//
	// IUnknown members
	//
	HRESULT __stdcall  QueryInterface (REFIID iid, void ** ppvObject);
	ULONG	__stdcall  AddRef (void);
	ULONG	__stdcall  Release (void);

	//
	// IEnumFormatEtc members
	//
	HRESULT __stdcall  Next  (ULONG celt, FORMATETC * rgelt, ULONG * pceltFetched);
	HRESULT __stdcall  Skip  (ULONG celt); 
	HRESULT __stdcall  Reset (void);
	HRESULT __stdcall  Clone (IEnumFORMATETC ** ppEnumFormatEtc);

	//
	// Construction / Destruction
	//
	CEnumFormatEtc(FORMATETC *pFormatEtc, int nNumFormats);
	~CEnumFormatEtc();

private:

	LONG		m_lRefCount;		// Reference count for this COM interface
	ULONG		m_nIndex;			// current enumerator index
	ULONG		m_nNumFormats;		// number of FORMATETC members
	FORMATETC * m_pFormatEtc;		// array of FORMATETC objects
};

//
//	"Drop-in" replacement for SHCreateStdEnumFmtEtc. Called by CDataObject::EnumFormatEtc
//
HRESULT CreateEnumFormatEtc(UINT nNumFormats, FORMATETC *pFormatEtc, IEnumFORMATETC **ppEnumFormatEtc)
{
	if(nNumFormats == 0 || pFormatEtc == 0 || ppEnumFormatEtc == 0)
		return E_INVALIDARG;

	*ppEnumFormatEtc = new CEnumFormatEtc(pFormatEtc, nNumFormats);

	return (*ppEnumFormatEtc) ? S_OK : E_OUTOFMEMORY;
}

//
//	Helper function to perform a "deep" copy of a FORMATETC
//
static void DeepCopyFormatEtc(FORMATETC *dest, FORMATETC *source)
{
	// copy the source FORMATETC into dest
	*dest = *source;
	
	if(source->ptd)
	{
		// allocate memory for the DVTARGETDEVICE if necessary
		dest->ptd = (DVTARGETDEVICE*)CoTaskMemAlloc(sizeof(DVTARGETDEVICE));

		// copy the contents of the source DVTARGETDEVICE into dest->ptd
		*(dest->ptd) = *(source->ptd);
	}
}

//
//	Constructor 
//
CEnumFormatEtc::CEnumFormatEtc(FORMATETC *pFormatEtc, int nNumFormats)
{
	m_lRefCount   = 1;
	m_nIndex      = 0;
	m_nNumFormats = nNumFormats;
	m_pFormatEtc  = new FORMATETC[nNumFormats];
	
	// copy the FORMATETC structures
	for(int i = 0; i < nNumFormats; i++)
	{	
		DeepCopyFormatEtc(&m_pFormatEtc[i], &pFormatEtc[i]);
	}
}

//
//	Destructor
//
CEnumFormatEtc::~CEnumFormatEtc()
{
	if(m_pFormatEtc)
	{
		for(ULONG i = 0; i < m_nNumFormats; i++)
		{
			if(m_pFormatEtc[i].ptd)
				CoTaskMemFree(m_pFormatEtc[i].ptd);
		}

		delete[] m_pFormatEtc;
	}
}

//
//	IUnknown::AddRef
//
ULONG __stdcall CEnumFormatEtc::AddRef(void)
{
    // increment object reference count
    return InterlockedIncrement(&m_lRefCount);
}

//
//	IUnknown::Release
//
ULONG __stdcall CEnumFormatEtc::Release(void)
{
    // decrement object reference count
	LONG count = InterlockedDecrement(&m_lRefCount);
		
	if(count == 0)
	{
		delete this;
		return 0;
	}
	else
	{
		return count;
	}
}

//
//	IUnknown::QueryInterface
//
HRESULT __stdcall CEnumFormatEtc::QueryInterface(REFIID iid, void **ppvObject)
{
    // check to see what interface has been requested
    if(iid == IID_IEnumFORMATETC || iid == IID_IUnknown)
    {
        AddRef();
        *ppvObject = this;
        return S_OK;
    }
    else
    {
        *ppvObject = 0;
        return E_NOINTERFACE;
    }
}

//
//	IEnumFORMATETC::Next
//
//	If the returned FORMATETC structure contains a non-null "ptd" member, then
//  the caller must free this using CoTaskMemFree (stated in the COM documentation)
//
HRESULT __stdcall CEnumFormatEtc::Next(ULONG celt, FORMATETC *pFormatEtc, ULONG * pceltFetched)
{
	ULONG copied  = 0;

	// validate arguments
	if(celt == 0 || pFormatEtc == 0)
		return E_INVALIDARG;

	// copy FORMATETC structures into caller's buffer
	while(m_nIndex < m_nNumFormats && copied < celt)
	{
		DeepCopyFormatEtc(&pFormatEtc[copied], &m_pFormatEtc[m_nIndex]);
		copied++;
		m_nIndex++;
	}

	// store result
	if(pceltFetched != 0) 
		*pceltFetched = copied;

	// did we copy all that was requested?
	return (copied == celt) ? S_OK : S_FALSE;
}

//
//	IEnumFORMATETC::Skip
//
HRESULT __stdcall CEnumFormatEtc::Skip(ULONG celt)
{
	m_nIndex += celt;
	return (m_nIndex <= m_nNumFormats) ? S_OK : S_FALSE;
}

//
//	IEnumFORMATETC::Reset
//
HRESULT __stdcall CEnumFormatEtc::Reset(void)
{
	m_nIndex = 0;
	return S_OK;
}

//
//	IEnumFORMATETC::Clone
//
HRESULT __stdcall CEnumFormatEtc::Clone(IEnumFORMATETC ** ppEnumFormatEtc)
{
	HRESULT hResult;

	// make a duplicate enumerator
	hResult = CreateEnumFormatEtc(m_nNumFormats, m_pFormatEtc, ppEnumFormatEtc);

	if(hResult == S_OK)
	{
		// manually set the index state
		((CEnumFormatEtc *) *ppEnumFormatEtc)->m_nIndex = m_nIndex;
	}

	return hResult;
}

