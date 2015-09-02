// Connect.h : Declaration of the CConnect

#pragma once
#include "resource.h"       // main symbols
#include "ResourceHelpers.h"
#include "TestAddin2_i.h"
#include "FormRegionWrapper.h";

// ###00000
// ------------------------------------------
// This file is built off the tutorial found at
// https://msdn.microsoft.com/en-us/library/office/ee941475%28v=office.14%29.aspx
// (retrieved: 1 sept 2015)
//    Added annotation is marked by the ###00nnn comments.

#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif

// ###00001
// >> Step 3: Implement Interfaces
//---------------------------------
// COM interfaces all derive from IDispatch.

// The IDispatchImpl template declares the major/minor version number
// of the interface to look up **in the registry**.

// Note: if the Major and minor version numbers are 0xFFFF, the
// IDispatchImpl class will look for the version numbers in the library
// file.
// see: http://stackoverflow.com/questions/20152185/using-idispatchimpl-with-an-unregistered-interface-in-an-mfcatl-exe

// Problem #1: if you have multiple version installed, you get pot luck
// which library gets found.  And if you are not writing for the least
// common denominator, you may get unexpected behavior.

// Problem #2: ... what if the library is not included in the path?

// Note that a side effect of looking for it in the registry is that
// you have to have it registered (IE installed).  That is, ... have
// Outlook installed.

// ------------------------------------------------------------------------


// ###00002
// >> Step 3: Implement Interfaces
//---------------------------------
// ... So Why a typedef in the first place?  Two reasons...
//
// 1) because this is a long line, subject to possible change, and it
// makes for a simpler (appearing) class definition.

// 2) because IRibbonExtensibility is not in a standard interface
// library used by VS 2010.  The walkthrough could not direct you to
// add it to the project in the same way it directed you to add the
// IDTExtensibility2 object. It needed to be added by hand, so why not
// have a copy/paste example to draw from?

// Side effect: as we didn't use the Wizard to add the Ribbon Extension
// object, we have to do the handling (like was done for
// IDTExtensibility2) by hand.

// See also: (creating a stand alone type library for the ribbon)
// https://xldennis.wordpress.com/2006/12/12/creating-a-standalone-type-library-for-iribbonextensibility/


typedef IDispatchImpl<_IDTExtensibility2, &__uuidof(_IDTExtensibility2), &__uuidof(__AddInDesignerObjects), /* wMajor = */ 1>
    IDTImpl;
// See also: (IDTExtensibility2 interface)
// https://msdn.microsoft.com/en-us/library/extensibility.idtextensibility2.aspx

typedef IDispatchImpl<IRibbonCallback, &__uuidof(IRibbonCallback)> 
	CallbackImpl;
typedef IDispatchImpl<_FormRegionStartup, &__uuidof(_FormRegionStartup), &__uuidof(__Outlook), /* wMajor = */ 9, /* wMinor = */ 4>
	FormImpl;

typedef IDispatchImpl<IRibbonExtensibility, &__uuidof(IRibbonExtensibility), &__uuidof(__Office), /* wMajor = */ 2, /* wMinor = */ 5>
	RibbonImpl;
// CConnect

class ATL_NO_VTABLE CConnect :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CConnect, &CLSID_Connect>,
	public IDispatchImpl<IConnect, &IID_IConnect, &LIBID_TestAddin2Lib, /*wMajor =*/ 1, /*wMinor =*/ 0>,
	public IDTImpl,	
	public RibbonImpl,
	public FormImpl,

	public CallbackImpl
{
public:
	CConnect()
	{
	}
	STDMETHOD(Invoke)(DISPID dispidMember, const IID &riid, LCID lcid, WORD wFlags, DISPPARAMS *pdispparams, VARIANT *pvarResult, EXCEPINFO *pexceptinfo, UINT *puArgErr)
	{
		HRESULT hr=DISP_E_MEMBERNOTFOUND;
		if(dispidMember==42)
		{
			hr  = CallbackImpl::Invoke(dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexceptinfo, puArgErr);
		}
		if (DISP_E_MEMBERNOTFOUND == hr)
			hr = FormImpl::Invoke(dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexceptinfo, puArgErr);

		if (DISP_E_MEMBERNOTFOUND == hr)
			hr = IDTImpl::Invoke(dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexceptinfo, puArgErr);

		return hr;
	}
	DECLARE_REGISTRY_RESOURCEID(IDR_CONNECT)

	BEGIN_COM_MAP(CConnect)
		COM_INTERFACE_ENTRY2(IDispatch, IRibbonCallback)
		COM_INTERFACE_ENTRY(IConnect)		
		COM_INTERFACE_ENTRY(_IDTExtensibility2)
		COM_INTERFACE_ENTRY(_FormRegionStartup)
		COM_INTERFACE_ENTRY(IRibbonExtensibility)
		COM_INTERFACE_ENTRY(IRibbonCallback)
	END_COM_MAP()

	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
	}

public:


	// _IDTExtensibility2 Methods
public:
        // ###00003
        // >> Step 3: Implement Interfaces
        //---------------------------------
        // These functions are all used by Outlook, called at
        // appropriate times.  Bad things happen if we blow off these
        // functions (by leaving the stubs returning E_NOTIMPL).  As we
        // don't have particular needs on these events, we don't need
        // to do more than let them return "success".

        STDMETHOD(OnConnection)(LPDISPATCH Application,
                                ext_ConnectMode ConnectMode,
                                LPDISPATCH AddInInst,
                                SAFEARRAY * * custom
                               )
        {
            // ###00004
            // >> Step 4: Set up loading add-in in Outlook
            // -------------------------------------------
            // for debugging purposes, the tutorial temporarily adds a
            // MessageBox call here, to show that the add-in has
            // actually been loaded.  It is taken out soon after, once
            // you've verified that your add-in is correctly loading.
		return S_OK;
	}
	STDMETHOD(OnDisconnection)(ext_DisconnectMode RemoveMode, SAFEARRAY * * custom)
	{
		return S_OK;
	}
	STDMETHOD(OnAddInsUpdate)(SAFEARRAY * * custom)
	{
		return S_OK;
	}
	STDMETHOD(OnStartupComplete)(SAFEARRAY * * custom)
	{
		return S_OK;
	}
	STDMETHOD(OnBeginShutdown)(SAFEARRAY * * custom)
	{
		if(m_pFormRegionWrapper)
			m_pFormRegionWrapper->OnFormRegionClose();
		return S_OK;
	}

	// _FormRegionStartup Methods
public:
	STDMETHOD(GetFormRegionStorage)(BSTR FormRegionName, LPDISPATCH Item, long LCID, OlFormRegionMode FormRegionMode, OlFormRegionSize FormRegionSize, VARIANT * Storage)
	{
			V_VT(Storage) = VT_ARRAY | VT_UI1;
			V_ARRAY(Storage) = GetOFSResource(IDR_FORMREGION);
			return S_OK;
	}
	STDMETHOD(BeforeFormRegionShow)(_FormRegion * FormRegion)
	{
		m_pFormRegionWrapper = new FormRegionWrapper();
		if (!m_pFormRegionWrapper)
			return E_OUTOFMEMORY;
		return m_pFormRegionWrapper->HrInit(FormRegion);

	}
	STDMETHOD(GetFormRegionManifest)(BSTR FormRegionName, long LCID, VARIANT * Manifest)
	{
			V_VT(Manifest) = VT_BSTR;
			BSTR bstrManifest = GetXMLResource(IDR_XML2);
			V_BSTR(Manifest) = bstrManifest;
			return S_OK;
	}
	STDMETHOD(GetFormRegionIcon)(BSTR FormRegionName, long LCID, OlFormRegionIcon Icon, VARIANT * Result)
	{
		return S_OK;
	}

	// IRibbonExtensibility Methods
public:

	STDMETHOD(GetCustomUI)(BSTR RibbonID, BSTR * RibbonXml)
	{
		if(!RibbonXml)
			return E_POINTER;
		*RibbonXml = GetXMLResource(IDR_XML3);
		return S_OK;
	}

	// IRibbonCallback Methods
public:
	STDMETHOD(ButtonClicked)(IDispatch* ribbon)
	{
		if(m_pFormRegionWrapper)
		{
			m_pFormRegionWrapper->Show();
			m_pFormRegionWrapper->SearchSelection();
			
		}
		return S_OK;
	}
HRESULT HrGetResource(int nId, LPCTSTR lpType, LPVOID* ppvResourceData, DWORD* pdwSizeInBytes)
{
	HMODULE hModule = _AtlBaseModule.GetModuleInstance();
	
	if (!hModule)
		return E_UNEXPECTED;

	HRSRC hRsrc = FindResource(hModule, MAKEINTRESOURCE(nId), lpType);

	if (!hRsrc)
		return HRESULT_FROM_WIN32(GetLastError());

	HGLOBAL hGlobal = LoadResource(hModule, hRsrc);

	if (!hGlobal)
		return HRESULT_FROM_WIN32(GetLastError());

	*pdwSizeInBytes = SizeofResource(hModule, hRsrc);
	*ppvResourceData = LockResource(hGlobal);
	
	return S_OK;
}

BSTR GetXMLResource(int nId)
{
	LPVOID pResourceData = NULL;
	DWORD dwSizeInBytes = 0;

	HRESULT hr = HrGetResource(nId, TEXT("XML"), &pResourceData, &dwSizeInBytes);
	if (FAILED(hr))
		return NULL;

	// Assumes that the data is not store in Unicode.
	CComBSTR cbstr(dwSizeInBytes, reinterpret_cast<LPCSTR>(pResourceData));

	return cbstr.Detach();
}

SAFEARRAY* GetOFSResource(int nId)
{
	LPVOID pResourceData = NULL;
	DWORD dwSizeInBytes = 0;

	if (FAILED(HrGetResource(nId, TEXT("OFS"), &pResourceData, &dwSizeInBytes)))
		return NULL;

	SAFEARRAY* psa;
	SAFEARRAYBOUND dim = {dwSizeInBytes, 0};

	psa = SafeArrayCreate(VT_UI1, 1, &dim);

	if (psa == NULL)
		return NULL;

	BYTE* pSafeArrayData;

	SafeArrayAccessData(psa, (void**)&pSafeArrayData);

	memcpy((void*)pSafeArrayData, pResourceData, dwSizeInBytes);
	
	SafeArrayUnaccessData(psa);
	
	return psa;
}
private:
	FormRegionWrapper* m_pFormRegionWrapper;
};

OBJECT_ENTRY_AUTO(__uuidof(Connect), CConnect)
