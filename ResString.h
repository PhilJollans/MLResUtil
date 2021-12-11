// ResString.h : Declaration of the CResString

#ifndef __RESSTRING_H_
#define __RESSTRING_H_

#include "resource.h"       // main symbols

/////////////////////////////////////////////////////////////////////////////
// CResString
class ATL_NO_VTABLE CResString : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CResString, &CLSID_ResString>,
	public ISupportErrorInfo,
	public IDispatchImpl<IResString, &IID_IResString, &LIBID_MLRESUTILLib>
{
public:
	CResString()
	{
      m_lLocaleId = -1 ;
      m_lResId    = -1 ;
	}

DECLARE_REGISTRY_RESOURCEID(IDR_RESSTRING)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CResString)
	COM_INTERFACE_ENTRY(IResString)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
END_COM_MAP()

// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

// IResString
public:
	STDMETHOD(get_Length)(/*[out, retval]*/ long *pVal);
	STDMETHOD(get_CombinedId)(/*[out, retval]*/ long *pVal);
	STDMETHOD(get_LocaleId)(/*[out, retval]*/ long *pVal);
	STDMETHOD(put_LocaleId)(/*[in]*/ long newVal);
	STDMETHOD(get_ResId)(/*[out, retval]*/ long *pVal);
	STDMETHOD(put_ResId)(/*[in]*/ long newVal);
	STDMETHOD(get_Text)(/*[out, retval]*/ BSTR *pVal);
	STDMETHOD(put_Text)(/*[in]*/ BSTR newVal);

private:
    _bstr_t   m_bstrText ;
    long      m_lLocaleId ;
    long      m_lResId ;
};

#endif //__RESSTRING_H_
