// ResStrings.h : Declaration of the CResStrings

#ifndef __RESSTRINGS_H_
#define __RESSTRINGS_H_

#include "resource.h"       // main symbols

/////////////////////////////////////////////////////////////////////////////
// CResStrings
class ATL_NO_VTABLE CResStrings : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CResStrings, &CLSID_ResStrings>,
	public ISupportErrorInfo,
	public IDispatchImpl<IResStrings, &IID_IResStrings, &LIBID_MLRESUTILLib>
{
public:
  CResStrings()
  {
    m_Dirty = VARIANT_FALSE ;
  }

DECLARE_REGISTRY_RESOURCEID(IDR_RESSTRINGS)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CResStrings)
  COM_INTERFACE_ENTRY(IResStrings)
  COM_INTERFACE_ENTRY(IDispatch)
  COM_INTERFACE_ENTRY(ISupportErrorInfo)
END_COM_MAP()

// ISupportsErrorInfo
  STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

// IResStrings
public:
	STDMETHOD(GetByLocaleAndResId)(/*[in]*/ long LocaleId, /*[in]*/ long ResId, /*[out, retval]*/ IResString** pVal);
	STDMETHOD(get_Dirty)(/*[out, retval]*/ VARIANT_BOOL *pVal);
	STDMETHOD(put_Dirty)(/*[in]*/ VARIANT_BOOL newVal);
  STDMETHOD(get_Locales)(/*[out, retval]*/ VARIANT *pVal);
  STDMETHOD(Exists)(/*[in]*/ long LocaleId, /*[in]*/ long ResId, /*[out, retval]*/ VARIANT_BOOL* StringExists);
  STDMETHOD(get_Item2)(long Index, /*[out, retval]*/ IResString* *pVal);
  STDMETHOD(SortById)();
  STDMETHOD(Clear)();
  STDMETHOD(Remove)(/*[in]*/ VARIANT Index);
  STDMETHOD(Add)(/*[in]*/ IResString* NewResString);
  STDMETHOD(get_Count)(/*[out, retval]*/ long *pVal);
  STDMETHOD(get_Item)(/*[in]*/ VARIANT Index, /*[out, retval]*/ IResString* *pVal);
  STDMETHOD(get__NewEnum)(/*[out, retval]*/ IEnumVARIANT* *pVal);

private:
  HRESULT GetIndex              ( VARIANT Index,      int* pOutIndex ) ;
  HRESULT GetIndexFromString    ( _bstr_t bstrSearch, int* pOutIndex ) ;
  HRESULT GetIndexFromInterface ( LPDISPATCH pDisp,   int* pOutIndex ) ;

  vector<IResString*>     m_vStrings ;
  VARIANT_BOOL            m_Dirty ;

};

#endif //__RESSTRINGS_H_
