// ResString.cpp : Implementation of CResString
#include "stdafx.h"
#include "ResString.h"

/////////////////////////////////////////////////////////////////////////////
// CResString

STDMETHODIMP CResString::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_IResString
	};
	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}

STDMETHODIMP CResString::get_Text(BSTR *pVal)
{
  *pVal = m_bstrText.copy() ;
  return S_OK;
}

STDMETHODIMP CResString::put_Text(BSTR newVal)
{
  m_bstrText = newVal ;
  return S_OK;
}

STDMETHODIMP CResString::get_ResId(long *pVal)
{
  *pVal = m_lResId ;
  return S_OK;
}

STDMETHODIMP CResString::put_ResId(long newVal)
{
  m_lResId = newVal ;
  return S_OK;
}

STDMETHODIMP CResString::get_LocaleId(long *pVal)
{
  *pVal = m_lLocaleId ;
  return S_OK;
}

STDMETHODIMP CResString::put_LocaleId(long newVal)
{
  m_lLocaleId = newVal ;
  return S_OK;
}

STDMETHODIMP CResString::get_CombinedId(long *pVal)
{
  *pVal = ( m_lLocaleId << 16 ) | m_lResId ;
  return S_OK;
}

STDMETHODIMP CResString::get_Length(long *pVal)
{
  *pVal = m_bstrText.length() ;
  return S_OK;
}
