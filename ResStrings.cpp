// ResStrings.cpp : Implementation of CResStrings
#include "stdafx.h"
#include "ResStrings.h"

/////////////////////////////////////////////////////////////////////////////
// CResStrings

STDMETHODIMP CResStrings::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_IResStrings
	};
	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}

STDMETHODIMP CResStrings::get__NewEnum(IEnumVARIANT* *pVal)
{
  // Boilerplate code!
  typedef CComObject < CComEnum < IEnumVARIANT,
                                  &IID_IEnumVARIANT,
                                  VARIANT,
                                  _Copy<VARIANT> > > EnumType ;

  long                            lCount = m_vStrings.size() ;
  VARIANT*                        var = new VARIANT[lCount] ;
  vector<IResString*>::iterator   it ;
  int                             i ;

  for ( it = m_vStrings.begin(), i = 0 ; 
        it != m_vStrings.end() ;
        it++, i++ )
  {
    VariantInit ( &var[i] ) ;
    var[i].vt = VT_DISPATCH ;
    (*it)->QueryInterface ( IID_IDispatch, (void**)&(var[i].pdispVal) ) ;
  }

  EnumType* pEnum = new EnumType ;
  pEnum->Init ( &var[0], &var[lCount], NULL, AtlFlagCopy ) ;

  for ( i = 0 ; i < lCount ; i++ )
  {
    VariantClear ( &var[i] ) ;
  }
  delete[] var ;

  pEnum->QueryInterface ( IID_IUnknown, (void**)pVal ) ;

  return S_OK;
}


STDMETHODIMP CResStrings::get_Item(VARIANT Index, IResString **pVal)
{
  HRESULT hRes = S_OK ;
  int     iInd ;

  hRes = GetIndex ( Index, &iInd ) ;

  if ( hRes == S_OK )
  {
    m_vStrings[iInd]->QueryInterface ( IID_IResString, (void**)pVal ) ;
  }

  return hRes ;
}

STDMETHODIMP CResStrings::get_Count(long *pVal)
{
  *pVal = m_vStrings.size() ;
  return S_OK;
}

STDMETHODIMP CResStrings::Add(IResString *NewResString)
{
  HRESULT         hr = S_OK ;
  IResString*     pResString ;

  // Reserve space in bigish chunks
  if ( m_vStrings.size() == m_vStrings.capacity() )
  {
    m_vStrings.reserve ( m_vStrings.size() + 200 ) ;
  } 

  // QI to check that it really is a ResString and AddRef at the same time.
  hr = NewResString->QueryInterface ( IID_IResString, (void**)&pResString ) ;
  if ( hr == S_OK )
  {
    m_vStrings.push_back ( pResString ) ;
    m_Dirty = VARIANT_TRUE ;
  }
  return hr ;
}

STDMETHODIMP CResStrings::Remove(VARIANT Index)
{
  HRESULT hRes = S_OK ;
  int     iInd ;

  hRes = GetIndex ( Index, &iInd ) ;

  if ( hRes == S_OK )
  {
    // Release the interface
    m_vStrings[iInd]->Release() ;
    
    // Remove the interface from the vector
    //m_vStrings.erase ( &m_vStrings[iInd] ) ;
    // Modified for VS 2005 (11-June-2007)
    // In fact, I don't think this method is used in the Multi-Language Add-In anyway!
    m_vStrings.erase ( m_vStrings.begin() + iInd ) ;

    // Set the dirty flag
    m_Dirty = VARIANT_TRUE ;
  }

  return hRes ;
}

HRESULT CResStrings::GetIndexFromString ( _bstr_t bstrSearch, int* pOutIndex )
{
  HRESULT hRes = S_OK ;
  int     iOutIndex ;
  int     iMax = m_vStrings.size() ;

  for ( iOutIndex = 0 ; iOutIndex < iMax ; iOutIndex++ )
  {
    BSTR        BSTRText ;
    hRes = m_vStrings[iOutIndex]->get_Text ( &BSTRText ) ;
    _bstr_t     bstrText ( BSTRText, false ) ;

    if ( bstrSearch == bstrText )
      break ;
  }

  // Pass the index back to the caller.
  *pOutIndex = iOutIndex ;

  if ( iOutIndex == iMax )
    hRes = Error("Resourse string not found.") ;
  else
    hRes = S_OK ;

  return hRes ;
}

HRESULT CResStrings::GetIndexFromInterface ( LPDISPATCH pDisp, int* pOutIndex )
{
  HRESULT   hRes = S_OK ;

  // Get the IUnknown pointer of the parameter
  if ( pDisp == NULL )
  {
    *pOutIndex = -1 ;
    hRes = Error ( "No interface specified to GetIndexFromInterface()" ) ;
  }
  else
  {
    int       iOutIndex ;
    int       iMax = m_vStrings.size() ;
    LPUNKNOWN pUnk ;

    hRes = pDisp->QueryInterface ( IID_IUnknown, (void**)&pUnk ) ;

    for ( iOutIndex = 0 ; iOutIndex < iMax ; iOutIndex++ )
    {
      LPUNKNOWN   pUnk2 ;
      hRes = m_vStrings[iOutIndex]->QueryInterface ( IID_IUnknown, (void**)&pUnk2 ) ;

      bool found = ( pUnk == pUnk2 ) ;
      pUnk2->Release() ;

      if ( found )
        break ;
    }
    pUnk->Release() ;

    if ( iOutIndex == iMax )
    {
      *pOutIndex = -1 ;
      hRes = Error("ResString not found in collection.") ;
    }
    else
    {
      *pOutIndex = iOutIndex ;
      hRes = S_OK ;
    }
  }

  return hRes ;
}


HRESULT CResStrings::GetIndex ( VARIANT Index, int* pOutIndex )
{
  HRESULT hRes = S_OK ;
  int     iOutIndex = 0 ;

  switch ( Index.vt )
  {
    case VT_I4:
      // Always zero based - even in VB
      iOutIndex = Index.lVal ;
      break ;

    case VT_I2:
      // Always zero based - even in VB
      iOutIndex = Index.iVal ;
      break ;

    case VT_BSTR:
      {
        _bstr_t strParam = Index.bstrVal ;
        hRes = GetIndexFromString ( strParam, &iOutIndex ) ;
      }
      break ;

    case VT_DISPATCH:
    case VT_UNKNOWN:
      {
        hRes = GetIndexFromInterface ( Index.pdispVal, &iOutIndex ) ;
      }
      break ;

    default:
      {
        hRes = Error("Index parameter has invalid variant type.") ;
      }
      break ;
  }

  if ( hRes == S_OK )
  {
    if ( iOutIndex >= m_vStrings.size() )
      hRes = Error ( "Index out of range for resource strings collection" ) ;
  }

  // Pass the index back to the caller.
  *pOutIndex = iOutIndex ;

  return hRes ;
}


STDMETHODIMP CResStrings::Clear()
{
  // Release all of the interfaces
  for ( int i = 0 ; i < m_vStrings.size() ; i++ )
    m_vStrings[i]->Release() ;

  // Clean up vectors.
  ClearVector ( m_vStrings ) ;

  // Set the dirty flag
  m_Dirty = VARIANT_TRUE ;

  return S_OK;
}

class CombiIdLess
{
public:
  bool operator()(IResString* p1, IResString* p2)
  {
    long  c1, c2 ;
    p1->get_CombinedId ( &c1 ) ;
    p2->get_CombinedId ( &c2 ) ;
    return ( c1 < c2 ) ;
  }
} ;

STDMETHODIMP CResStrings::SortById()
{
  sort ( m_vStrings.begin(), m_vStrings.end(), CombiIdLess() ) ;
  return S_OK ;
}

STDMETHODIMP CResStrings::get_Item2(long Index, IResString **pVal)
{
  HRESULT hRes = S_OK ;

  if ( Index >= m_vStrings.size() )
    hRes = Error ( "Index out of range for resource strings collection" ) ;
  else
    hRes = m_vStrings[Index]->QueryInterface ( IID_IResString, (void**)pVal ) ;

  return hRes ;
}

STDMETHODIMP CResStrings::Exists(long LocaleId, long ResId, VARIANT_BOOL* StringExists)
{
  HRX             hr ;

  try
  {
    // Assume it doesn't exist
    *StringExists = VARIANT_FALSE ;
  
    // Define an iterator for the vector
    vector<IResString*>::iterator   it ;

    for ( it = m_vStrings.begin() ; it != m_vStrings.end() ; it++ )
    {
      long  lStringResId ;
      long  lStringLocaleId ;

      hr = (*it)->get_ResId ( &lStringResId ) ;
      if ( lStringResId == ResId )
      {
        hr = (*it)->get_LocaleId ( &lStringLocaleId ) ;
        if ( lStringLocaleId == LocaleId )
        {
          *StringExists = VARIANT_TRUE ;
          break ;
        }
      }
    }
  }
  catch ( _com_error e )  
  { 
    hr.SetNoThrow ( Error( (LPCOLESTR)e.Description()) ) ; 
  }
  catch ( ... )
  {
    if ( hr == S_OK )
      hr.SetNoThrow ( Error("Exception in CResStrings::Exists()") ) ;
  }

  return hr ;
}

STDMETHODIMP CResStrings::get_Locales(VARIANT *pVal)
{
  HRX             hr ;

  try
  {
    // Get the Locales in an STL set
    set<long>   SetOfLocales ;

    // Define an iterator for the vector
    vector<IResString*>::iterator   itStrings ;

    for ( itStrings = m_vStrings.begin() ; itStrings != m_vStrings.end() ; itStrings++ )
    {
      long  lStringLocaleId ;
      hr = (*itStrings)->get_LocaleId ( &lStringLocaleId ) ;

      SetOfLocales.insert ( lStringLocaleId ) ; 
    }

    // Create a SAFEARRAY
    SAFEARRAY*  psa;
    long*       pArray ;

    psa = SafeArrayCreateVector ( VT_I4, 0, SetOfLocales.size() ) ;
    SafeArrayAccessData ( psa, (void**)&pArray ) ;

    // Copy the data
    set<long>::iterator   itLocales ;
    for ( itLocales = SetOfLocales.begin() ; itLocales != SetOfLocales.end() ; itLocales++ )
      *pArray++ = *itLocales ;

    // Release access to the safearray
    SafeArrayUnaccessData(psa);

    // Initialise the variant return value
    pVal->vt     = VT_ARRAY | VT_I4 ;
    pVal->parray = psa ;

  }
  catch ( _com_error e )  
  { 
    hr.SetNoThrow ( Error( (LPCOLESTR)e.Description()) ) ; 
  }
  catch ( ... )
  {
    if ( hr == S_OK )
      hr.SetNoThrow ( Error("Exception in CResStrings::get_Locales()") ) ;
  }

  return hr ;
}

//
// The dirty property is exposed via the IResFile interface.
// The dirty property of the IResStrings interface is used
// internally and is hidden and restricted for external
// components.
//
STDMETHODIMP CResStrings::get_Dirty(VARIANT_BOOL *pVal)
{
  *pVal = m_Dirty ;
  return S_OK ;
}

STDMETHODIMP CResStrings::put_Dirty(VARIANT_BOOL newVal)
{
  m_Dirty = newVal ;
  return S_OK ;
}

STDMETHODIMP CResStrings::GetByLocaleAndResId(long LocaleId, long ResId, IResString **pVal)
{
  HRX             hr ;

  try
  {
    // Assume it doesn't exist
    *pVal = NULL ;
  
    // Define an iterator for the vector
    vector<IResString*>::iterator   it ;

    for ( it = m_vStrings.begin() ; it != m_vStrings.end() ; it++ )
    {
      long  lStringResId ;
      long  lStringLocaleId ;

      hr = (*it)->get_ResId ( &lStringResId ) ;
      if ( lStringResId == ResId )
      {
        hr = (*it)->get_LocaleId ( &lStringLocaleId ) ;
        if ( lStringLocaleId == LocaleId )
        {
          hr = (*it)->QueryInterface ( IID_IResString, (void**)pVal ) ;
          break ;
        }
      }
    }
  }
  catch ( _com_error e )  
  { 
    hr.SetNoThrow ( Error( (LPCOLESTR)e.Description()) ) ; 
  }
  catch ( ... )
  {
    if ( hr == S_OK )
      hr.SetNoThrow ( Error("Exception in CResStrings::GetByLocaleAndResId()") ) ;
  }

  return hr ;
}
