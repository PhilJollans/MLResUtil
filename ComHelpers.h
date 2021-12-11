//
// some often used Macros for COM..
//
#define CHECK_PTR(p)	if (p==NULL) return E_POINTER;

// Wrapper class for HRESULT, which throws an exception when
// assigned a "FAILED" value.
class HRX
{
public:
  HRX() 
  { 
    hr = S_OK ;
  }

  HRX ( HRESULT hresult )
  {
    hr = hresult ;
    if ( FAILED(hr) )
      throw hr ; 
  }

  HRX& operator= ( HRESULT hresult )
  {
    hr = hresult ;
    if ( FAILED(hr) )
      throw hr ; 
    return *this ;
  }

  operator HRESULT()
  { 
    return hr ;
  }

  // Set the internal value without throwing an exception.
  void SetNoThrow ( HRESULT hresult )
  {
    hr = hresult ;
  }

private:
  HRESULT hr ;
} ;

//
// ClearVector function for any vector class
//
template<class T>
void ClearVector(T& vec)
{
  vec.erase ( vec.begin(), vec.end() ) ;
}

