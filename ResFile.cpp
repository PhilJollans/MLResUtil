// ResFile.cpp : Implementation of CResFile
#include "stdafx.h"
#include "ResFile.h"

/////////////////////////////////////////////////////////////////////////////
// CResFile

STDMETHODIMP CResFile::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_IResFile
	};
	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}

STDMETHODIMP CResFile::get_Filename(BSTR *pVal)
{
  *pVal = m_bstrFilename.copy() ;
  return S_OK;
}

STDMETHODIMP CResFile::put_Filename(BSTR newVal)
{
  HRESULT     hRes ;

  m_bstrFilename = newVal ;

  // Set the dirty flag.
  hRes = m_strings->put_Dirty ( VARIANT_TRUE ) ;

  return hRes ;
}

STDMETHODIMP CResFile::Import()
{
  HRX           hr ;
  ifstream      input ;

  try
  {
    // Clear the string collection
    m_strings->Clear() ;

    // Open the input file
    input.open ( (const wchar_t *)m_bstrFilename, ios::in | ios::binary ) ;
    if ( !input.is_open() )
      hr = Error ( "Failed to open resource file" ) ;

    // Enable exceptions on the input stream
    input.exceptions ( ios_base::badbit | ios_base::failbit | ios_base::eofbit ) ;

    while ( !input.eof() )
    {
      hr = import_single_resource ( input ) ;
    }

    // Clear the dirty flag.
    hr = m_strings->put_Dirty ( VARIANT_FALSE ) ;

    // John Charlton has been having problems, which might be because the 
    // resource file is not being closed. Previously, we relied on the 
    // destructor to close the files, which was reliable for me, but 
    // nevertheless, we will now close the files explicitly.
    input.close() ;
  }
  catch ( long )
  {
    // This catches the HRESULT type, so check hr.
    if ( hr == S_OK )
    {
      if ( input.eof() )
        hr.SetNoThrow ( Error ( "Unexpected end of file reading resource file" ) ) ;
      else if ( input.fail() )
        hr.SetNoThrow ( Error ( "Unexpected I/O error reading resource file" ) ) ;
      else
        hr.SetNoThrow ( Error ( "Unexpected exception reading resource file" ) ) ;
    }
  }
  catch ( _com_error e )  { hr.SetNoThrow ( Error( (LPCOLESTR)e.Description()) ) ; }
  catch ( ... )           { hr.SetNoThrow ( Error("Exception in CResFile::Import()") ) ;    }

  return hr ;
}

STDMETHODIMP CResFile::Export(VARIANT_BOOL PreserveStrings)
{
  USES_CONVERSION ;
  HRX             hr ;
  ofstream        output ;
  ifstream        input ;
  BOOL		      bOK ;

  try
  {
    TCHAR       szdrive      [ _MAX_DRIVE ] ;
    TCHAR       szdir        [ _MAX_DIR   ] ;
    TCHAR       szfname      [ _MAX_FNAME ] ;
    TCHAR       szext        [ _MAX_EXT   ] ;
    TCHAR       szBackupName [ _MAX_PATH  ] ;
    CStdString  ssMsg ;
    LPVOID      lpMsgBuf ;


    // Create a name for the backup file.
    _splitpath ( OLE2T(m_bstrFilename), szdrive, szdir, szfname, szext ) ;
    _makepath ( szBackupName, szdrive, szdir, szfname, ".BAK" ) ;

    // Determine whether the file exists
    if ( ( _taccess ( szBackupName, 0 ) ) == 0 ) 
    {
      // Backup file already exists
      // Don't check read only bit. Just reset it!
      // Then delete the file.
      _tchmod  ( szBackupName, S_IREAD | S_IWRITE ) ;
      _tunlink ( szBackupName ) ;
    }

    // Now move the resource file to the backup.
    bOK = MoveFile ( OLE2T(m_bstrFilename), szBackupName ) ;
    if ( !bOK )
    {
      FormatMessage ( FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
                      NULL,
                      GetLastError(),
                      MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), // Default language
                      (LPTSTR) &lpMsgBuf,
                      0,
                      NULL ) ;

      ssMsg.Format ( "Failed to rename file %s to %s\n%s", OLE2T(m_bstrFilename), szBackupName, lpMsgBuf ) ;
      hr = Error ( ssMsg.c_str() ) ;
    }

    // Open the backup file for input
    input.open ( szBackupName, ios::in | ios::binary  ) ;
    if ( !input.is_open() )
    {
      FormatMessage ( FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
                      NULL,
                      GetLastError(),
                      MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), // Default language
                      (LPTSTR) &lpMsgBuf,
                      0,
                      NULL ) ;

      ssMsg.Format ( "Failed to open resource file %s for input\n%s", szBackupName, lpMsgBuf ) ;
      hr = Error ( ssMsg.c_str() ) ;
    }

    // Enable exceptions on the input stream
    input.exceptions ( ios_base::badbit | ios_base::failbit | ios_base::eofbit ) ;

    // Open the output file
    output.open ( (const wchar_t *)m_bstrFilename, ios::binary  ) ;
    if ( !output.is_open() )
    {
      FormatMessage ( FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
                      NULL,
                      GetLastError(),
                      MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), // Default language
                      (LPTSTR) &lpMsgBuf,
                      0,
                      NULL ) ;

      ssMsg.Format ( "Failed to open resource file %s for output\n%s", OLE2T(m_bstrFilename), lpMsgBuf ) ;
      hr = Error ( ssMsg.c_str() ) ;
    }

    // Enable exceptions for the output stream.
    output.exceptions ( ios_base::badbit | ios_base::failbit | ios_base::eofbit ) ;

    // Loop through the input file and:
    // - Copy all non string resources to the output file.
    // - If PreserveStrings flag is set, import any strings which
    //   do not conflict with strings already in the collection.
    //   These will later be written to the output file.
    while ( !input.eof() )
    {
      hr = transfer_single_resource ( input, output, PreserveStrings ) ;
    }

    // Common function for actually exporting the strings.
    export_string_resources ( output ) ;
    
    // John Charlton has been having problems, which might be because the 
    // resource file is not being closed. Previously, we relied on the 
    // destructor to close the files, which was reliable for me, but 
    // nevertheless, we will now close the files explicitly.
    output.close() ;
    input.close() ;
  }
  catch ( long )
  {
    // This catches the HRESULT type, so check hr.
    if ( hr == S_OK )
    {
      if ( input.eof() )
        hr.SetNoThrow ( Error ( "Unexpected end of file reading resource file" ) ) ;
      else if ( input.fail() )
        hr.SetNoThrow ( Error ( "Unexpected I/O error reading resource file" ) ) ;
      else if ( output.fail() )
        hr.SetNoThrow ( Error ( "Unexpected I/O error writing resource file" ) ) ;
      else
        hr.SetNoThrow ( Error ( "Unexpected exception writing resource file" ) ) ;
    }
  }
  catch ( _com_error e )  { hr.SetNoThrow ( Error( (LPCOLESTR)e.Description()) ) ; }
  catch ( ... )           { hr.SetNoThrow ( Error("Exception in CResFile::Export()") ) ;    }

  return hr ;
}

STDMETHODIMP CResFile::get_Strings(IResStrings **pVal)
{
  HRESULT   hr = S_OK ;
  hr = m_strings.QueryInterface ( IID_IResStrings, *pVal ) ;
  return hr ;
}

HRESULT CResFile::import_single_resource ( ifstream& input )
{
  DWORD           pos ;
  DWORD           nextpos ;
  ResourceHeader  hdr ;

  pos = input.tellg() ;

  // Turn off exception handling and check for end of file
  input.exceptions ( 0 ) ;
  if ( input.peek() == EOF )
  {
    return S_OK ;
  }
  input.exceptions ( ios_base::badbit | ios_base::failbit | ios_base::eofbit ) ;

  // Read the resource header
  input.read ( (char*)&hdr, sizeof hdr ) ;

  // We are only interested in string resources.
  if ( ( hdr.TypeNameDummy == 0xFFFF ) && ( hdr.TypeId == (WORD)RT_STRING ) )
  {
    import_string_block ( input, hdr.DataSize, hdr.LanguageId, hdr.Id ) ;
  }


  // seek past the header and data to the next block
  nextpos = pos + hdr.HeaderSize + hdr.DataSize ;
  // pad to DWORD boundary
  nextpos = (nextpos + 3) & ~3 ;
  input.seekg ( nextpos ) ;

  return S_OK;
}

HRESULT CResFile::import_string_block ( ifstream& input,
                                        DWORD     BlockSize, 
                                        WORD      LanguageId,
                                        WORD      Id )
{
  // This function may throw an exception!
  // *************************************
  HRX   hr ;

  // Get a block of memory and read the whole block
  char*     pBuf = new char[BlockSize] ;
  input.read ( pBuf, BlockSize ) ;

  wchar_t*  pwchar = (wchar_t*)pBuf ;

  for ( long lStringNum = 0 ; lStringNum < 16 ; lStringNum++ )
  {
    WORD  slen = *pwchar++ ;
    
    if ( slen > 0 )
    {
      wstring   str ;
      str.assign ( pwchar, slen ) ;
      pwchar += slen ;

      WORD ResId = ((Id - 1) << 4) + lStringNum ;

      // If necessary speed this up by creating the string object directly with
      //
      // typedef CComObject<CResString> StrObj ;
      // StrObj*            pStr = new StrObj ;
      // IResStringPtr      NewString ( pStr ) ;
      //
      // (or use CComObject::CreateInstance() if FinalConstruct is required).

      IResStringPtr   NewString ( CLSID_ResString ) ;
      hr = NewString->put_LocaleId ( LanguageId ) ;
      hr = NewString->put_ResId    ( ResId ) ;
      hr = NewString->put_Text     ( _bstr_t(str.c_str()) ) ;
      hr = m_strings->Add ( NewString ) ;
    }
  }

  delete[] pBuf ;
  
  return hr ;
}

HRESULT CResFile::export_string_block ( ofstream& output,
                                        long      BlockId )
{
  // This function may throw an exception!
  // *************************************
  HRX             hr ;
  ResourceHeader  hdr ;
  long            Index ;
  long            Size ;
  long            LangId ;
  long            Top12 ;
  BSTR            BSTRText ;

  // Extract the Language ID and top 12 bits of the resource ID.
  LangId = ( BlockId >> 16 ) & 0xFFFF ;
  Top12  = ( ( BlockId >> 4  ) & 0x0FFF ) + 1 ;

  // Determine the length of the block
  Size = 0 ;
  for ( Index = 0 ; Index < 16 ; Index++ )
    Size += m_StringBlock[Index].Length + 1 ;
  // Double it, because we are using Unicode characters
  Size *= 2 ;

  // Initialise all fields in the header.
  hdr.DataSize        = Size ;
  hdr.HeaderSize      = sizeof hdr ;
  hdr.TypeNameDummy   = 0xFFFF ;
  hdr.TypeId          = (WORD)RT_STRING ;
  hdr.IdNameDummy     = 0xFFFF ;
  hdr.Id              = Top12 ;
  hdr.DataVersion     = 0 ;
  hdr.MemoryFlags     = 0 ;
  hdr.LanguageId      = LangId ;
  hdr.Version         = 0 ;
  hdr.Characteristics = 0 ;

  // Write it to the stream
  output.write ( (const char*)&hdr, sizeof hdr ) ;

  // Now write the strings
  for ( Index = 0 ; Index < 16 ; Index++ )
  {
    output.write ( (const char*)&(m_StringBlock[Index].Length), 2 ) ;

    if ( m_StringBlock[Index].pResString != NULL )
    {
      hr = m_StringBlock[Index].pResString->get_Text ( &BSTRText ) ;
      output.write ( (const char*)BSTRText, 2 * m_StringBlock[Index].Length ) ;
      SysFreeString ( BSTRText ) ;
    }
  }

  // Align to DWORD boundary.
  // Note: Padding may also be required at the end of the file to 
  //       make the whole file length a multiple of 4. This is only
  //       effective if we actually write 0's to the stream.
  //       Seeking with output.seekp() works otherwise, but does not
  //       introduce any padding if we never write any more bytes.
  DWORD   dw0 = 0 ;
  long    pad = ( -Size ) & 3 ;
  output.write ( (const char*)&dw0, pad ) ;

  return S_OK ;
}

HRESULT CResFile::export_empty_resource ( ofstream& output )
{
  // This function may throw an exception!
  // *************************************
  HRX             hr ;        // Catch block in caller!
  ResourceHeader  hdr ;

  // Initialise all fields in the header.
  hdr.DataSize        = 0 ;
  hdr.HeaderSize      = sizeof hdr ;
  hdr.TypeNameDummy   = 0xFFFF ;
  hdr.TypeId          = 0 ;
  hdr.IdNameDummy     = 0xFFFF ;
  hdr.Id              = 0 ;
  hdr.DataVersion     = 0 ;
  hdr.MemoryFlags     = 0 ;
  hdr.LanguageId      = 0 ;
  hdr.Version         = 0 ;
  hdr.Characteristics = 0 ;

  // Write it to the stream
  output.write ( (const char*)&hdr, sizeof hdr ) ;

  return S_OK ;
}


STDMETHODIMP CResFile::SaveAsNewFile()
{
  HRX             hr ;
  ofstream        output ;

  try
  {
    // Open the file
    output.open ( (const wchar_t *)m_bstrFilename, ios::binary  ) ;
    if ( !output.is_open() )
      hr = Error ( "Failed to open resource file" ) ;

    // Enable exceptions for the output stream.
    output.exceptions ( ios_base::badbit | ios_base::failbit | ios_base::eofbit ) ;

    // Write an emtpy resource at the start of the file.
    // I havn't found a documented requirement for this, but all 
    // resource files appear to have it.
    export_empty_resource ( output ) ;

    // Common function for actually exporting the strings.
    export_string_resources ( output ) ;
  }
  catch ( long )
  {
    // This catches the HRESULT type, so check hr.
    if ( hr == S_OK )
    {
      if ( output.fail() )
        hr.SetNoThrow ( Error ( "Unexpected I/O error writing resource file" ) ) ;
      else
        hr.SetNoThrow ( Error ( "Unexpected exception writing resource file" ) ) ;
    }
  }
  catch ( _com_error e )  { hr.SetNoThrow ( Error( (LPCOLESTR)e.Description()) ) ; }
  catch ( ... )           { hr.SetNoThrow ( Error("Exception in CResFile::SaveAsNewFile()") ) ;    }

  return hr ;
}

HRESULT CResFile::transfer_single_resource ( ifstream&    input,
                                             ofstream&    output,
                                             VARIANT_BOOL PreserveStrings )
{
  DWORD pos = input.tellg() ;

  // Turn off exception handling and check for end of file
  input.exceptions ( 0 ) ;
  if ( input.peek() == EOF )
  {
    return S_OK ;
  }
  input.exceptions ( ios_base::badbit | ios_base::failbit | ios_base::eofbit ) ;

  // Read the resource header
  ResourceHeader hdr ;
  input.read ( (char*)&hdr, sizeof hdr ) ;

  // Separate handling for Strings and all other resources.
  if ( ( hdr.TypeNameDummy == 0xFFFF ) && ( hdr.TypeId == (WORD)RT_STRING ) )
  {
    // It's a string
    if ( PreserveStrings )
    {
      conditional_import_string_block ( input, hdr.DataSize, hdr.LanguageId, hdr.Id ) ;
    }

    // seek past the header and data to the next block
    DWORD nextpos = pos + hdr.HeaderSize + hdr.DataSize ;
    // pad to DWORD boundary
    nextpos = (nextpos + 3) & ~3 ;
    input.seekg ( nextpos ) ;
  }
  else
  {
    // It's some other resource.
    // Write the (full or partial) header to the output.
    output.write ( (const char *)&hdr, sizeof hdr ) ;

    // Now determine how many more bytes to copy
    DWORD RemainingBytes = hdr.DataSize + hdr.HeaderSize - sizeof hdr ;
    
    // Round up to DWORD boundary
    RemainingBytes = (RemainingBytes + 3) & ~3 ;

    // We could probably do this more efficiently
    // (on CPU time, if not on my time), but allocate
    // a memory block for the data and release it 
    // immediately.
    char*     pBuf = new char[RemainingBytes] ;
    input.read ( pBuf, RemainingBytes ) ;
    output.write ( (const char *)pBuf, RemainingBytes ) ;
    delete[] pBuf ;
  }

  return S_OK;
}

HRESULT CResFile::conditional_import_string_block ( ifstream& input,
                                                    DWORD     BlockSize, 
                                                    WORD      LanguageId,
                                                    WORD      Id )
{
  // This function may throw an exception!
  // *************************************
  HRX   hr ;

  // Get a block of memory and read the whole block
  char*     pBuf = new char[BlockSize] ;
  input.read ( pBuf, BlockSize ) ;

  wchar_t*  pwchar = (wchar_t*)pBuf ;

  for ( long lStringNum = 0 ; lStringNum < 16 ; lStringNum++ )
  {
    WORD  slen = *pwchar++ ;
    
    if ( slen > 0 )
    {
      // Determine the Resource ID
      WORD ResId = ((Id - 1) << 4) + lStringNum ;

      // Determine whether we already have a string with the
      // identical Locale AND Resource IDs
      VARIANT_BOOL vbExists ;
      hr = m_strings->Exists ( LanguageId, ResId, &vbExists ) ;

      if ( !vbExists )
      {
        // Extract the string
        wstring   str ;
        str.assign ( pwchar, slen ) ;

        // Make an object and add it to our collection.
        IResStringPtr   NewString ( CLSID_ResString ) ;
        hr = NewString->put_LocaleId ( LanguageId ) ;
        hr = NewString->put_ResId    ( ResId ) ;
        hr = NewString->put_Text     ( _bstr_t(str.c_str()) ) ;
        hr = m_strings->Add ( NewString ) ;
      }

      // Advance past the string
      pwchar += slen ;
    }
  }

  delete[] pBuf ;
  
  return hr ;
}

HRESULT CResFile::export_string_resources ( ofstream& output )
{
  // This function may throw an exception!
  // *************************************
  HRX   hr ;

  // First sort the strings
  hr = m_strings->SortById() ;

  // Now loop over the strings
  long          Index ;
  long          Count ;
  long          BlockId = -1 ;

  IResString*   pString ;
  long          CombiId ;
  long          Length ;
  long          FourBitId ;
  
  hr = m_strings->get_Count ( &Count ) ;
  for ( Index = 0 ; Index < Count ; Index++ )
  {
    // Get as IResString interface
    hr = m_strings->get_Item2 ( Index, &pString ) ;

    // Get the CombiId and length of the string
    hr = pString->get_CombinedId ( &CombiId ) ;
    hr = pString->get_Length     ( &Length ) ;

    // Does it belong in the current block?
    if ( ( CombiId & 0xFFFFFFF0 ) != BlockId )
    {
      // no
      // If BlockId is not -1, then write it to the file.
      if ( BlockId != -1 )
      {
        export_string_block ( output, BlockId ) ;
      }

      // Clear the 16 string info block
      ClearStringBlock() ;

      // Initialise the BlockId for the new block
      BlockId = CombiId & 0xFFFFFFF0 ;
    }

    // Store info about this string in the string block.
    // NOTE: Tentativly a non QI'd pointer!!!
    FourBitId = CombiId & 0x0F ;
    m_StringBlock[FourBitId].pResString = pString ;
    m_StringBlock[FourBitId].Length     = Length ;
  
    // Release the string pointer
    pString->Release() ;
  }

  // Save the final block to the file.
  if ( BlockId != -1 )
  {
    export_string_block ( output, BlockId ) ;
  }

  // Clear the dirty flag.
  hr = m_strings->put_Dirty ( VARIANT_FALSE ) ;

  return hr ;
}

STDMETHODIMP CResFile::get_Locales(VARIANT *pVal)
{
  // After defining this function, I decided it would be easier to
  // implement it in the CResStrings class, but I left this function
  // anyway. From VB, it seems more logical to have it here than in
  // the Strings member.
  return m_strings->get_Locales ( pVal ) ;
}

STDMETHODIMP CResFile::get_Dirty(VARIANT_BOOL *pVal)
{
  // The dirty flag is a cooperative effort between the CResStrings
  // class and this class. Effectivly, it is set by changes to the
  // string collection (in CResStrings), but cleared by saving the
  // file (in this class).
  // Rather than have the Strings class generate events when a change
  // occurs, the dirty flag is actually implemented in that class,
  // but as a hidden property. It is exposed to the user through 
  // this class.

  return m_strings->get_Dirty ( pVal ) ;
}

STDMETHODIMP CResFile::MakeDirty()
{
  // This method was added as an afterthought.
  // The VB client can call it to set the dirty flag after 
  // modifying a string in an existing string object.
  // This is easier than handling it in the ATL component.
  return m_strings->put_Dirty ( VARIANT_TRUE ) ;
}
