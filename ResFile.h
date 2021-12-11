// ResFile.h : Declaration of the CResFile

#ifndef __RESFILE_H_
#define __RESFILE_H_

#include "resource.h"       // main symbols

//
// struct ResourceHeader
//
// The resourse header cannot always be defined as a C structure,
// because two of the fields may have a variable length.
// HOWEVER, 
// - for the special case of string resources, these fields have 
//   a fixed length and
// - for other resource types, we only require the first two fields
//   which are always at fixed positions.
//

struct ResourceHeader
{
   DWORD     DataSize;           // Size of data without header
   DWORD     HeaderSize;         // Length of the additional header
   //[Ordinal or name TYPE]      // Type identifier, id or string
   WORD      TypeNameDummy ;     // For String resourses 0xFFFF
   WORD      TypeId ;
   //[Ordinal or name NAME]      // Name identifier, id or string
   WORD      IdNameDummy ;       // For String resourses 0xFFFF
   WORD      Id ;
   DWORD     DataVersion;        // Predefined resource data version
   WORD      MemoryFlags;        // State of the resource
   WORD      LanguageId;         // Unicode support for NLS
   DWORD     Version;            // Version of the resource data
   DWORD     Characteristics;    // Characteristics of the data
} ;

//
// Private structures used on exporting a block of 16 strings
//
struct StringDescriptor
{
  IResString*   pResString ;
  WORD          Length ;
} ;

/////////////////////////////////////////////////////////////////////////////
// CResFile
class ATL_NO_VTABLE CResFile : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CResFile, &CLSID_ResFile>,
	public ISupportErrorInfo,
	public IDispatchImpl<IResFile, &IID_IResFile, &LIBID_MLRESUTILLib>
{
public:
	CResFile() :
    m_strings ( CLSID_ResStrings ) 
	{
	}

DECLARE_REGISTRY_RESOURCEID(IDR_RESFILE)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CResFile)
	COM_INTERFACE_ENTRY(IResFile)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
END_COM_MAP()

// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

// IResFile
public:
	STDMETHOD(MakeDirty)();
	STDMETHOD(get_Dirty)(/*[out, retval]*/ VARIANT_BOOL *pVal);
	STDMETHOD(get_Locales)(/*[out, retval]*/ VARIANT *pVal);
	STDMETHOD(SaveAsNewFile)();
	STDMETHOD(get_Strings)(/*[out, retval]*/ IResStrings* *pVal);
	STDMETHOD(Export)(/*[in, optional, defaultvalue(FALSE)]*/ VARIANT_BOOL PreserveStrings);
	STDMETHOD(Import)();
	STDMETHOD(get_Filename)(/*[out, retval]*/ BSTR *pVal);
	STDMETHOD(put_Filename)(/*[in]*/ BSTR newVal);

private:

    HRESULT import_single_resource   ( ifstream&    input ) ;

    HRESULT import_string_block      ( ifstream&    input, 
                                       DWORD        BlockSize, 
                                       WORD         LanguageId,
                                       WORD         Id ) ;

    HRESULT conditional_import_string_block      
                                     ( ifstream&    input, 
                                       DWORD        BlockSize, 
                                       WORD         LanguageId,
                                       WORD         Id ) ;

    HRESULT export_string_resources  ( ofstream&    output ) ;

    HRESULT export_string_block      ( ofstream&    output,
                                       long         BlockId ) ;

    HRESULT export_empty_resource    ( ofstream&    output ) ;

    HRESULT transfer_single_resource ( ifstream&    input,
                                       ofstream&    output,
                                       VARIANT_BOOL PreserveStrings ) ;

    _bstr_t           m_bstrFilename ;
    IResStringsPtr    m_strings ;

    void ClearStringBlock()
    {
      for ( int i = 0 ; i < 16 ; i++ )
      {
        m_StringBlock[i].Length     = 0 ;
        m_StringBlock[i].pResString = 0 ;
      }
    }

    // Info about 16 strings to export in a block.
    StringDescriptor  m_StringBlock[16] ;

};

#endif //__RESFILE_H_
