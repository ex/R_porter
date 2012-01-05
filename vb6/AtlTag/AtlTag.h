/* this ALWAYS GENERATED file contains the definitions for the interfaces */


/* File created by MIDL compiler version 5.01.0164 */
/* at Sun Mar 21 10:54:53 2004
 */
/* Compiler settings for C:\Esau\Dev\R_porter\AtlTag\AtlTag.idl:
    Oicf (OptLev=i2), W1, Zp8, env=Win32, ms_ext, c_ext
    error checks: allocation ref bounds_check enum stub_data 
*/
//@@MIDL_FILE_HEADING(  )


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 440
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif // __RPCNDR_H_VERSION__

#ifndef COM_NO_WINDOWS_H
#include "windows.h"
#include "ole2.h"
#endif /*COM_NO_WINDOWS_H*/

#ifndef __AtlTag_h__
#define __AtlTag_h__

#ifdef __cplusplus
extern "C"{
#endif 

/* Forward Declarations */ 

#ifndef __IexAtlTag_FWD_DEFINED__
#define __IexAtlTag_FWD_DEFINED__
typedef interface IexAtlTag IexAtlTag;
#endif 	/* __IexAtlTag_FWD_DEFINED__ */


#ifndef __IexAtlTag2_FWD_DEFINED__
#define __IexAtlTag2_FWD_DEFINED__
typedef interface IexAtlTag2 IexAtlTag2;
#endif 	/* __IexAtlTag2_FWD_DEFINED__ */


#ifndef __exAtlTag_FWD_DEFINED__
#define __exAtlTag_FWD_DEFINED__

#ifdef __cplusplus
typedef class exAtlTag exAtlTag;
#else
typedef struct exAtlTag exAtlTag;
#endif /* __cplusplus */

#endif 	/* __exAtlTag_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

void __RPC_FAR * __RPC_USER MIDL_user_allocate(size_t);
void __RPC_USER MIDL_user_free( void __RPC_FAR * ); 

#ifndef __IexAtlTag_INTERFACE_DEFINED__
#define __IexAtlTag_INTERFACE_DEFINED__

/* interface IexAtlTag */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_IexAtlTag;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("ED693590-6D10-4270-A778-45D0094CBF6D")
    IexAtlTag : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE SetPathFile( 
            /* [in] */ BSTR pathFile,
            /* [in] */ LONG BufferLen,
            /* [in] */ BYTE bVerifyBR,
            /* [in] */ LONG cod) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Bitrate( 
            /* [retval][out] */ INT __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Mpeg( 
            /* [retval][out] */ short __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Layer( 
            /* [retval][out] */ short __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_SampleRate( 
            /* [retval][out] */ INT __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_Mode( 
            /* [retval][out] */ short __RPC_FAR *pVal) = 0;
        
        virtual /* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE get_ErrorNumber( 
            /* [retval][out] */ short __RPC_FAR *pVal) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IexAtlTagVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IexAtlTag __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IexAtlTag __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IexAtlTag __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            IexAtlTag __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            IexAtlTag __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            IexAtlTag __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            IexAtlTag __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetPathFile )( 
            IexAtlTag __RPC_FAR * This,
            /* [in] */ BSTR pathFile,
            /* [in] */ LONG BufferLen,
            /* [in] */ BYTE bVerifyBR,
            /* [in] */ LONG cod);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Bitrate )( 
            IexAtlTag __RPC_FAR * This,
            /* [retval][out] */ INT __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Mpeg )( 
            IexAtlTag __RPC_FAR * This,
            /* [retval][out] */ short __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Layer )( 
            IexAtlTag __RPC_FAR * This,
            /* [retval][out] */ short __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_SampleRate )( 
            IexAtlTag __RPC_FAR * This,
            /* [retval][out] */ INT __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Mode )( 
            IexAtlTag __RPC_FAR * This,
            /* [retval][out] */ short __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ErrorNumber )( 
            IexAtlTag __RPC_FAR * This,
            /* [retval][out] */ short __RPC_FAR *pVal);
        
        END_INTERFACE
    } IexAtlTagVtbl;

    interface IexAtlTag
    {
        CONST_VTBL struct IexAtlTagVtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IexAtlTag_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IexAtlTag_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IexAtlTag_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IexAtlTag_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IexAtlTag_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IexAtlTag_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IexAtlTag_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define IexAtlTag_SetPathFile(This,pathFile,BufferLen,bVerifyBR,cod)	\
    (This)->lpVtbl -> SetPathFile(This,pathFile,BufferLen,bVerifyBR,cod)

#define IexAtlTag_get_Bitrate(This,pVal)	\
    (This)->lpVtbl -> get_Bitrate(This,pVal)

#define IexAtlTag_get_Mpeg(This,pVal)	\
    (This)->lpVtbl -> get_Mpeg(This,pVal)

#define IexAtlTag_get_Layer(This,pVal)	\
    (This)->lpVtbl -> get_Layer(This,pVal)

#define IexAtlTag_get_SampleRate(This,pVal)	\
    (This)->lpVtbl -> get_SampleRate(This,pVal)

#define IexAtlTag_get_Mode(This,pVal)	\
    (This)->lpVtbl -> get_Mode(This,pVal)

#define IexAtlTag_get_ErrorNumber(This,pVal)	\
    (This)->lpVtbl -> get_ErrorNumber(This,pVal)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IexAtlTag_SetPathFile_Proxy( 
    IexAtlTag __RPC_FAR * This,
    /* [in] */ BSTR pathFile,
    /* [in] */ LONG BufferLen,
    /* [in] */ BYTE bVerifyBR,
    /* [in] */ LONG cod);


void __RPC_STUB IexAtlTag_SetPathFile_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IexAtlTag_get_Bitrate_Proxy( 
    IexAtlTag __RPC_FAR * This,
    /* [retval][out] */ INT __RPC_FAR *pVal);


void __RPC_STUB IexAtlTag_get_Bitrate_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IexAtlTag_get_Mpeg_Proxy( 
    IexAtlTag __RPC_FAR * This,
    /* [retval][out] */ short __RPC_FAR *pVal);


void __RPC_STUB IexAtlTag_get_Mpeg_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IexAtlTag_get_Layer_Proxy( 
    IexAtlTag __RPC_FAR * This,
    /* [retval][out] */ short __RPC_FAR *pVal);


void __RPC_STUB IexAtlTag_get_Layer_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IexAtlTag_get_SampleRate_Proxy( 
    IexAtlTag __RPC_FAR * This,
    /* [retval][out] */ INT __RPC_FAR *pVal);


void __RPC_STUB IexAtlTag_get_SampleRate_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IexAtlTag_get_Mode_Proxy( 
    IexAtlTag __RPC_FAR * This,
    /* [retval][out] */ short __RPC_FAR *pVal);


void __RPC_STUB IexAtlTag_get_Mode_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id][propget] */ HRESULT STDMETHODCALLTYPE IexAtlTag_get_ErrorNumber_Proxy( 
    IexAtlTag __RPC_FAR * This,
    /* [retval][out] */ short __RPC_FAR *pVal);


void __RPC_STUB IexAtlTag_get_ErrorNumber_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IexAtlTag_INTERFACE_DEFINED__ */


#ifndef __IexAtlTag2_INTERFACE_DEFINED__
#define __IexAtlTag2_INTERFACE_DEFINED__

/* interface IexAtlTag2 */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_IexAtlTag2;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("CDB2899F-AFDE-47f4-B152-56B3CD4C633F")
    IexAtlTag2 : public IexAtlTag
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE SetBufferLength( 
            INT BufferLen) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE SetPathFile2( 
            /* [in] */ BSTR pathFile,
            /* [in] */ BYTE bVerifyBR,
            /* [in] */ LONG cod) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE FreeBuffer( void) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct IexAtlTag2Vtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *QueryInterface )( 
            IexAtlTag2 __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void __RPC_FAR *__RPC_FAR *ppvObject);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *AddRef )( 
            IexAtlTag2 __RPC_FAR * This);
        
        ULONG ( STDMETHODCALLTYPE __RPC_FAR *Release )( 
            IexAtlTag2 __RPC_FAR * This);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfoCount )( 
            IexAtlTag2 __RPC_FAR * This,
            /* [out] */ UINT __RPC_FAR *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetTypeInfo )( 
            IexAtlTag2 __RPC_FAR * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo __RPC_FAR *__RPC_FAR *ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE __RPC_FAR *GetIDsOfNames )( 
            IexAtlTag2 __RPC_FAR * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR __RPC_FAR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID __RPC_FAR *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *Invoke )( 
            IexAtlTag2 __RPC_FAR * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS __RPC_FAR *pDispParams,
            /* [out] */ VARIANT __RPC_FAR *pVarResult,
            /* [out] */ EXCEPINFO __RPC_FAR *pExcepInfo,
            /* [out] */ UINT __RPC_FAR *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetPathFile )( 
            IexAtlTag2 __RPC_FAR * This,
            /* [in] */ BSTR pathFile,
            /* [in] */ LONG BufferLen,
            /* [in] */ BYTE bVerifyBR,
            /* [in] */ LONG cod);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Bitrate )( 
            IexAtlTag2 __RPC_FAR * This,
            /* [retval][out] */ INT __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Mpeg )( 
            IexAtlTag2 __RPC_FAR * This,
            /* [retval][out] */ short __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Layer )( 
            IexAtlTag2 __RPC_FAR * This,
            /* [retval][out] */ short __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_SampleRate )( 
            IexAtlTag2 __RPC_FAR * This,
            /* [retval][out] */ INT __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_Mode )( 
            IexAtlTag2 __RPC_FAR * This,
            /* [retval][out] */ short __RPC_FAR *pVal);
        
        /* [helpstring][id][propget] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *get_ErrorNumber )( 
            IexAtlTag2 __RPC_FAR * This,
            /* [retval][out] */ short __RPC_FAR *pVal);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetBufferLength )( 
            IexAtlTag2 __RPC_FAR * This,
            INT BufferLen);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *SetPathFile2 )( 
            IexAtlTag2 __RPC_FAR * This,
            /* [in] */ BSTR pathFile,
            /* [in] */ BYTE bVerifyBR,
            /* [in] */ LONG cod);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE __RPC_FAR *FreeBuffer )( 
            IexAtlTag2 __RPC_FAR * This);
        
        END_INTERFACE
    } IexAtlTag2Vtbl;

    interface IexAtlTag2
    {
        CONST_VTBL struct IexAtlTag2Vtbl __RPC_FAR *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IexAtlTag2_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define IexAtlTag2_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define IexAtlTag2_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define IexAtlTag2_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define IexAtlTag2_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define IexAtlTag2_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define IexAtlTag2_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)


#define IexAtlTag2_SetPathFile(This,pathFile,BufferLen,bVerifyBR,cod)	\
    (This)->lpVtbl -> SetPathFile(This,pathFile,BufferLen,bVerifyBR,cod)

#define IexAtlTag2_get_Bitrate(This,pVal)	\
    (This)->lpVtbl -> get_Bitrate(This,pVal)

#define IexAtlTag2_get_Mpeg(This,pVal)	\
    (This)->lpVtbl -> get_Mpeg(This,pVal)

#define IexAtlTag2_get_Layer(This,pVal)	\
    (This)->lpVtbl -> get_Layer(This,pVal)

#define IexAtlTag2_get_SampleRate(This,pVal)	\
    (This)->lpVtbl -> get_SampleRate(This,pVal)

#define IexAtlTag2_get_Mode(This,pVal)	\
    (This)->lpVtbl -> get_Mode(This,pVal)

#define IexAtlTag2_get_ErrorNumber(This,pVal)	\
    (This)->lpVtbl -> get_ErrorNumber(This,pVal)


#define IexAtlTag2_SetBufferLength(This,BufferLen)	\
    (This)->lpVtbl -> SetBufferLength(This,BufferLen)

#define IexAtlTag2_SetPathFile2(This,pathFile,bVerifyBR,cod)	\
    (This)->lpVtbl -> SetPathFile2(This,pathFile,bVerifyBR,cod)

#define IexAtlTag2_FreeBuffer(This)	\
    (This)->lpVtbl -> FreeBuffer(This)

#endif /* COBJMACROS */


#endif 	/* C style interface */



/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IexAtlTag2_SetBufferLength_Proxy( 
    IexAtlTag2 __RPC_FAR * This,
    INT BufferLen);


void __RPC_STUB IexAtlTag2_SetBufferLength_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IexAtlTag2_SetPathFile2_Proxy( 
    IexAtlTag2 __RPC_FAR * This,
    /* [in] */ BSTR pathFile,
    /* [in] */ BYTE bVerifyBR,
    /* [in] */ LONG cod);


void __RPC_STUB IexAtlTag2_SetPathFile2_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);


/* [helpstring][id] */ HRESULT STDMETHODCALLTYPE IexAtlTag2_FreeBuffer_Proxy( 
    IexAtlTag2 __RPC_FAR * This);


void __RPC_STUB IexAtlTag2_FreeBuffer_Stub(
    IRpcStubBuffer *This,
    IRpcChannelBuffer *_pRpcChannelBuffer,
    PRPC_MESSAGE _pRpcMessage,
    DWORD *_pdwStubPhase);



#endif 	/* __IexAtlTag2_INTERFACE_DEFINED__ */



#ifndef __ATLTAGLib_LIBRARY_DEFINED__
#define __ATLTAGLib_LIBRARY_DEFINED__

/* library ATLTAGLib */
/* [helpstring][version][uuid] */ 


EXTERN_C const IID LIBID_ATLTAGLib;

EXTERN_C const CLSID CLSID_exAtlTag;

#ifdef __cplusplus

class DECLSPEC_UUID("D020FF56-24F5-4F5E-A776-7ECC9CB872ED")
exAtlTag;
#endif
#endif /* __ATLTAGLib_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

unsigned long             __RPC_USER  BSTR_UserSize(     unsigned long __RPC_FAR *, unsigned long            , BSTR __RPC_FAR * ); 
unsigned char __RPC_FAR * __RPC_USER  BSTR_UserMarshal(  unsigned long __RPC_FAR *, unsigned char __RPC_FAR *, BSTR __RPC_FAR * ); 
unsigned char __RPC_FAR * __RPC_USER  BSTR_UserUnmarshal(unsigned long __RPC_FAR *, unsigned char __RPC_FAR *, BSTR __RPC_FAR * ); 
void                      __RPC_USER  BSTR_UserFree(     unsigned long __RPC_FAR *, BSTR __RPC_FAR * ); 

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif
