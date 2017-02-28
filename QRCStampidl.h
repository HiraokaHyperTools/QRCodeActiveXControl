

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 7.00.0555 */
/* at Tue Feb 28 10:59:12 2017
 */
/* Compiler settings for .\QRCStamp.idl:
    Oicf, W1, Zp8, env=Win32 (32b run), target_arch=X86 7.00.0555 
    protocol : dce , ms_ext, c_ext
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
/* @@MIDL_FILE_HEADING(  ) */

#pragma warning( disable: 4049 )  /* more than 64k source lines */


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 440
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __QRCStampidl_h__
#define __QRCStampidl_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef ___DQRCStamp_FWD_DEFINED__
#define ___DQRCStamp_FWD_DEFINED__
typedef interface _DQRCStamp _DQRCStamp;
#endif 	/* ___DQRCStamp_FWD_DEFINED__ */


#ifndef ___DQRCStampEvents_FWD_DEFINED__
#define ___DQRCStampEvents_FWD_DEFINED__
typedef interface _DQRCStampEvents _DQRCStampEvents;
#endif 	/* ___DQRCStampEvents_FWD_DEFINED__ */


#ifndef __QRCStamp_FWD_DEFINED__
#define __QRCStamp_FWD_DEFINED__

#ifdef __cplusplus
typedef class QRCStamp QRCStamp;
#else
typedef struct QRCStamp QRCStamp;
#endif /* __cplusplus */

#endif 	/* __QRCStamp_FWD_DEFINED__ */


#ifdef __cplusplus
extern "C"{
#endif 


/* interface __MIDL_itf_QRCStamp_0000_0000 */
/* [local] */ 

typedef 
enum tagSizeModeEnum
    {	SizeModeCenter	= 0,
	SizeModeStretch	= 1,
	SizeModeZoom	= 2
    } 	SizeModeEnum;

typedef 
enum tagQRErrorCorrectLevel
    {	ErrorCorrectLevelL	= 1,
	ErrorCorrectLevelM	= 0,
	ErrorCorrectLevelQ	= 3,
	ErrorCorrectLevelH	= 2
    } 	QRErrorCorrectLevelEnum;

typedef 
enum tagEncodeMode
    {	EncodeModeShiftJIS	= 0,
	EncodeModeAlphaNum	= ( EncodeModeShiftJIS + 1 ) ,
	EncodeModeNum	= ( EncodeModeAlphaNum + 1 ) ,
	EncodeMode8bit	= ( EncodeModeNum + 1 ) ,
	EncodeModeStrucShiftJIS	= 256,
	EncodeModeStrucAlphaNum	= ( EncodeModeStrucShiftJIS + 1 ) ,
	EncodeModeStrucNum	= ( EncodeModeStrucAlphaNum + 1 ) ,
	EncodeModeStruc8bit	= ( EncodeModeStrucNum + 1 ) ,
	EncodeModeVertStrucShiftJIS	= 512,
	EncodeModeVertStrucAlphaNum	= ( EncodeModeVertStrucShiftJIS + 1 ) ,
	EncodeModeVertStrucNum	= ( EncodeModeVertStrucAlphaNum + 1 ) ,
	EncodeModeVertStruc8bit	= ( EncodeModeVertStrucNum + 1 ) 
    } 	EncodeModeEnum;



extern RPC_IF_HANDLE __MIDL_itf_QRCStamp_0000_0000_v0_0_c_ifspec;
extern RPC_IF_HANDLE __MIDL_itf_QRCStamp_0000_0000_v0_0_s_ifspec;


#ifndef __QRCStampLib_LIBRARY_DEFINED__
#define __QRCStampLib_LIBRARY_DEFINED__

/* library QRCStampLib */
/* [control][helpstring][helpfile][version][uuid] */ 


EXTERN_C const IID LIBID_QRCStampLib;

#ifndef ___DQRCStamp_DISPINTERFACE_DEFINED__
#define ___DQRCStamp_DISPINTERFACE_DEFINED__

/* dispinterface _DQRCStamp */
/* [helpstring][uuid] */ 


EXTERN_C const IID DIID__DQRCStamp;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("D94D9B2B-F60D-498B-8F56-4D067680E08B")
    _DQRCStamp : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct _DQRCStampVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            _DQRCStamp * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            __RPC__deref_out  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            _DQRCStamp * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            _DQRCStamp * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            _DQRCStamp * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            _DQRCStamp * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            _DQRCStamp * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            _DQRCStamp * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        END_INTERFACE
    } _DQRCStampVtbl;

    interface _DQRCStamp
    {
        CONST_VTBL struct _DQRCStampVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _DQRCStamp_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define _DQRCStamp_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define _DQRCStamp_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define _DQRCStamp_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define _DQRCStamp_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define _DQRCStamp_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define _DQRCStamp_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* ___DQRCStamp_DISPINTERFACE_DEFINED__ */


#ifndef ___DQRCStampEvents_DISPINTERFACE_DEFINED__
#define ___DQRCStampEvents_DISPINTERFACE_DEFINED__

/* dispinterface _DQRCStampEvents */
/* [helpstring][uuid] */ 


EXTERN_C const IID DIID__DQRCStampEvents;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("834C0E14-6FEE-484E-A826-450701DA595B")
    _DQRCStampEvents : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct _DQRCStampEventsVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            _DQRCStampEvents * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            __RPC__deref_out  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            _DQRCStampEvents * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            _DQRCStampEvents * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            _DQRCStampEvents * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            _DQRCStampEvents * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            _DQRCStampEvents * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            _DQRCStampEvents * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        END_INTERFACE
    } _DQRCStampEventsVtbl;

    interface _DQRCStampEvents
    {
        CONST_VTBL struct _DQRCStampEventsVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _DQRCStampEvents_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define _DQRCStampEvents_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define _DQRCStampEvents_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define _DQRCStampEvents_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define _DQRCStampEvents_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define _DQRCStampEvents_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define _DQRCStampEvents_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* ___DQRCStampEvents_DISPINTERFACE_DEFINED__ */


EXTERN_C const CLSID CLSID_QRCStamp;

#ifdef __cplusplus

class DECLSPEC_UUID("D6FEB896-8CB3-4B3A-B997-4ED31B7670EC")
QRCStamp;
#endif
#endif /* __QRCStampLib_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


