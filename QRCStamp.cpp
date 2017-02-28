// QRCStamp.cpp : CQRCStampApp および DLL 登録の実装

#include "stdafx.h"
#include "QRCStamp.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


CQRCStampApp theApp;

const GUID CDECL BASED_CODE _tlid =
		{ 0x60028474, 0x1671, 0x4F2D, { 0x94, 0xCF, 0x86, 0x63, 0x87, 0x96, 0x99, 0x7F } };

// QRCStamp.idl の version(x.y) も変更する。
const WORD _wVerMajor = 1;
const WORD _wVerMinor = 3;



// CQRCStampApp::InitInstance - DLL 初期化

BOOL CQRCStampApp::InitInstance()
{
	BOOL bInit = COleControlModule::InitInstance();

	if (bInit)
	{
		// TODO: この位置にモジュールの初期化処理コードを追加してください。
	}

	return bInit;
}



// CQRCStampApp::ExitInstance - DLL 終了

int CQRCStampApp::ExitInstance()
{
	// TODO: この位置にモジュールの終了処理を追加してください。

	return COleControlModule::ExitInstance();
}



// DllRegisterServer - エントリをシステム レジストリに追加します。

STDAPI DllRegisterServer(void)
{
	AFX_MANAGE_STATE(_afxModuleAddrThis);

	if (!AfxOleRegisterTypeLib(AfxGetInstanceHandle(), _tlid))
		return ResultFromScode(SELFREG_E_TYPELIB);

	if (!COleObjectFactoryEx::UpdateRegistryAll(TRUE))
		return ResultFromScode(SELFREG_E_CLASS);

	return NOERROR;
}



// DllUnregisterServer - エントリをレジストリから削除します。

STDAPI DllUnregisterServer(void)
{
	AFX_MANAGE_STATE(_afxModuleAddrThis);

	if (!AfxOleUnregisterTypeLib(_tlid, _wVerMajor, _wVerMinor))
		return ResultFromScode(SELFREG_E_TYPELIB);

	if (!COleObjectFactoryEx::UpdateRegistryAll(FALSE))
		return ResultFromScode(SELFREG_E_CLASS);

	return NOERROR;
}
