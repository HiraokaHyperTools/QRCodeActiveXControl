#pragma once

// QRCStamp.h : QRCStamp.DLL のメイン ヘッダー ファイル

#if !defined( __AFXCTL_H__ )
#error "このファイルをインクルードする前に 'afxctl.h' をインクルードしてください。"
#endif

#include "resource.h"       // メイン シンボル


// CQRCStampApp : 実装に関しては QRCStamp.cpp を参照してください。

class CQRCStampApp : public COleControlModule
{
public:
	BOOL InitInstance();
	int ExitInstance();
};

extern const GUID CDECL _tlid;
extern const WORD _wVerMajor;
extern const WORD _wVerMinor;

