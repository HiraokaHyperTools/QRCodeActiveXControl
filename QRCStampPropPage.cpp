// QRCStampPropPage.cpp : CQRCStampPropPage プロパティ ページ クラスの実装

#include "stdafx.h"
#include "QRCStamp.h"
#include "QRCStampPropPage.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


IMPLEMENT_DYNCREATE(CQRCStampPropPage, COlePropertyPage)



// メッセージ マップ

BEGIN_MESSAGE_MAP(CQRCStampPropPage, COlePropertyPage)
END_MESSAGE_MAP()



// クラス ファクトリおよび GUID を初期化します。

IMPLEMENT_OLECREATE_EX(CQRCStampPropPage, "QRCSTAMP.QRCStampPropPage.1",
	0x29f3d6f9, 0x9312, 0x46c8, 0xb6, 0x6f, 0x36, 0x87, 0x98, 0x49, 0x8d, 0x51)



// CQRCStampPropPage::CQRCStampPropPageFactory::UpdateRegistry -
// CQRCStampPropPage のシステム レジストリ エントリを追加または削除します。

BOOL CQRCStampPropPage::CQRCStampPropPageFactory::UpdateRegistry(BOOL bRegister)
{
	if (bRegister)
		return AfxOleRegisterPropertyPageClass(AfxGetInstanceHandle(),
			m_clsid, IDS_QRCSTAMP_PPG);
	else
		return AfxOleUnregisterClass(m_clsid, NULL);
}



// CQRCStampPropPage::CQRCStampPropPage - コンストラクタ

CQRCStampPropPage::CQRCStampPropPage() :
	COlePropertyPage(IDD, IDS_QRCSTAMP_PPG_CAPTION)
		, m_strValue(_T(""))
		, m_dblPPM(0)
		, m_iSizeMode(0)
		, m_iECL(0)
	{
}



// CQRCStampPropPage::DoDataExchange - ページおよびプロパティ間でデータを移動します。

void CQRCStampPropPage::DoDataExchange(CDataExchange* pDX)
{
	DDP_Text(pDX, IDC_EDIT_VALUE, m_strValue, _T("Value"));
	DDX_Text(pDX, IDC_EDIT_VALUE, m_strValue);

	DDP_Text(pDX, IDC_EDIT_PPM, m_dblPPM, _T("PelsPerMeter"));
	DDX_Text(pDX, IDC_EDIT_PPM, m_dblPPM);

	DDP_CBIndex(pDX, IDC_COMBO_SIZEMODE, m_iSizeMode, _T("SizeMode"));
	DDX_CBIndex(pDX, IDC_COMBO_SIZEMODE, m_iSizeMode);

	DDP_CBIndex(pDX, IDC_COMBO_ECL, m_iECL, _T("ErrorCorrectLevel"));
	DDX_CBIndex(pDX, IDC_COMBO_ECL, m_iECL);

	DDP_PostProcessing(pDX);
}



// CQRCStampPropPage メッセージ ハンドラ
