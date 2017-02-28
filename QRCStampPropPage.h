#pragma once

// QRCStampPropPage.h : CQRCStampPropPage プロパティ ページ クラスの宣言です。


// CQRCStampPropPage : 実装に関しては QRCStampPropPage.cpp を参照してください。

class CQRCStampPropPage : public COlePropertyPage
{
	DECLARE_DYNCREATE(CQRCStampPropPage)
	DECLARE_OLECREATE_EX(CQRCStampPropPage)

// コンストラクタ
public:
	CQRCStampPropPage();

// ダイアログ データ
	enum { IDD = IDD_PROPPAGE_QRCSTAMP };

// 実装
protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV サポート

// メッセージ マップ
protected:
	DECLARE_MESSAGE_MAP()
public:
	CString m_strValue;
	double m_dblPPM;
	int m_iSizeMode;
	int m_iECL;
};

