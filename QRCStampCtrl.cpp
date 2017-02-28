// QRCStampCtrl.cpp :  CQRCStampCtrl ActiveX コントロール クラスの実装

#include "stdafx.h"
#include "QRCStamp.h"
#include "QRCStampCtrl.h"
#include "QRCStampPropPage.h"

#include <qrencode.h>

#pragma comment(lib, "version.lib")

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


IMPLEMENT_DYNCREATE(CQRCStampCtrl, COleControl)



// メッセージ マップ

BEGIN_MESSAGE_MAP(CQRCStampCtrl, COleControl)
	ON_OLEVERB(AFX_IDS_VERB_PROPERTIES, OnProperties)
END_MESSAGE_MAP()



// ディスパッチ マップ

BEGIN_DISPATCH_MAP(CQRCStampCtrl, COleControl)
	DISP_FUNCTION_ID(CQRCStampCtrl, "AboutBox", DISPID_ABOUTBOX, AboutBox, VT_EMPTY, VTS_NONE)

	DISP_PROPERTY_EX_ID(CQRCStampCtrl, "Value", dispidValue, GetValue, SetValue, VT_VARIANT)
	DISP_PROPERTY_EX_ID(CQRCStampCtrl, "PelsPerMeter", dispidPelsPerMeter, GetPelsPerMeter, SetPelsPerMeter, VT_R8)
	DISP_PROPERTY_EX_ID(CQRCStampCtrl, "Picture", dispidPicture, GetPicture, SetPicture, VT_PICTURE)
	DISP_PROPERTY_EX_ID(CQRCStampCtrl, "SizeMode", dispidSizeMode, GetSizeMode, SetSizeMode, VT_I4)
	DISP_PROPERTY_EX_ID(CQRCStampCtrl, "ErrorCorrectLevel", dispidErrorCorrectLevel, GetErrorCorrectLevel, SetErrorCorrectLevel, VT_I4)
	DISP_PROPERTY_EX_ID(CQRCStampCtrl, "EncodeMode", dispidEncodeMode, GetEncodeMode, SetEncodeMode, VT_I4)
	DISP_PROPERTY_EX_ID(CQRCStampCtrl, "PixScale", dispidPixScale, GetPixScale, SetPixScale, VT_I4)
	DISP_PROPERTY_EX_ID(CQRCStampCtrl, "QRVersion", dispidQRVersion, GetQRVersion, SetQRVersion, VT_I4)
	DISP_PROPERTY_EX_ID(CQRCStampCtrl, "HasErrors", dispidHasErrors, GetHasErrors, SetHasErrors, VT_BOOL)
	DISP_PROPERTY_EX_ID(CQRCStampCtrl, "ErrorMessage", dispidErrorMessage, GetErrorMessage, SetErrorMessage, VT_BSTR)

	DISP_STOCKPROP_BACKCOLOR()
	DISP_STOCKPROP_FORECOLOR()
	DISP_STOCKPROP_FONT()
END_DISPATCH_MAP()



// イベント マップ

BEGIN_EVENT_MAP(CQRCStampCtrl, COleControl)
	EVENT_CUSTOM_ID("ValueChange", eventidValueChange, ValueChange, VTS_UNKNOWN)
	EVENT_CUSTOM_ID("PelsPerMeterChange", eventidPelsPerMeterChange, PelsPerMeterChange, VTS_R8)
	EVENT_CUSTOM_ID("PictureChange", eventidPictureChange, PictureChange, VTS_PICTURE)
	EVENT_CUSTOM_ID("SizeModeChange", eventidSizeModeChange, SizeModeChange, VTS_NONE)
	EVENT_CUSTOM_ID("ErrorCorrectLevelChange", eventidErrorCorrectLevelChange, ErrorCorrectLevelChange, VTS_NONE)
	EVENT_CUSTOM_ID("EncodeModeChange", eventidEncodeModeChange, EncodeModeChange, VTS_NONE)
	EVENT_CUSTOM_ID("PixScaleChange", eventidPixScaleChange, PixScaleChange, VTS_NONE)
	EVENT_CUSTOM_ID("QRVersionChange", eventidQRVersionChange, QRVersionChange, VTS_NONE)
END_EVENT_MAP()



// プロパティ ページ

// TODO: プロパティ ページを追加して、BEGIN 行の最後にあるカウントを増やしてください。
BEGIN_PROPPAGEIDS(CQRCStampCtrl, 2)
	PROPPAGEID(CQRCStampPropPage::guid)
	PROPPAGEID(CLSID_CColorPropPage)
END_PROPPAGEIDS(CQRCStampCtrl)



// クラス ファクトリおよび GUID を初期化します。

IMPLEMENT_OLECREATE_EX(CQRCStampCtrl, "QRCSTAMP.QRCStampCtrl.1",
	0xd6feb896, 0x8cb3, 0x4b3a, 0xb9, 0x97, 0x4e, 0xd3, 0x1b, 0x76, 0x70, 0xec)



// タイプ ライブラリ ID およびバージョン

IMPLEMENT_OLETYPELIB(CQRCStampCtrl, _tlid, _wVerMajor, _wVerMinor)



// インターフェイス ID

const IID BASED_CODE IID_DQRCStamp =
		{ 0xD94D9B2B, 0xF60D, 0x498B, { 0x8F, 0x56, 0x4D, 0x6, 0x76, 0x80, 0xE0, 0x8B } };
const IID BASED_CODE IID_DQRCStampEvents =
		{ 0x834C0E14, 0x6FEE, 0x484E, { 0xA8, 0x26, 0x45, 0x7, 0x1, 0xDA, 0x59, 0x5B } };



// コントロールの種類の情報

static const DWORD BASED_CODE _dwQRCStampOleMisc =
	OLEMISC_ACTIVATEWHENVISIBLE |
	OLEMISC_SETCLIENTSITEFIRST |
	OLEMISC_INSIDEOUT |
	OLEMISC_CANTLINKINSIDE |
	OLEMISC_RECOMPOSEONRESIZE;

IMPLEMENT_OLECTLTYPE(CQRCStampCtrl, IDS_QRCSTAMP, _dwQRCStampOleMisc)



// CQRCStampCtrl::CQRCStampCtrlFactory::UpdateRegistry -
// CQRCStampCtrl のシステム レジストリ エントリを追加または削除します。

BOOL CQRCStampCtrl::CQRCStampCtrlFactory::UpdateRegistry(BOOL bRegister)
{
	// TODO: コントロールが apartment モデルのスレッド処理の規則に従っていることを
	// 確認してください。詳細については MFC のテクニカル ノート 64 を参照してください。
	// アパートメント モデルのスレッド処理の規則に従わないコントロールの場合は、6 番目
	// のパラメータを以下のように変更してください。
	// afxRegApartmentThreading を 0 に設定します。

	if (bRegister)
		return AfxOleRegisterControlClass(
			AfxGetInstanceHandle(),
			m_clsid,
			m_lpszProgID,
			IDS_QRCSTAMP,
			IDB_QRCSTAMP,
			afxRegApartmentThreading,
			_dwQRCStampOleMisc,
			_tlid,
			_wVerMajor,
			_wVerMinor);
	else
		return AfxOleUnregisterClass(m_clsid, m_lpszProgID);
}



// CQRCStampCtrl::CQRCStampCtrl - コンストラクタ

CQRCStampCtrl::CQRCStampCtrl()
{
	InitializeIIDs(&IID_DQRCStamp, &IID_DQRCStampEvents);
}



// CQRCStampCtrl::~CQRCStampCtrl - デストラクタ

CQRCStampCtrl::~CQRCStampCtrl()
{
	// TODO: この位置にコントロールのインスタンス データの後処理用のコードを追加してください
}



// CQRCStampCtrl::OnDraw - 描画関数

void CQRCStampCtrl::OnDraw(
			CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid)
{
	if (!pdc)
		return;

	CDC &dc = *pdc;

	CComQIPtr<IPicture> pict = m_Picture;

	CRect rc = rcBounds;

	CBrush br;
	COLORREF backclr = TranslateColor(GetBackColor());
	COLORREF foreclr = TranslateColor(GetForeColor());
	br.CreateSolidBrush(backclr);
	dc.FillRect(rcBounds, &br);

	if (pict == NULL) {
		rc.DeflateRect(2, 2);
		CFont *pfont = SelectStockFont(pdc);
		COLORREF priorFore = dc.SetTextColor(foreclr);
		COLORREF priorBack = dc.SetBkColor(backclr);
		dc.DrawText(_T("QRCodeActiveXControl (libqrencode)"), rc, DT_LEFT|DT_TOP);
		dc.SetBkColor(priorBack);
		dc.SetTextColor(priorFore);
		pdc->SelectObject(pfont);
	}
	else {
		OLE_XSIZE_HIMETRIC cxSrc;
		OLE_YSIZE_HIMETRIC cySrc;

		pict->get_Width(&cxSrc);
		pict->get_Height(&cySrc);

		int xDst = rc.left;
		int yDst = rc.top;
		int cxDst = rc.right - rc.left;
		int cyDst = rc.bottom - rc.top;

		switch (m_sizeMode) {
			case SizeModeCenter:
				{
					SIZEL sizeHiM = {
						cxSrc,
						cySrc
					};
					SIZEL sizePix = sizeHiM;
					switch (dc.GetMapMode()) {
						case MM_ANISOTROPIC:
						case MM_ISOTROPIC:
							sizePix.cx = LONG(sizePix.cx / (float)m_cxExtent * m_rcBounds.Width());
							sizePix.cy = LONG(sizePix.cy / (float)m_cyExtent * m_rcBounds.Height());
							break;
						default:
							dc.HIMETRICtoDP(&sizePix);
							break;
					}
					xDst += (cxDst - sizePix.cx) / 2;
					yDst += (cyDst - sizePix.cy) / 2;
					cxDst = sizePix.cx;
					cyDst = sizePix.cy;
					break;
				}
			case SizeModeZoom:
				if (cyDst != 0 && cySrc != 0) {
					float fVwpt = cxDst / (float)cyDst;
					float fPic = cxSrc / (float)cySrc;
					if (fVwpt >= fPic) {
						// ビューポートが横長
						int c = int(cxDst / fVwpt * fPic);
						xDst += (cxDst - c) / 2;
						cxDst = c;
					}
					else {
						// ビューポートが縦長
						int c = int(cyDst / fPic * fVwpt);
						yDst += (cyDst - c) / 2;
						cyDst = c;
					}
					break;
				}
		}

		int state = dc.SaveDC();
		dc.IntersectClipRect(rcBounds);

		CRect rcDst(xDst, yDst, xDst +yDst, yDst +cyDst);
		dc.LPtoDP(&rcDst);

		pict->Render(
			dc,
			xDst, yDst, cxDst, cyDst,
			0, cySrc -1,
			cxSrc, -cySrc,
			&rc
			);

		dc.RestoreDC(state);
	}
}



// CQRCStampCtrl::DoPropExchange - 永続性のサポート

#define DEFAULT_PIX_SCALE 10

void CQRCStampCtrl::DoPropExchange(CPropExchange* pPX)
{
	ExchangeVersion(pPX, MAKELONG(_wVerMinor, _wVerMajor));
	COleControl::DoPropExchange(pPX);

	PX_Double(pPX, _T("PelsPerMeter"), m_PelsPerMeter, 1000);
	PX_Long(pPX, _T("ErrorCorrectLevel"), m_ErrorCorrectLevel, ErrorCorrectLevelM);
	PX_Long(pPX, _T("EncodeMode"), m_EncodeMode, EncodeModeShiftJIS);
	PX_Long(pPX, _T("SizeMode"), m_sizeMode, SizeModeZoom);
	PX_String(pPX, _T("Value"), m_strValue);

	if (pPX->GetVersion() >= MAKELONG(1, 1)) {
		PX_Long(pPX, _T("PixScale"), m_PixScale, DEFAULT_PIX_SCALE);
	}
	else {
		if (pPX->IsLoading()) {
			m_PixScale = DEFAULT_PIX_SCALE;
		}
	}

	if (pPX->GetVersion() >= MAKELONG(2, 1)) {
		PX_Long(pPX, _T("Version"), m_QRVersion, 0);
	}
	else {
		if (pPX->IsLoading()) {
			m_QRVersion = 0;
		}
	}

	if (pPX->IsLoading()) {
		UpdateBarcode();
	}
}



// CQRCStampCtrl::GetControlFlags -
// MFC の ActiveX コントロールの実装のカスタマイズ用フラグです。
//
DWORD CQRCStampCtrl::GetControlFlags()
{
	DWORD dwFlags = COleControl::GetControlFlags();


	// コントロールはウィンドウを作成せずにアクティベート可能です。
	// TODO: コントロールのメッセージ ハンドラを作成する場合、m_hWnd
	//		m_hWnd メンバ変数の値が NULL 以外であることを最初に確認
	//		してから使用してください。
	dwFlags |= windowlessActivate;
	return dwFlags;
}



// CQRCStampCtrl::OnResetState - コントロールを既定の状態にリセットします。

void CQRCStampCtrl::OnResetState()
{
	COleControl::OnResetState();  // DoPropExchange を呼び出して既定値にリセット
	ResetStockProps();

	m_Picture.Release();
	m_PelsPerMeter = 1000;
	m_strValue.Empty();
	m_sizeMode = SizeModeZoom;
	m_ErrorCorrectLevel = ErrorCorrectLevelM;
	m_EncodeMode = EncodeModeShiftJIS;
}



// CQRCStampCtrl::AboutBox - "バージョン情報" ボックスをユーザーに表示します。

class CAboutDlg : public CDialog {
public:
	CAboutDlg(): CDialog(IDD_ABOUTBOX_QRCSTAMP) {
	}

	BOOL OnInitDialog() {
		CWnd *pWnd = GetDlgItem(IDC_STATIC_VER);
		if (pWnd != NULL) {
			HMODULE hMod = AfxGetResourceHandle();
			HRSRC hVeri = FindResource(hMod, MAKEINTRESOURCE(VS_VERSION_INFO), RT_VERSION);
			PVOID pvVeri = LockResource(LoadResource(hMod, hVeri));
			DWORD cbVeri = SizeofResource(hMod, hVeri);
			if (pvVeri != NULL && cbVeri != 0) {
				VS_FIXEDFILEINFO *pffi = NULL;
				UINT cb = 0;
				if (VerQueryValue(pvVeri, _T("\\"), reinterpret_cast<LPVOID *>(&pffi), &cb)) {
					if (pffi->dwSignature == 0xFEEF04BD) {
						CString strText;
						pWnd->GetWindowText(strText);
						CString strVer;
						strVer.Format(_T("%u.%u.%u")
							, 0U +(WORD)(pffi->dwFileVersionMS >> 16)
							, 0U +(WORD)(pffi->dwFileVersionMS >> 0)
							, 0U +(WORD)(pffi->dwFileVersionLS >> 16)
							, 0U +(WORD)(pffi->dwFileVersionLS >> 0)
							);
						strText.Replace(_T("x.x.x"), strVer);
						pWnd->SetWindowText(strText);
					}
				}
			}
		}
		return true;
	}
};

void CQRCStampCtrl::AboutBox()
{
	CAboutDlg wndDlg;
	wndDlg.DoModal();
}


// CQRCStampCtrl メッセージ ハンドラ

VARIANT CQRCStampCtrl::GetValue(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	VARIANT vaResult;
	VariantInit(&vaResult);

	vaResult.vt = VT_BSTR;
	vaResult.bstrVal = m_strValue.AllocSysString();

	return vaResult;
}

void CQRCStampCtrl::SetValue(VARIANT *newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	COleVariant vt = *newVal;
	vt.ChangeType(VT_BSTR);

	m_strValue = vt.bstrVal;

	ValueChange();

	UpdateBarcode();

	SetModifiedFlag();
}

template<typename T>
void ReleaseIt(T *pInstance) {
	throw std::exception();
}

void ReleaseIt(QRcode *pInstance) {
	QRcode_free(pInstance);
}

void ReleaseIt(QRcode_List *pInstance) {
	QRcode_List_free(pInstance);
}

template<typename T>
class AutoRelease {
	T *pInstance;
public:
	AutoRelease(T *pInstance)
		: pInstance(pInstance) {

	}

	~AutoRelease() {
		ReleaseIt(pInstance);
	}
};

void CQRCStampCtrl::UpdateBarcode() {
	HRESULT hr;

	if (m_strValue.IsEmpty()) {
		SetEmptyPicture();
		SetErrorInfo();
		return;
	}

	CStringA str = CW2A(CT2W(m_strValue), 932);

	QRecLevel level;
	switch (m_ErrorCorrectLevel) {
		default:
		case ErrorCorrectLevelL: level = QR_ECLEVEL_L; break;
		case ErrorCorrectLevelM: level = QR_ECLEVEL_M; break;
		case ErrorCorrectLevelQ: level = QR_ECLEVEL_Q; break;
		case ErrorCorrectLevelH: level = QR_ECLEVEL_H; break;
	}

	int qrVersion = std::min<int>(QRSPEC_VERSION_MAX, std::max<int>(0, m_QRVersion));

	QRcode *pCode = NULL;
	QRcode_List *pCodeListOrg = NULL;
	switch (m_EncodeMode) {
		default:
		case EncodeModeShiftJIS: // 有効
		case EncodeModeAlphaNum: // ↓無効
		case EncodeModeNum: // ↓無効
		case EncodeMode8bit: // ここに集約
			{
				QRencodeMode mode = (m_EncodeMode == EncodeModeShiftJIS)
					? QR_MODE_KANJI
					: QR_MODE_8;
				while (true) {
					pCode = QRcode_encodeString(str, qrVersion, level, mode, 1);
					if (pCode != NULL) {
						break;
					}
					if (qrVersion == QRSPEC_VERSION_MAX) {
						SetErrorInfo(true, ObtainErrorMessage(IDS_LIMIT_OVER), IDB_LIMIT_OVER);
						return;
					}
					qrVersion++;
				}
				break;
			}

			pCodeListOrg = QRcode_encodeStringStructured(str, m_QRVersion, level, QR_MODE_KANJI, 1);
			break;

		case EncodeModeStrucShiftJIS: // 有効
		case EncodeModeVertStrucShiftJIS: // 有効
		case EncodeModeStrucAlphaNum: // ↓ 無効
		case EncodeModeStrucNum: // ↓ 無効
		case EncodeModeStruc8bit: // ここに集約
		case EncodeModeVertStrucAlphaNum: // ↓無効
		case EncodeModeVertStrucNum: // ↓無効
		case EncodeModeVertStruc8bit: // ここに集約
			{
				QRencodeMode mode = (m_EncodeMode == EncodeModeStrucShiftJIS || m_EncodeMode == EncodeModeVertStrucShiftJIS)
					? QR_MODE_KANJI
					: QR_MODE_8;
				if (qrVersion == 0) {
					qrVersion = 10;
				}
				while (true) {
					pCodeListOrg = QRcode_encodeStringStructured(str, qrVersion, level, mode, 1);
					if (pCodeListOrg != NULL) {
						break;
					}
					if (qrVersion == QRSPEC_VERSION_MAX) {
						SetErrorInfo(true, ObtainErrorMessage(IDS_LIMIT_OVER), IDB_LIMIT_OVER);
						return;
					}
					qrVersion++;
				}
				break;
			}
	}

	AutoRelease<QRcode> ar1(pCode);
	AutoRelease<QRcode_List> ar2(pCodeListOrg);

	const int pixScale = std::max<int>(1, std::min<int>(20, m_PixScale));

	if (pCode != NULL) {
		// Single, BMP
		int cx = pCode->width;
		int cy = pCode->width;
		int vcx = cx * pixScale;
		int vcy = cy * pixScale;
		int stride = ((vcx +31) >> 3) & (~3);

		CByteArray bits;
		bits.SetSize(stride * vcy);

		for (int y=0; y<cy; y++) {
			for (int x=0; x<cx; x++) {
				BYTE b = pCode->data[cx * y + x];
				if (0 == (1 & b)) {
					for (int vy=0; vy<pixScale; vy++) {
						for (int vx=0; vx<pixScale; vx++) {
							const int ax = (x * pixScale + vx);
							bits.ElementAt(stride * (vcy - (y * pixScale + vy) - 1) + (ax >> 3)) |= (128 >> (ax & 7));
						}
					}
				}
			}
		}

		char cBI[sizeof(BITMAPINFOHEADER) + sizeof(RGBQUAD)*256];
		PBITMAPINFO pBI = reinterpret_cast<PBITMAPINFO>(cBI);
		pBI->bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
		pBI->bmiHeader.biWidth = vcx;
		pBI->bmiHeader.biHeight = vcx;
		pBI->bmiHeader.biPlanes = 1;
		pBI->bmiHeader.biBitCount = 1;
		pBI->bmiHeader.biCompression = BI_RGB;
		pBI->bmiHeader.biXPelsPerMeter = int(m_PelsPerMeter);
		pBI->bmiHeader.biYPelsPerMeter = int(m_PelsPerMeter);
		pBI->bmiHeader.biClrUsed = 2;
		pBI->bmiHeader.biClrImportant = 2;
		{
			static const RGBQUAD black = {0, 0, 0, 0};
			static const RGBQUAD white = {255,255,255,0};
			pBI->bmiColors[0] = black;
			pBI->bmiColors[1] = white;
		}

		void *pvBits = NULL;
		HBITMAP hbm = CreateDIBSection(NULL, pBI, DIB_RGB_COLORS, &pvBits, NULL, 0);
		if (hbm != NULL) {
			memcpy(pvBits, bits.GetData(), stride * vcy);
			PICTDESC pd;
			pd.cbSizeofstruct = sizeof(pd);
			pd.picType = PICTYPE_BITMAP;
			pd.bmp.hbitmap = hbm;
			pd.bmp.hpal = NULL;

			CComPtr<IPictureDisp> pict;
			if (SUCCEEDED(hr = OleCreatePictureIndirect(&pd, IID_IPictureDisp, true, reinterpret_cast<void **>(&pict)))) {
				hbm = NULL;
				SetPicture(pict);
				SetErrorInfo();
				return;
			}

			DeleteObject(hbm);
		}
		else {
			SetErrorInfo(true, ObtainErrorMessage(IDS_GENERATION_FAILURE), IDB_GENERATION_FAILURE);
		}
	}
	else if (pCodeListOrg != NULL) {
		// Multi, BMP
		int maxlen = 0;
		QRcode_List *pCodeList;
		for (pCodeList = pCodeListOrg; pCodeList != NULL; pCodeList = pCodeList->next) {
			maxlen += pCodeList->code->width;
		}
		const int seplen = 15;

		int cx = 0;
		int cy = 0;
		for (pCodeList = pCodeListOrg; pCodeList != NULL; pCodeList = pCodeList->next) {
			if (cx != 0) {
				cx += seplen;
			}
			cx += pCodeList->code->width;
			cy = max(cy, pCodeList->code->width);
		}

		bool isHorz = true;
		switch (m_EncodeMode) {
			case EncodeModeVertStrucShiftJIS:
			case EncodeModeVertStrucAlphaNum:
			case EncodeModeVertStrucNum:
			case EncodeModeVertStruc8bit:
				isHorz = false;
				break;
		}

		int vcx = (isHorz ? cx : cy) * pixScale;
		int vcy = (isHorz ? cy : cx) * pixScale;
		int stride = ((vcx +31) >> 3) & (~3);

		CByteArray bits;
		bits.SetSize(stride * vcy);

		int step = 0;
		for (pCodeList = pCodeListOrg; pCodeList != NULL; pCodeList = pCodeList->next) {
			if (step != 0) {
				for (int wx=0; wx<seplen; wx++) {
					for (int wy=0; wy<cy; wy++) {
						const int bx = (isHorz ? step + wx : wy) * pixScale;
						const int by = (isHorz ? wy : step + wx) * pixScale;
						for (int vy=0; vy<pixScale; vy++) {
							for (int vx=0; vx<pixScale; vx++) {
								const int ax = (bx +vx);
								bits.ElementAt(stride * (vcy - (by +vy) - 1) + (ax >> 3)) |= (128 >> (ax & 7));
							}
						}
					}
				}
				step += seplen;
			}
			int cx1 = pCodeList->code->width;
			int cy1 = pCodeList->code->width;
			for (int y=0; y<cy1; y++) {
				for (int x=0; x<cx1; x++) {
					const int bx = (isHorz ? step + x : x) * pixScale;
					const int by = (isHorz ? y : step + y) * pixScale;
					const BYTE b = pCodeList->code->data[cx1 * y + x];
					if (0 == (1 & b)) {
						for (int vy=0; vy<pixScale; vy++) {
							for (int vx=0; vx<pixScale; vx++) {
								const int ax = (bx +vx);
								bits.ElementAt(stride * (vcy - (by +vy) - 1) + (ax >> 3)) |= (128 >> (ax & 7));
							}
						}
					}
				}
			}
			step += cx1;
		}

		char cBI[sizeof(BITMAPINFOHEADER) + sizeof(RGBQUAD)*256];
		PBITMAPINFO pBI = reinterpret_cast<PBITMAPINFO>(cBI);
		pBI->bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
		pBI->bmiHeader.biWidth = vcx;
		pBI->bmiHeader.biHeight = vcy;
		pBI->bmiHeader.biPlanes = 1;
		pBI->bmiHeader.biBitCount = 1;
		pBI->bmiHeader.biCompression = BI_RGB;
		pBI->bmiHeader.biXPelsPerMeter = int(m_PelsPerMeter);
		pBI->bmiHeader.biYPelsPerMeter = int(m_PelsPerMeter);
		pBI->bmiHeader.biClrUsed = 2;
		pBI->bmiHeader.biClrImportant = 2;
		{
			static const RGBQUAD black = {0, 0, 0, 0};
			static const RGBQUAD white = {255,255,255,0};
			pBI->bmiColors[0] = black;
			pBI->bmiColors[1] = white;
		}

		void *pvBits = NULL;
		HBITMAP hbm = CreateDIBSection(NULL, pBI, DIB_RGB_COLORS, &pvBits, NULL, 0);
		if (hbm != NULL) {
			memcpy(pvBits, bits.GetData(), stride * vcy);
			PICTDESC pd;
			pd.cbSizeofstruct = sizeof(pd);
			pd.picType = PICTYPE_BITMAP;
			pd.bmp.hbitmap = hbm;
			pd.bmp.hpal = NULL;

			CComPtr<IPictureDisp> pict;
			if (SUCCEEDED(hr = OleCreatePictureIndirect(&pd, IID_IPictureDisp, true, reinterpret_cast<void **>(&pict)))) {
				hbm = NULL;
				SetPicture(pict);
				SetErrorInfo();
				return;
			}

			DeleteObject(hbm);
		}
		else {
			SetErrorInfo(true, ObtainErrorMessage(IDS_GENERATION_FAILURE), IDB_GENERATION_FAILURE);
		}
	}
	else {
		// Other
		SetErrorInfo(true, ObtainErrorMessage(IDS_GENERATION_FAILURE), IDB_GENERATION_FAILURE);
	}

	//SetEmptyPicture();
	return;
}

DOUBLE CQRCStampCtrl::GetPelsPerMeter(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	return m_PelsPerMeter;
}

void CQRCStampCtrl::SetPelsPerMeter(DOUBLE newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	if (m_PelsPerMeter != newVal) {
		m_PelsPerMeter = newVal;
		PelsPerMeterChange();
		UpdateBarcode();

		SetModifiedFlag();
	}
}

IPictureDisp* CQRCStampCtrl::GetPicture(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	IPictureDisp *pRes = NULL;
	m_Picture.CopyTo(&pRes);
	return pRes;
}

void CQRCStampCtrl::SetPicture(IPictureDisp* pVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	if (m_Picture != pVal) {
		m_Picture = pVal;
		PictureChange();
		InvalidateControl();

		SetModifiedFlag();
	}
}

LONG CQRCStampCtrl::GetSizeMode(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	return m_sizeMode;
}

void CQRCStampCtrl::SetSizeMode(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	if (m_sizeMode != newVal) {
		m_sizeMode = static_cast<SizeModeEnum>(newVal);
		SizeModeChange();
		InvalidateControl();
		SetModifiedFlag();
	}

}

LONG CQRCStampCtrl::GetErrorCorrectLevel(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	return m_ErrorCorrectLevel;
}

void CQRCStampCtrl::SetErrorCorrectLevel(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	if (m_ErrorCorrectLevel != newVal) {
		m_ErrorCorrectLevel = static_cast<QRErrorCorrectLevelEnum>(newVal);
		ErrorCorrectLevelChange();
		UpdateBarcode();
		SetModifiedFlag();
	}
}

LONG CQRCStampCtrl::GetEncodeMode(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	return m_EncodeMode;
}

void CQRCStampCtrl::SetEncodeMode(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	if (m_EncodeMode != newVal) {
		m_EncodeMode = static_cast<EncodeModeEnum>(newVal);
		EncodeModeChange();
		UpdateBarcode();
		SetModifiedFlag();
	}
}

void CQRCStampCtrl::OnFontChanged() {
	InvalidateControl();
}

void CQRCStampCtrl::OnBackColorChanged() {
	UpdateBarcode();
	InvalidateControl();
}

void CQRCStampCtrl::OnForeColorChanged() {
	UpdateBarcode();
	InvalidateControl();
}

LONG CQRCStampCtrl::GetPixScale(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	return m_PixScale;
}

void CQRCStampCtrl::SetPixScale(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	m_PixScale = newVal;

	PixScaleChange();

	UpdateBarcode();

	SetModifiedFlag();
}

LONG CQRCStampCtrl::GetQRVersion(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	return m_QRVersion;
}

void CQRCStampCtrl::SetQRVersion(LONG newVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	m_QRVersion = newVal;

	QRVersionChange();

	UpdateBarcode();

	SetModifiedFlag();
}

void CQRCStampCtrl::SetEmptyPicture() {
#if 1
	SetPicture(NULL);
#else
	HRESULT hr;

	char cBI[sizeof(BITMAPINFOHEADER) + sizeof(RGBQUAD)*256];
	PBITMAPINFO pBI = reinterpret_cast<PBITMAPINFO>(cBI);
	pBI->bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
	pBI->bmiHeader.biWidth = 1;
	pBI->bmiHeader.biHeight = 1;
	pBI->bmiHeader.biPlanes = 1;
	pBI->bmiHeader.biBitCount = 8;
	pBI->bmiHeader.biCompression = BI_RGB;
	pBI->bmiHeader.biXPelsPerMeter = int(m_PelsPerMeter);
	pBI->bmiHeader.biYPelsPerMeter = int(m_PelsPerMeter);
	pBI->bmiHeader.biClrUsed = 256;
	pBI->bmiHeader.biClrImportant = 256;
	for (int x=0; x<256; x++) {
		BYTE v = ((x & 1) == 1) ? 0 : 255;
		RGBQUAD q = {v, v, v, 0};
		pBI->bmiColors[x] = q;
	}

	void *pvBits = NULL;
	HBITMAP hbm = CreateDIBSection(NULL, pBI, DIB_RGB_COLORS, &pvBits, NULL, 0);
	if (hbm != NULL) {
		memcpy(pvBits, "\x00", 1);
		PICTDESC pd;
		pd.cbSizeofstruct = sizeof(pd);
		pd.picType = PICTYPE_BITMAP;
		pd.bmp.hbitmap = hbm;
		pd.bmp.hpal = NULL;

		CComPtr<IPictureDisp> pict;
		if (SUCCEEDED(hr = OleCreatePictureIndirect(&pd, IID_IPictureDisp, true, reinterpret_cast<void **>(&pict)))) {
			hbm = NULL;
			SetPicture(pict);
			return;
		}

		DeleteObject(hbm);
	}
#endif
}

BOOL CQRCStampCtrl::GetHasErrors() {
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return m_HasErrors;
}

void CQRCStampCtrl::SetHasErrors(BOOL newVal) {
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	// なにもしない
}

BSTR CQRCStampCtrl::GetErrorMessage() {
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return m_ErrorMessage.Copy();
}

void CQRCStampCtrl::SetErrorMessage(BSTR newVal) {
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	// なにもしない
}
