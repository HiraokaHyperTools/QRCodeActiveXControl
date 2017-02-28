#pragma once

// QRCStampCtrl.h : CQRCStampCtrl ActiveX コントロール クラスの宣言です。

typedef enum tagSizeModeEnum {
	SizeModeCenter = 0,
	SizeModeStretch = 1,
	SizeModeZoom = 2,
} SizeModeEnum;

typedef enum tagQRErrorCorrectLevel {
	ErrorCorrectLevelL = 1,
	ErrorCorrectLevelM = 0,
	ErrorCorrectLevelQ = 3,
	ErrorCorrectLevelH = 2,
} QRErrorCorrectLevelEnum;

typedef enum tagEncodeMode {
	EncodeModeShiftJIS = 0,
	EncodeModeAlphaNum,
	EncodeModeNum,
	EncodeMode8bit,
	EncodeModeStrucShiftJIS = 256,
	EncodeModeStrucAlphaNum,
	EncodeModeStrucNum,
	EncodeModeStruc8bit,
	EncodeModeVertStrucShiftJIS = 512,
	EncodeModeVertStrucAlphaNum,
	EncodeModeVertStrucNum,
	EncodeModeVertStruc8bit,
} EncodeModeEnum;

// CQRCStampCtrl : 実装に関しては QRCStampCtrl.cpp を参照してください。

class CQRCStampCtrl : public COleControl
{
	DECLARE_DYNCREATE(CQRCStampCtrl)

// コンストラクタ
public:
	CQRCStampCtrl();

// オーバーライド
public:
	virtual void OnDraw(CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid);
	virtual void DoPropExchange(CPropExchange* pPX);
	virtual void OnResetState();
	virtual DWORD GetControlFlags();

// 実装
protected:
	~CQRCStampCtrl();

	DECLARE_OLECREATE_EX(CQRCStampCtrl)    // クラス ファクトリ と guid
	DECLARE_OLETYPELIB(CQRCStampCtrl)      // GetTypeInfo
	DECLARE_PROPPAGEIDS(CQRCStampCtrl)     // プロパティ ページ ID
	DECLARE_OLECTLTYPE(CQRCStampCtrl)		// タイプ名とその他のステータス

// メッセージ マップ
	DECLARE_MESSAGE_MAP()

// ディスパッチ マップ
	DECLARE_DISPATCH_MAP()

	afx_msg void AboutBox();

// イベント マップ
	DECLARE_EVENT_MAP()

// ディスパッチ と イベント ID
public:
	enum {
		eventidErrorMessageChange = 1010L,
		dispidErrorMessage = 1010,

		eventidHasErrorsChange = 1009L,
		dispidHasErrors = 1009,

		eventidQRVersionChange = 1008L,
		dispidQRVersion = 1008,

		eventidPixScaleChange = 1007L,
		dispidPixScale = 1007,

		eventidEncodeModeChange = 1006L,
		dispidEncodeMode = 1006,

		eventidErrorCorrectLevelChange = 1005L,
		dispidErrorCorrectLevel = 1005,

		eventidSizeModeChange = 1004L,
		dispidSizeMode = 1004,

		eventidPictureChange = 1003L,
		dispidPicture = 1003,

		eventidPelsPerMeterChange = 1002L,
		dispidPelsPerMeter = 1002,

		eventidValueChange = 1001L,
		dispidValue = 1001
	};
protected:
	VARIANT GetValue(void);
	void SetValue(VARIANT *newVal);
	CString m_strValue;
	void UpdateBarcode();
	double m_PelsPerMeter;
	CComPtr<IPictureDisp> m_Picture;
	LONG m_sizeMode;
	LONG m_ErrorCorrectLevel;
	bool m_HasErrors;
	CComBSTR m_ErrorMessage;

	void ValueChange() {
		FireEvent(eventidValueChange, EVENT_PARAM(VTS_NONE)
			);
	}

	void PelsPerMeterChange() {
		FireEvent(eventidPelsPerMeterChange, EVENT_PARAM(VTS_NONE)
			);
	}

	void PictureChange() {
		FireEvent(eventidPictureChange, EVENT_PARAM(VTS_NONE)
			);
	}
	DOUBLE GetPelsPerMeter(void);
	void SetPelsPerMeter(DOUBLE newVal);
	IPictureDisp* GetPicture(void);
	void SetPicture(IPictureDisp* pVal);
	void SetEmptyPicture();
	LONG GetSizeMode(void);
	void SetSizeMode(LONG newVal);
	LONG GetErrorCorrectLevel(void);
	void SetErrorCorrectLevel(LONG newVal);

	void SizeModeChange(void) {
		FireEvent(eventidSizeModeChange, EVENT_PARAM(VTS_NONE));
	}

	void ErrorCorrectLevelChange(void) {
		FireEvent(eventidErrorCorrectLevelChange, EVENT_PARAM(VTS_NONE));
	}
	LONG GetEncodeMode(void);
	void SetEncodeMode(LONG newVal);
	LONG m_EncodeMode;

	void EncodeModeChange(void)
	{
		FireEvent(eventidEncodeModeChange, EVENT_PARAM(VTS_NONE));
	}

	LONG GetPixScale(void);
	void SetPixScale(LONG newVal);
	LONG m_PixScale;

	void PixScaleChange(void)
	{
		FireEvent(eventidPixScaleChange, EVENT_PARAM(VTS_NONE));
	}

	LONG GetQRVersion(void);
	void SetQRVersion(LONG newVal);
	LONG m_QRVersion;

	void QRVersionChange(void)
	{
		FireEvent(eventidQRVersionChange, EVENT_PARAM(VTS_NONE));
	}

	BOOL GetHasErrors();
	void SetHasErrors(BOOL newVal);
	void HasErrorsChange(void)
	{
		FireEvent(eventidHasErrorsChange, EVENT_PARAM(VTS_NONE));
	}

	BSTR GetErrorMessage();
	void SetErrorMessage(BSTR newVal);
	void ErrorMessageChange(void)
	{
		FireEvent(eventidErrorMessageChange, EVENT_PARAM(VTS_NONE));
	}

	CComBSTR ObtainErrorMessage(int idString) {
		CString str;
		str.LoadString(idString);
		return CComBSTR(str);
	}

	void SetErrorInfo(bool hasErrors = false, BSTR errorMessage = NULL, int idBitmap = -1) {
		HRESULT hr;

		bool fireHasErrorsChange = false;
		if (m_HasErrors != hasErrors) {
			m_HasErrors = hasErrors;
			fireHasErrorsChange = true;
		}

		bool fireErrorMessageChange = false;
		if (m_ErrorMessage != errorMessage) {
			m_ErrorMessage = errorMessage;
			fireErrorMessageChange = true;
		}

		if (fireHasErrorsChange) {
			HasErrorsChange();
		}
		if (fireErrorMessageChange) {
			ErrorMessageChange();
		}

		if (idBitmap > 0) {
			CBitmap bitmap;
			bitmap.LoadBitmap(idBitmap);

			if (bitmap.m_hObject != NULL) {
				PICTDESC pd;
				pd.cbSizeofstruct = sizeof(pd);
				pd.picType = PICTYPE_BITMAP;
				pd.bmp.hbitmap = reinterpret_cast<HBITMAP>(bitmap.Detach());
				pd.bmp.hpal = NULL;

				CComPtr<IPictureDisp> pict;
				if (SUCCEEDED(hr = OleCreatePictureIndirect(&pd, IID_IPictureDisp, true, reinterpret_cast<void **>(&pict)))) {
					SetPicture(pict);
				}
				else {
					DeleteObject(pd.bmp.hbitmap);
				}
			}
		}
	}

public:
	virtual void OnFontChanged();
	virtual void OnBackColorChanged();
	virtual void OnForeColorChanged();
};
