
// KTVDlg.h : 头文件
//

#pragma once

#include <strmif.h>
#include <MMSystem.h>
#include <control.h>
#include "afxwin.h"
#include <vector>
#include <string>

#ifndef SAFE_RELEASE
#define SAFE_RELEASE(x)  if(x) {x->Release(); x=0;}
#endif

struct IGraphBuilder;
struct ICaptureGraphBuilder2;
struct ICreateDevEnum;
struct IBaseFilter;
struct IMoniker;

enum WinVersion {
    WINVERSION_PRE_2000 = 0,  // Not supported
    WINVERSION_2000 = 1,
    WINVERSION_XP = 2,
    WINVERSION_SERVER_2003 = 3,
    WINVERSION_VISTA = 4,
    WINVERSION_2008 = 5,
    WINVERSION_WIN7 = 6,
};


// CKTVDlg 对话框
class CKTVDlg : public CDialog
{
// 构造
public:
	CKTVDlg(CWnd* pParent = NULL);	// 标准构造函数

// 对话框数据
	enum { IDD = IDD_KTV_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持


// 实现
protected:
	HICON m_hIcon;
    IGraphBuilder* m_playsoundGraphBuilder;
    IMediaControl* m_playsoundMediaControl;
    IGraphBuilder* m_microphoneGraphBuilder;
    IMediaControl* m_microphoneMediaControl;
    IGraphBuilder* m_mixGraphBuilder;
    IMediaControl* m_mixMediaControl;

	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
    afx_msg void OnBnClickedButton1();
    void FreeInterfaces();
    HRESULT initPlaySoundGraph();
    HRESULT initMicrophoneGraph();
    HRESULT initMixGraph();
    HRESULT initCapture();
    HRESULT EnumFiltersAndMonikersToList(IEnumMoniker *pEnumCat);
    HRESULT EnumPinsOnFilter (IBaseFilter *pFilter, PIN_DIRECTION PinDir, int index);
    HRESULT ActivateSelectedInputPin();
    HRESULT SetAudioProperties(IBaseFilter* filter);
    bool checkCapture(); 
    bool selectCapture(); 
    int findFilter(IBaseFilter** filter, wchar_t* str = NULL, int index = -1);
    HRESULT activePin(IBaseFilter* filter, int index, wchar_t* str = NULL, int pinIndex = -1);

    bool buildMicrophoneGraph(int filterIndex = -1, int pinIndex = -1);
    bool buildMixGraph(int filterIndex = -1, int pinIndex = -1);


    afx_msg void OnBnClickedButton2();
    afx_msg void OnBnClickedCancel();
    CComboBox m_microphoneCombox;
    CComboBox m_microphonePinCombox;
    CComboBox m_mixCaptureCombox;
    CComboBox m_mixPinCombox;

    struct TCaptureFilter
    {
        std::wstring Name;
        IMoniker* Moniker;
        std::vector<std::wstring> PinVec; 
    };
    std::vector<TCaptureFilter> m_captureFilterVec;

    IBaseFilter* m_pInputDevice;
    IBaseFilter* m_mixDevice;
    IBaseFilter* m_aviMuxer;
    IBaseFilter* m_pFileWriter;

    int m_captureFilter;
    int m_capturePin;
    int m_mixFilter;
    int m_mixPin;

private:
    //std::vector<>
public:
    afx_msg void OnBnClickedButtonSet();
    afx_msg void OnCbnSelchangeCombo1();
    afx_msg void OnCbnSelchangeCombo2();
    afx_msg void OnCbnSelchangeCombo3();
    afx_msg void OnCbnSelchangeCombo4();
};
