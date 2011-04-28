
// KTVDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "KTV.h"
#include "KTVDlg.h"

#include <streams.h>
#include <DShow.h>
#include <InitGuid.h>
#include <uuids.h>
#include <Mmdeviceapi.h>
#include <Functiondiscoverykeys_devpkey.h>

#include "debug_util.h"
#include "dshow_enum_helper.h"
#include "intrusive_ptr_helper.h"

using std::wstring;
using boost::intrusive_ptr;

#define DEFAULT_BUFFER_TIME ((float) 0.05)  /* 10 milliseconds*/


#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CKTVDlg 对话框

WinVersion GetWinVersion() {
    static bool checked_version = false;
    static WinVersion win_version = WINVERSION_PRE_2000;
    if (!checked_version) {
        OSVERSIONINFOEX version_info;
        version_info.dwOSVersionInfoSize = sizeof version_info;
        GetVersionEx(reinterpret_cast<OSVERSIONINFO*>(&version_info));
        if (version_info.dwMajorVersion == 5) {
            switch (version_info.dwMinorVersion) {
        case 0:
            win_version = WINVERSION_2000;
            break;
        case 1:
            win_version = WINVERSION_XP;
            break;
        case 2:
        default:
            win_version = WINVERSION_SERVER_2003;
            break;
            }
        } else if (version_info.dwMajorVersion == 6) {
            if (version_info.wProductType != VER_NT_WORKSTATION) {
                // 2008 is 6.0, and 2008 R2 is 6.1.
                win_version = WINVERSION_2008;
            } else {
                if (version_info.dwMinorVersion == 0) {
                    win_version = WINVERSION_VISTA;
                } else {
                    win_version = WINVERSION_WIN7;
                }
            }
        } else if (version_info.dwMajorVersion > 6) {
            win_version = WINVERSION_WIN7;
        }
        checked_version = true;
    }
    return win_version;
}

HRESULT GetPin( IBaseFilter * pFilter, PIN_DIRECTION dirrequired, int iNum, IPin **ppPin)
{
    CComPtr< IEnumPins > pEnum;
    *ppPin = NULL;

    if (!pFilter)
        return E_POINTER;

    HRESULT hr = pFilter->EnumPins(&pEnum);
    if(FAILED(hr)) 
        return hr;

    ULONG ulFound;
    IPin *pPin;
    hr = E_FAIL;

    while(S_OK == pEnum->Next(1, &pPin, &ulFound))
    {
        PIN_DIRECTION pindir = (PIN_DIRECTION)3;

        pPin->QueryDirection(&pindir);
        if(pindir == dirrequired)
        {
            if(iNum == 0)
            {
                *ppPin = pPin;  // Return the pin's interface
                hr = S_OK;      // Found requested pin, so clear error
                break;
            }
            iNum--;
        } 

        pPin->Release();
    } 

    return hr;
}

void UtilDeleteMediaType(AM_MEDIA_TYPE *pmt)
{
    // Allow NULL pointers for coding simplicity
    if (pmt == NULL) {
        return;
    }

    // Free media type's format data
    if (pmt->cbFormat != 0) 
    {
        CoTaskMemFree((PVOID)pmt->pbFormat);

        // Strictly unnecessary but tidier
        pmt->cbFormat = 0;
        pmt->pbFormat = NULL;
    }

    // Release interface
    if (pmt->pUnk != NULL) 
    {
        pmt->pUnk->Release();
        pmt->pUnk = NULL;
    }

    // Free media type
    CoTaskMemFree((PVOID)pmt);
}


CKTVDlg::CKTVDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CKTVDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CKTVDlg::DoDataExchange(CDataExchange* pDX)
{
    CDialog::DoDataExchange(pDX);
    DDX_Control(pDX, IDC_COMBO1, m_microphoneCombox);
    DDX_Control(pDX, IDC_COMBO2, m_microphonePinCombox);
    DDX_Control(pDX, IDC_COMBO3, m_mixCaptureCombox);
    DDX_Control(pDX, IDC_COMBO4, m_mixPinCombox);
}

BEGIN_MESSAGE_MAP(CKTVDlg, CDialog)
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	//}}AFX_MSG_MAP
    ON_BN_CLICKED(IDC_BUTTON1, &CKTVDlg::OnBnClickedButton1)
    ON_BN_CLICKED(IDC_BUTTON2, &CKTVDlg::OnBnClickedButton2)
    ON_BN_CLICKED(IDCANCEL, &CKTVDlg::OnBnClickedCancel)
    ON_BN_CLICKED(IDC_BUTTON_SET, &CKTVDlg::OnBnClickedButtonSet)
    ON_CBN_SELCHANGE(IDC_COMBO1, &CKTVDlg::OnCbnSelchangeCombo1)
    ON_CBN_SELCHANGE(IDC_COMBO2, &CKTVDlg::OnCbnSelchangeCombo2)
    ON_CBN_SELCHANGE(IDC_COMBO3, &CKTVDlg::OnCbnSelchangeCombo3)
    ON_CBN_SELCHANGE(IDC_COMBO4, &CKTVDlg::OnCbnSelchangeCombo4)
END_MESSAGE_MAP()


// CKTVDlg 消息处理程序

BOOL CKTVDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// 设置此对话框的图标。当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO: 在此添加额外的初始化代码
    CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);

    m_pInputDevice = NULL;
    m_mixDevice = NULL;
    m_captureFilter = -1;
    m_capturePin = -1;
    m_mixFilter = -1;
    m_mixPin = -1;
    HRESULT r = initCapture();

    initPlaySoundGraph();
    initMicrophoneGraph();
    initMixGraph();
    
    if (!checkCapture())
    {
        m_microphoneCombox.ShowWindow(SW_SHOW);
        m_microphonePinCombox.ShowWindow(SW_SHOW);
        m_mixCaptureCombox.ShowWindow(SW_SHOW);
        m_mixPinCombox.ShowWindow(SW_SHOW);
        GetDlgItem(IDC_STATIC)->ShowWindow(SW_SHOW);
        GetDlgItem(IDC_STATIC1)->ShowWindow(SW_SHOW);
        GetDlgItem(IDC_BUTTON_SET)->ShowWindow(SW_SHOW);
        m_microphoneCombox.ResetContent();
        m_mixCaptureCombox.ResetContent();
        for (int i = 0; i < m_captureFilterVec.size(); i++)
        {
            m_microphoneCombox.AddString(m_captureFilterVec[i].Name.c_str());
            m_mixCaptureCombox.AddString(m_captureFilterVec[i].Name.c_str());
        }
        m_microphoneCombox.SetCurSel(0);
        m_mixCaptureCombox.SetCurSel(0);

        m_microphonePinCombox.ResetContent();
        m_mixPinCombox.ResetContent();
        for (int i = 0; i < m_captureFilterVec[0].PinVec.size(); i++)
        {
            m_microphonePinCombox.AddString(m_captureFilterVec[0].PinVec[i].c_str());
            m_mixPinCombox.AddString(m_captureFilterVec[0].PinVec[i].c_str());
        }

        m_microphonePinCombox.SetCurSel(0);
        m_mixPinCombox.SetCurSel(0);
        m_captureFilter = 0;
        m_capturePin = 0;
        m_mixFilter = 0;
        m_mixPin = 0;
    }
    else
    {
        if (!buildMicrophoneGraph())
             PostQuitMessage(0);

        if (!buildMixGraph())
            PostQuitMessage(0);      
    }




	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CKTVDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作区矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
HCURSOR CKTVDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}


void CKTVDlg::OnBnClickedButton1()
{
    // TODO: 在此添加控件通知处理程序代码
    m_playsoundMediaControl->Run();
    m_microphoneMediaControl->Run();
    m_mixMediaControl->Run();
}

void CKTVDlg::FreeInterfaces()
{
    SAFE_RELEASE(m_playsoundGraphBuilder);
    SAFE_RELEASE(m_playsoundMediaControl);
    SAFE_RELEASE(m_microphoneGraphBuilder);
    SAFE_RELEASE(m_microphoneMediaControl);    
    SAFE_RELEASE(m_pInputDevice);
    SAFE_RELEASE(m_mixDevice);
}

HRESULT CKTVDlg::initPlaySoundGraph()
{
    HRESULT r = CoCreateInstance(CLSID_FilterGraph, NULL, CLSCTX_INPROC_SERVER, 
        IID_IGraphBuilder, (void**)&m_playsoundGraphBuilder);
    if (FAILED(r))
        return r;

    r = m_playsoundGraphBuilder->QueryInterface(IID_IMediaControl,
        (void **)&m_playsoundMediaControl);
    if (FAILED(r))
        return r;

    r = m_playsoundGraphBuilder->RenderFile(L"abc.mp3",NULL);
    if (FAILED(r))
        return r;

    return S_OK;
}

void CKTVDlg::OnBnClickedButton2()
{
    // TODO: 在此添加控件通知处理程序代码
    m_playsoundMediaControl->Pause();
    m_microphoneMediaControl->Pause();
    m_mixMediaControl->Pause();
}

void CKTVDlg::OnBnClickedCancel()
{
    // TODO: 在此添加控件通知处理程序代码
    m_playsoundMediaControl->Stop();
    m_microphoneMediaControl->Stop();
    m_mixMediaControl->Stop();
    CoUninitialize();
    OnCancel();
}

HRESULT CKTVDlg::initCapture()
{
    int mixerNum = 0;
    int microphoneNum = 0;
    ICreateDevEnum* sysDevEnum = NULL;
    HRESULT r = CoCreateInstance(CLSID_SystemDeviceEnum, NULL, 
                                 CLSCTX_INPROC, IID_ICreateDevEnum, 
                                 (void **)&sysDevEnum);

    if FAILED(r)
        return r;

    IEnumMoniker *pEnumCat = NULL;
    r = sysDevEnum->CreateClassEnumerator(CLSID_AudioInputDeviceCategory, &pEnumCat, 0);

    if (SUCCEEDED(r))
    {
        // Enumerate all filters using the category enumerator
        r = EnumFiltersAndMonikersToList(pEnumCat);

        pEnumCat->Release();
    }
    sysDevEnum->Release();


    IMoniker *pMoniker=0;
    IBaseFilter* filter;
    for (int i = 0;i < m_captureFilterVec.size(); i++)
    {
        pMoniker = m_captureFilterVec[i].Moniker;
        // Use the moniker to create the specified audio capture device
        r = pMoniker->BindToObject(0, 0, IID_IBaseFilter, (void**)&filter);   
        if (FAILED(r))
            return r;

        r = EnumPinsOnFilter(filter, PINDIR_INPUT, i);
        if (FAILED(r))
            return r;

        filter->Release();
    }
    
    return S_OK;
    

//     if (GetWinVersion() >= WINVERSION_VISTA)
//     {
//         IMMDeviceEnumerator* pEnumerator = NULL;
//         HRESULT hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), NULL,
//             CLSCTX_ALL, __uuidof(IMMDeviceEnumerator),
//             (void**)&pEnumerator);
// 
//         IMMDeviceCollection* pCollection;
//         hr = pEnumerator->EnumAudioEndpoints(
//             eCapture, DEVICE_STATE_ACTIVE,
//             &pCollection);
// 
// 
//         UINT  count;
//         hr = pCollection->GetCount(&count);
//         IMMDevice  *pEndpoint = NULL;
//         LPWSTR pwszID = NULL;
//         IPropertyStore  *pProps = NULL;
// 
//         // Each loop prints the name of an endpoint device.
//         for (ULONG i = 0; i < count; i++)
//         {
//             // Get pointer to endpoint number i.
//             hr = pCollection->Item(i, &pEndpoint);
// 
// 
//             // Get the endpoint ID string.
//             hr = pEndpoint->GetId(&pwszID);
// 
// 
//             hr = pEndpoint->OpenPropertyStore(
//                 STGM_READ, &pProps);
// 
//             PROPVARIANT varName;
//             // Initialize container for property value.
//             PropVariantInit(&varName);
// 
//             // Get the endpoint's friendly-name property.
//             hr = pProps->GetValue(
//                 PKEY_Device_FriendlyName, &varName);
// 
//             // Print endpoint friendly name and endpoint ID.
//             //         printf("Endpoint %d: \"%S\" (%S)\n",
//             //             i, varName.pwszVal, pwszID);
//             wstring str = varName.pwszVal;
//             if (str.find(L"混音") != wstring::npos || 
//                 str.find(L"Stereo Mix") != wstring::npos)
//             {
//                 mixerNum++;
//             }
// 
//             CoTaskMemFree(pwszID);
//             pwszID = NULL;
//             PropVariantClear(&varName);
//             SAFE_RELEASE(pProps);
//             SAFE_RELEASE(pEndpoint);
//         }
//         SAFE_RELEASE(pEnumerator);
//         SAFE_RELEASE(pCollection);
//     }
//     else
//     {
//         HMIXER m_hmx; //
//         UINT m_uMxId; //mixer的ID
//         MMRESULT err = mixerOpen(&m_hmx, 0,(DWORD)0, 0, 0);
//         if (MMSYSERR_NOERROR != err)
//         {
//             return E_FAIL;
//         }
//         err = mixerGetID((HMIXEROBJ)m_hmx, &m_uMxId, MIXER_OBJECTF_HMIXER);
//         if (MMSYSERR_NOERROR != err)
//         {
//             return E_FAIL;
//         }
//         MIXERCAPS     mixcaps;
//         unsigned long iNumDevs;
// 
//         /* Get the number of Mixer devices in this computer */
//         iNumDevs = mixerGetNumDevs();
// 
//         /* Go through all of those devices, displaying their IDs/names */
//         for (int i = 0; i < iNumDevs; i++)
//         {
//             /* Get info about the next device */
//             if (!mixerGetDevCaps(i, &mixcaps, sizeof(MIXERCAPS)))
//             {
//                 /* Display its ID number and name */
//                 wstring str =  mixcaps.szPname;
//                 if (str.find(L"立体声混音") != wstring::npos || 
//                     str.find(L"Stereo Mix") != wstring::npos)
//                 {
//                     mixerNum++;
//                 }
//                 //OutputDebugString(str);
//             }
//         }
// 
//         int num = waveInGetNumDevs();
//         WAVEINCAPS waveinCaps;
//         for (int i = 0; i <num; i++)
//         {
//             if (!waveInGetDevCaps(i, &waveinCaps, sizeof(WAVEINCAPS)))
//             {
//                 wchar_t str[1024];
//                 _sntprintf_s(str, 1024, L"Wavein ID #%u: %s\r\n", i, waveinCaps.szPname);
// 
//             }
//         }
//     }
//     if (mixerNum == 0)
//         return E_FAIL;
// 
// 
//     if (mixerNum != 1)
//         return S_FALSE;
// 
// 
//     return S_OK;
}

HRESULT CKTVDlg::EnumFiltersAndMonikersToList( IEnumMoniker *pEnumCat )
{
    HRESULT r=S_OK;
    IMoniker* pMoniker=0;
    ULONG cFetched=0;
    VARIANT varName={0};
    int nFilters=0;

    // If there are no filters of a requested type, show default string
    if (!pEnumCat)
    {
        return S_FALSE;
    }

    // Enumerate all items associated with the moniker
    while(pEnumCat->Next(1, &pMoniker, &cFetched) == S_OK)
    {
        IPropertyBag *pPropBag;
        ASSERT(pMoniker);

        // Associate moniker with a file
        r = pMoniker->BindToStorage(0, 0, IID_IPropertyBag, 
            (void **)&pPropBag);
        ASSERT(SUCCEEDED(r));
        ASSERT(pPropBag);
        if (FAILED(r))
            continue;

        // Read filter name from property bag
        varName.vt = VT_BSTR;
        r = pPropBag->Read(L"FriendlyName", &varName, 0);
        if (FAILED(r))
            continue;

        // Get filter name (converting BSTR name to a CString)
        CString str(varName.bstrVal);
        SysFreeString(varName.bstrVal);
        nFilters++;

        TCaptureFilter filter = {str, pMoniker};
        m_captureFilterVec.push_back(filter);

        // Cleanup interfaces
        SAFE_RELEASE(pPropBag);

        // Intentionally DO NOT release the pMoniker, since it is
        // stored in a listbox for later use
    }

    return r;
}

HRESULT CKTVDlg::EnumPinsOnFilter( IBaseFilter *pFilter, PIN_DIRECTION PinDir , int index)
{
    HRESULT r;
    IEnumPins  *pEnum = NULL;
    IPin *pPin = NULL;

    // Verify filter interface
    if (!pFilter)
        return E_NOINTERFACE;

    // Get pin enumerator
    r = pFilter->EnumPins(&pEnum);
    if (FAILED(r))
        return r;

    pEnum->Reset();

    // Enumerate all pins on this filter
    while((r = pEnum->Next(1, &pPin, 0)) == S_OK)
    {
        PIN_DIRECTION PinDirThis;

        r = pPin->QueryDirection(&PinDirThis);
        if (FAILED(r))
        {
            pPin->Release();
            continue;
        }

        // Does the pin's direction match the requested direction?
        if (PinDir == PinDirThis)
        {
            PIN_INFO pininfo={0};

            // Direction matches, so add pin name to listbox
            r = pPin->QueryPinInfo(&pininfo);
            if (SUCCEEDED(r))
            {
                wstring str = pininfo.achName;
                m_captureFilterVec[index].PinVec.push_back(str);
            }

            // The pininfo structure contains a reference to an IBaseFilter,
            // so you must release its reference to prevent resource a leak.
            pininfo.pFilter->Release();
        }
        pPin->Release();
    }
    pEnum->Release();

    return r;
}

bool CKTVDlg::checkCapture()
{
    int mixCount = 0;
    int captureCount = 0;

    for (int i = 0; i < m_captureFilterVec.size(); i++)
    {
        if (GetWinVersion() >= WINVERSION_VISTA)
        {
            if (m_captureFilterVec[i].Name.find(L"麦克风") != wstring::npos || m_captureFilterVec[i].Name.find(L"Mic") != wstring::npos)
            {
                captureCount++;
            }

            if (m_captureFilterVec[i].Name.find(L"混音") != wstring::npos || m_captureFilterVec[i].Name.find(L"Mix") != wstring::npos)
            {
                mixCount++;
            }
        }
        else
        {
            for (int j = 0; j < m_captureFilterVec[i].PinVec.size(); j++)
            {
                if (m_captureFilterVec[i].PinVec[j].find(L"麦克风") != wstring::npos || m_captureFilterVec[i].PinVec[j].find(L"Mic") != wstring::npos)
                {
                    captureCount++;
                }

                if (m_captureFilterVec[i].PinVec[j].find(L"混音") != wstring::npos || m_captureFilterVec[i].PinVec[j].find(L"Mix") != wstring::npos)
                {
                    mixCount++;
                }
            }
        }

    }
    
    if (captureCount != 1 || mixCount != 1)
        return false;

    return true;
}

int CKTVDlg::findFilter(IBaseFilter** filter,  wchar_t* str , int index)
{
    int r = -1;

    if (index != -1)
    {
        IMoniker *pMoniker = m_captureFilterVec[index].Moniker;
        HRESULT hr = pMoniker->BindToObject(0, 0, IID_IBaseFilter, (void**)filter);
        if (FAILED(hr))
            return -1;

        return index;
    }

    if (str == NULL)
    {
        IMoniker *pMoniker = m_captureFilterVec[0].Moniker;
        HRESULT hr = pMoniker->BindToObject(0, 0, IID_IBaseFilter, (void**)filter);
        if (FAILED(hr))
            return -1;

        return 0;
    }

    for (int i = 0; i < m_captureFilterVec.size(); i++)
    {
        if (m_captureFilterVec[i].Name.find(str) != wstring::npos)
        {
            IMoniker *pMoniker = m_captureFilterVec[i].Moniker;
            HRESULT hr = pMoniker->BindToObject(0, 0, IID_IBaseFilter, (void**)filter);
            if (FAILED(hr))
                return -1;

            return i;
        }

    }       

    return -1;
}

HRESULT CKTVDlg::initMicrophoneGraph()
{
    HRESULT r;
    r = CoCreateInstance(CLSID_FilterGraph, NULL, CLSCTX_INPROC_SERVER, 
        IID_IGraphBuilder, (void**)&m_microphoneGraphBuilder);

    if (FAILED(r) || !m_microphoneGraphBuilder)
        return E_NOINTERFACE;

    // Get useful interfaces from the graph builder
    r = m_microphoneGraphBuilder->QueryInterface(IID_IMediaControl,
        (void **)&m_microphoneMediaControl);

    if (FAILED(r))
        return r;

    return r;
}

HRESULT CKTVDlg::activePin(IBaseFilter* filter, int index, wchar_t* str /*= NULL*/ , int pinIndex  )
{
    HRESULT r = -1;
    int nActivePin = -1;

    if (pinIndex == -1)
    {
        if (GetWinVersion() >= WINVERSION_VISTA)
        {
            nActivePin = 0;
        }
        else
        {
            for (int j = 0; j < m_captureFilterVec[index].PinVec.size(); j++)
            {
                if (m_captureFilterVec[index].PinVec[j].find(str) != wstring::npos)
                {
                    nActivePin = j;
                    break;
                }
            }
        }
    }
    else
    {
        nActivePin = pinIndex;
    }

    IPin *pPin=0;
    IAMAudioInputMixer *pPinMixer;

    // How many pins are in the input pin list?
    int nPins = m_captureFilterVec[index].PinVec.size();


    // Activate the selected input pin and deactivate all others
    for (int i=0; i<nPins; i++)
    {
        // Get this pin's interface
        r = GetPin(filter, PINDIR_INPUT, i, &pPin);
        if (SUCCEEDED(r))
        {
            r = pPin->QueryInterface(IID_IAMAudioInputMixer, (void **)&pPinMixer);
            if (SUCCEEDED(r))
            {
                // If this is our selected pin, enable it
                if (i == nActivePin)
                {
                    // Set any other audio properties on this pin
                    //r = SetInputPinProperties(pPinMixer);

                    // If there is only one input pin, this method
                    // might return E_NOTIMPL.
                    r = pPinMixer->put_Enable(TRUE);
                }
                // Otherwise, disable it
                else
                {
                    //r = pPinMixer->put_Enable(FALSE);
                }

                pPinMixer->Release();
            }

            // Release pin interfaces
            pPin->Release();
        }
    }

    return S_OK;
}

HRESULT CKTVDlg::SetAudioProperties(IBaseFilter* filter)
{
    HRESULT hr=0;
    IPin *pPin=0;
    IAMBufferNegotiation *pNeg=0;
    IAMStreamConfig *pCfg=0;
    int nFrequency=44100;
    int nChannels = 2;
    int nBytesPerSample = 2;
    // Determine audio properties
    //     int nChannels = IsDlgButtonChecked(IDC_RADIO_MONO) ? 1 : 2;
    //     int nBytesPerSample = IsDlgButtonChecked(IDC_RADIO_8) ? 1 : 2;

    //     // Determine requested frequency
    //     if (IsDlgButtonChecked(IDC_RADIO_11KHZ))      
    //         nFrequency = 11025;
    //     else if (IsDlgButtonChecked(IDC_RADIO_22KHZ)) 
    //         nFrequency = 22050;
    //     else 
    //         nFrequency = 44100;

    // Find number of bytes in one second
    long lBytesPerSecond = (long) (nBytesPerSample * nFrequency * nChannels);

    // Set to 50ms worth of data    
    long lBufferSize = (long) ((float) lBytesPerSecond * DEFAULT_BUFFER_TIME);

    for (int i=0; i<2; i++)
    {
        hr = GetPin(filter, PINDIR_OUTPUT, i, &pPin);
        if (SUCCEEDED(hr))
        {
            // Get buffer negotiation interface
            hr = pPin->QueryInterface(IID_IAMBufferNegotiation, (void **)&pNeg);
            if (FAILED(hr))
            {
                pPin->Release();
                return hr;
            }

            // Set the buffer size based on selected settings
            ALLOCATOR_PROPERTIES prop={0};
            prop.cbBuffer = lBufferSize;
            prop.cBuffers = 6;
            prop.cbAlign = nBytesPerSample * nChannels;
            hr = pNeg->SuggestAllocatorProperties(&prop);
            pNeg->Release();

            // Now set the actual format of the audio data
            hr = pPin->QueryInterface(IID_IAMStreamConfig, (void **)&pCfg);
            if (FAILED(hr))
            {
                pPin->Release();
                return hr;
            }            

            // Read current media type/format
            AM_MEDIA_TYPE *pmt={0};
            hr = pCfg->GetFormat(&pmt);

            if (SUCCEEDED(hr))
            {
                // Fill in values for the new format
                WAVEFORMATEX *pWF = (WAVEFORMATEX *) pmt->pbFormat;
                pWF->nChannels = (WORD) nChannels;
                pWF->nSamplesPerSec = nFrequency;
                pWF->nAvgBytesPerSec = lBytesPerSecond;
                pWF->wBitsPerSample = (WORD) (nBytesPerSample * 8);
                pWF->nBlockAlign = (WORD) (nBytesPerSample * nChannels);

                // Set the new formattype for the output pin
                hr = pCfg->SetFormat(pmt);
                UtilDeleteMediaType(pmt);
            }

            // Release interfaces
            pCfg->Release();
            pPin->Release();
        }
        else
            break;
    }
    // No more output pins on this filter
    return hr;
}

HRESULT CKTVDlg::initMixGraph()
{
    HRESULT r;
    r = CoCreateInstance(CLSID_FilterGraph, NULL, CLSCTX_INPROC_SERVER, 
        IID_IGraphBuilder, (void**)&m_mixGraphBuilder);

    if (FAILED(r) || !m_mixGraphBuilder)
        return E_NOINTERFACE;

    // Get useful interfaces from the graph builder
    r = m_mixGraphBuilder->QueryInterface(IID_IMediaControl,
        (void **)&m_mixMediaControl);

    if (FAILED(r))
        return r;

    return r;
}

bool CKTVDlg::buildMicrophoneGraph(int filterIndex, int pinIndex)
{
    HRESULT r;
    int index = -1;

    if (filterIndex != -1)
    {
        index = findFilter(&m_pInputDevice, NULL, filterIndex);
    }
    else
    {
        if (GetWinVersion() >= WINVERSION_VISTA)
            index = findFilter(&m_pInputDevice, L"麦克风");
        else
            index = findFilter(&m_pInputDevice);
    }


    if (index == -1)
        return false;

    if (pinIndex != -1)
    {
        activePin(m_pInputDevice, index, NULL, pinIndex);
    }
    else
    {
        if (GetWinVersion() >= WINVERSION_VISTA)
            activePin(m_pInputDevice, index);
        else
        {
            activePin(m_pInputDevice, index, L"麦克风");
        }
    }


    SetAudioProperties(m_pInputDevice);

    r = m_microphoneGraphBuilder->AddFilter(m_pInputDevice, L"Microphone Capture");
    if (FAILED(r))
        return false;

    intrusive_ptr<IBaseFilter> renderer;

    r = CoCreateInstance(CLSID_AudioRender, 
        NULL, CLSCTX_INPROC_SERVER,IID_IBaseFilter,
        reinterpret_cast<void**>(&renderer));

    if (FAILED(r))
        return false;

    intrusive_ptr<IPin> sourceOutPin;
    r = GetPinByDirection(m_pInputDevice, 
        reinterpret_cast<IPin**>(&sourceOutPin), 
        PINDIR_OUTPUT);

    if (FAILED(r))
        return false;

    r = m_microphoneGraphBuilder->AddFilter((IBaseFilter *)renderer.get(), L"DShow Render");
    if (FAILED(r))
        return false;

    intrusive_ptr<IPin> rendererInputPin;
    r = GetPinByDirection(renderer.get(), 
        reinterpret_cast<IPin**>(&rendererInputPin), 
        PINDIR_INPUT);

    if (FAILED(r))
        return false;


    r = m_microphoneGraphBuilder->ConnectDirect(sourceOutPin.get(), rendererInputPin.get(), NULL);
    if (FAILED(r))
        return false;  

    return true;
}

bool CKTVDlg::buildMixGraph(int filterIndex, int pinIndex)
{
    int index = -1;
    if (filterIndex != -1)
    {
        index = findFilter(&m_mixDevice, NULL, filterIndex);
    }
    else
    {
        if (GetWinVersion() >= WINVERSION_VISTA)
            index = findFilter(&m_mixDevice, L"混音");
        else
            index = findFilter(&m_mixDevice);   
    }


    if (index == -1)
        return false;

    if (pinIndex != -1)
    {
        activePin(m_mixDevice, index, NULL, pinIndex);
    }
    else
    {
        if (GetWinVersion() >= WINVERSION_VISTA)
            activePin(m_mixDevice, index);
        else
            activePin(m_mixDevice, index, L"麦克风");
    }

    SetAudioProperties(m_mixDevice);

    HRESULT r = m_mixGraphBuilder->AddFilter(m_mixDevice, L"mix Capture");
    if (FAILED(r))
        return false;

    intrusive_ptr<IPin> sourceOutPin1;
    r = GetPinByDirection(m_mixDevice, 
        reinterpret_cast<IPin**>(&sourceOutPin1), 
        PINDIR_OUTPUT);

    if (FAILED(r))
        return false;


    // {3C78B8E2-6C4D-11d1-ADE2-0000F8754B99}
    static const GUID CLSID_WavDest =
    { 0x3c78b8e2, 0x6c4d, 0x11d1, { 0xad, 0xe2, 0x0, 0x0, 0xf8, 0x75, 0x4b, 0x99 } };

    HMODULE m = LoadLibrary(L"wavdest.ax");
    if (!m)
        return false;

    typedef HRESULT (__stdcall* dllGetClassObjectProc)(const IID&, const IID&,
        void**);
    typedef HRESULT (__stdcall* dllCanUnloadNowProc)();
    do 
    {
        dllGetClassObjectProc getObjProc =
            reinterpret_cast<dllGetClassObjectProc>(GetProcAddress(
            m, "DllGetClassObject"));
        dllCanUnloadNowProc canUnloadProc =
            reinterpret_cast<dllCanUnloadNowProc>(GetProcAddress(
            m, "DllCanUnloadNow"));
        if (!getObjProc || !canUnloadProc)
            break;

        intrusive_ptr<IClassFactory> factory;
        r = getObjProc(CLSID_WavDest, IID_IClassFactory,
            reinterpret_cast<void**>(&factory));
        if (SUCCEEDED(r))
            r = factory->CreateInstance(NULL, IID_IBaseFilter, (void**)&m_aviMuxer);

        if (FAILED(r))
        {
            return false;
        }
    } while (0);


//     r = CoCreateInstance(CLSID_WavDest, NULL, CLSCTX_INPROC,
//         IID_IBaseFilter, (void **)&m_aviMuxer);
    if (FAILED(r))
        return false;

    r = m_mixGraphBuilder->AddFilter(m_aviMuxer, L"wav MUX");
    if (FAILED(r))
        return false;

    intrusive_ptr<IPin> muxerInPin;
    r = GetPinByDirection(m_aviMuxer, 
        reinterpret_cast<IPin**>(&muxerInPin), 
        PINDIR_INPUT);

    if (FAILED(r))
        return false;

    r = m_mixGraphBuilder->ConnectDirect(sourceOutPin1.get(), muxerInPin.get(), NULL);
    if (FAILED(r))
        return false;


    intrusive_ptr<IFileSinkFilter2> pFileSink;
    r = CoCreateInstance(CLSID_FileWriter, NULL, CLSCTX_INPROC, IID_IFileSinkFilter,
        (void**)&pFileSink);

    if (FAILED(r))
        return false;

    // Get the file sink interface from the File Writer
    r = pFileSink->QueryInterface(IID_IBaseFilter, (void **)&m_pFileWriter);
    if (FAILED(r))
        return false;

    // Add the FileWriter filter to the graph
    r = m_mixGraphBuilder->AddFilter((IBaseFilter *)m_pFileWriter, L"File Writer");
    if (FAILED(r))
        return false;

    intrusive_ptr<IPin> writerInPin;
    r = GetPinByDirection(m_pFileWriter, 
        reinterpret_cast<IPin**>(&writerInPin), 
        PINDIR_INPUT);
    if (FAILED(r))
        return false;

    IFileSinkFilter* fileSinkFilter;
    r = m_pFileWriter->QueryInterface(
        IID_IFileSinkFilter, reinterpret_cast<void**>(&fileSinkFilter));
    if (FAILED(r))
        return false;


    r = fileSinkFilter->SetFileName(L"D:\\1111.wav", NULL);
    if (FAILED(r))
        return false;

    intrusive_ptr<IPin> muxerOutPin;
    r = GetPinByDirection(m_aviMuxer, 
        reinterpret_cast<IPin**>(&muxerOutPin), 
        PINDIR_OUTPUT);
    if (FAILED(r))
        return false;

    r = m_mixGraphBuilder->ConnectDirect(muxerOutPin.get(), writerInPin.get(), NULL);
    if (FAILED(r))
        return false;


    return true;
}


void CKTVDlg::OnBnClickedButtonSet()
{
    // TODO: 在此添加控件通知处理程序代码
    if (!buildMicrophoneGraph(m_captureFilter, m_capturePin))
        PostQuitMessage(0);

    if (!buildMixGraph(m_mixFilter, m_mixPin))
        PostQuitMessage(0); 
    GetDlgItem(IDC_BUTTON_SET)->EnableWindow(FALSE);
}

void CKTVDlg::OnCbnSelchangeCombo1()
{
    // TODO: 在此添加控件通知处理程序代码
    m_captureFilter = m_microphoneCombox.GetCurSel();
    m_microphonePinCombox.ResetContent();

    for (int i = 0; i < m_captureFilterVec[m_captureFilter].PinVec.size(); i++)
    {
        m_microphonePinCombox.AddString(m_captureFilterVec[m_captureFilter].PinVec[i].c_str());
    }
  
    m_microphonePinCombox.SetCurSel(0);

}

void CKTVDlg::OnCbnSelchangeCombo2()
{
    // TODO: 在此添加控件通知处理程序代码
    m_capturePin = m_microphonePinCombox.GetCurSel();
}

void CKTVDlg::OnCbnSelchangeCombo3()
{
    // TODO: 在此添加控件通知处理程序代码
    m_mixFilter = m_mixCaptureCombox.GetCurSel();

    m_mixPinCombox.ResetContent();
    for (int i = 0; i < m_captureFilterVec[m_mixFilter].PinVec.size(); i++)
    {
        m_mixPinCombox.AddString(m_captureFilterVec[m_mixFilter].PinVec[i].c_str());
    }

    m_mixPinCombox.SetCurSel(0);


}

void CKTVDlg::OnCbnSelchangeCombo4()
{
    // TODO: 在此添加控件通知处理程序代码
    m_mixPin = m_mixPinCombox.GetCurSel();
}
