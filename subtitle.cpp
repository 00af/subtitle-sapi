#define _WIN32_WINNT 0x0603
#include <windows.h>
#include <sapi.h>
#include <sphelper.h>

#pragma comment(lib, "sapi.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdi32.lib")

// 窗口句柄
HWND hWndSubtitle = NULL;
HWND hTextDisplay = NULL;
HFONT hFont = NULL;

// 语音识别接口
ISpRecognizer* g_pReco = NULL;
ISpRecoContext* g_pContext = NULL;
ISpRecoGrammar* g_pGrammar = NULL;

// 窗口消息回调
LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
        case WM_DESTROY:
            PostQuitMessage(0);
            break;
        default:
            return DefWindowProc(hWnd, msg, wParam, lParam);
    }
    return 0;
}

// 更新字幕文字
void SetSubtitleText(LPCWSTR text) {
    if (hTextDisplay) {
        SetWindowTextW(hTextDisplay, text);
    }
}

// 初始化 Win8.1 自带语音识别
HRESULT InitSpeechRecognition(HWND hWnd) {
    HRESULT hr;

    // 初始化COM
    hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    if (FAILED(hr)) return hr;

    // 创建识别引擎
    hr = CoCreateInstance(CLSID_SpSharedRecognizer, NULL, CLSCTX_ALL, IID_ISpRecognizer, (void**)&g_pReco);
    if (FAILED(hr)) return hr;

    // 创建上下文
    hr = g_pReco->CreateRecoContext(&g_pContext);
    if (FAILED(hr)) return hr;

    // 设置窗口接收消息
    hr = g_pContext->SetMessageWindow(hWnd, WM_USER + 100);
    if (FAILED(hr)) return hr;

    // 创建语法
    hr = g_pContext->CreateGrammar(1, &g_pGrammar);
    if (FAILED(hr)) return hr;

    // 加载 dictation 模式（自由听写）
    hr = g_pGrammar->LoadDictation(NULL, SPLO_STATIC);
    if (FAILED(hr)) return hr;

    // 启用听写
    hr = g_pGrammar->SetDictationState(SPRS_ACTIVE);
    if (FAILED(hr)) return hr;

    return S_OK;
}

// 处理识别结果
void ProcessRecoResult(LPARAM lParam) {
    SPEVENT eventItem;
    ULONG eventCount = 0;
    WCHAR resultBuf[1024] = { 0 };

    while (SUCCEEDED(g_pContext->GetEvents(1, &eventItem, &eventCount)) && eventCount > 0) {
        if (eventItem.eEventId == SPEI_RECOGNITION) {
            ISpRecoResult* pResult = (ISpRecoResult*)eventItem.lParam;
            pResult->GetText(SP_GETWHOLEPHRASE, SP_GETWHOLEPHRASE, TRUE, resultBuf, NULL);
            SetSubtitleText(resultBuf);
            pResult->Release();
        }
    }
}

// 主函数
int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrev, LPSTR cmd, int nShow) {
    // 1. 注册窗口
    WNDCLASSEX wc = { 0 };
    wc.cbSize = sizeof(WNDCLASSEX);
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = L"SubtitleClass";
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    RegisterClassEx(&wc);

    // 2. 创建悬浮字幕窗口（屏幕底部半透明）
    int scrW = GetSystemMetrics(SM_CXSCREEN);
    int winW = scrW - 200;
    int winH = 140;

    hWndSubtitle = CreateWindowExW(
        WS_EX_LAYERED | WS_EX_TRANSPARENT | WS_EX_TOPMOST | WS_EX_NOACTIVATE,
        L"SubtitleClass",
        L"Real-time Subtitle",
        WS_POPUP,
        (scrW - winW) / 2,
        GetSystemMetrics(SM_CYSCREEN) - winH - 40,
        winW, winH,
        NULL, NULL, hInstance, NULL
    );

    // 设置透明度
    SetLayeredWindowAttributes(hWndSubtitle, 0, 230, LWA_ALPHA);

    // 3. 创建文字显示控件
    hTextDisplay = CreateWindowExW(
        0, L"STATIC", L"请开始说话...",
        WS_CHILD | WS_VISIBLE | SS_CENTER,
        0, 0, winW, winH,
        hWndSubtitle, NULL, hInstance, NULL
    );

    // 4. 创建字体
    hFont = CreateFontW(
        56, 0, 0, 0, FW_NORMAL,
        FALSE, FALSE, FALSE,
        DEFAULT_CHARSET,
        OUT_DEFAULT_PRECIS,
        CLIP_DEFAULT_PRECIS,
        ANTIALIASED_QUALITY,
        DEFAULT_PITCH,
        L"Microsoft YaHei UI"
    );
    SendMessageW(hTextDisplay, WM_SETFONT, (WPARAM)hFont, TRUE);

    ShowWindow(hWndSubtitle, SW_SHOW);
    UpdateWindow(hWndSubtitle);

    // 5. 初始化语音识别
    InitSpeechRecognition(hWndSubtitle);

    // 6. 消息循环
    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        if (msg.message == WM_USER + 100) {
            ProcessRecoResult(msg.lParam);
        }
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    // 释放资源
    DeleteObject(hFont);
    if (g_pGrammar) g_pGrammar->Release();
    if (g_pContext) g_pContext->Release();
    if (g_pReco) g_pReco->Release();
    CoUninitialize();
    return 0;
}
