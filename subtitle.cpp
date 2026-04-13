#define _WIN32_WINNT 0x0603
#include <windows.h>
#include <sapi.h>
#include <sphelper.h>

#pragma comment(lib, "sapi.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdi32.lib")

HWND hWndSubtitle = NULL;
HWND hTextDisplay = NULL;
HFONT hFont = NULL;

ISpRecognizer* g_pReco = NULL;
ISpRecoContext* g_pContext = NULL;
ISpRecoGrammar* g_pGrammar = NULL;

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

void SetSubtitleText(LPCWSTR text) {
    if (hTextDisplay) {
        SetWindowTextW(hTextDisplay, text);
    }
}

HRESULT InitSpeechRecognition(HWND hWnd) {
    HRESULT hr;
    hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    if (FAILED(hr)) return hr;

    hr = CoCreateInstance(CLSID_SpSharedRecognizer, NULL, CLSCTX_ALL, IID_ISpRecognizer, (void**)&g_pReco);
    if (FAILED(hr)) return hr;

    hr = g_pReco->CreateRecoContext(&g_pContext);
    if (FAILED(hr)) return hr;

    hr = g_pContext->CreateGrammar(1, &g_pGrammar);
    if (FAILED(hr)) return hr;

    hr = g_pGrammar->LoadDictation(NULL, SPLO_STATIC);
    if (FAILED(hr)) return hr;

    hr = g_pGrammar->SetDictationState(SPRS_ACTIVE);
    if (FAILED(hr)) return hr;

    return S_OK;
}

void ProcessReco() {
    SPEVENT eventItem;
    ULONG fetched = 0;

    while (g_pContext->GetEvents(1, &eventItem, &fetched) == S_OK && fetched > 0) {
        if (eventItem.eEventId == SPEI_RECOGNITION) {
            ISpRecoResult* pResult = (ISpRecoResult*)eventItem.lParam;
            LPWSTR pText = NULL;
            pResult->GetText(SP_GETWHOLEPHRASE, SP_GETWHOLEPHRASE, TRUE, &pText, NULL);
            SetSubtitleText(pText);
            CoTaskMemFree(pText);
            pResult->Release();
        }
    }
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrev, LPSTR cmd, int nShow) {
    WNDCLASSEX wc = { 0 };
    wc.cbSize = sizeof(WNDCLASSEX);
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = L"SubtitleClass";
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    RegisterClassEx(&wc);

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

    SetLayeredWindowAttributes(hWndSubtitle, 0, 230, LWA_ALPHA);

    hTextDisplay = CreateWindowExW(
        0, L"STATIC", L"请说话...",
        WS_CHILD | WS_VISIBLE | SS_CENTER,
        0, 0, winW, winH,
        hWndSubtitle, NULL, hInstance, NULL
    );

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

    InitSpeechRecognition(hWndSubtitle);

    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        ProcessReco();
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    DeleteObject(hFont);
    if (g_pGrammar) g_pGrammar->Release();
    if (g_pContext) g_pContext->Release();
    if (g_pReco) g_pReco->Release();
    CoUninitialize();
    return 0;
}
