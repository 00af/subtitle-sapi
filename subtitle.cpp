#define UNICODE
#define _UNICODE
#include <windows.h>
#include <sapi.h>
#include <sphelper.h>
#include <string>
#include <chrono>

#pragma comment(lib, "sapi.lib")

HWND hLabel;
std::chrono::steady_clock::time_point lastUpdate;

LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_LBUTTONDOWN:
        SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0);
        break;
    case WM_DESTROY:
        PostQuitMessage(0);
        break;
    }
    return DefWindowProc(hwnd, msg, wParam, lParam);
}

void UpdateSubtitle(const std::wstring& text) {
    SetWindowTextW(hLabel, text.c_str());
    ShowWindow(hLabel, SW_SHOW);
    lastUpdate = std::chrono::steady_clock::now();
}

DWORD WINAPI HideThread(LPVOID) {
    while (true) {
        auto now = std::chrono::steady_clock::now();
        if (std::chrono::duration_cast<std::chrono::seconds>(now - lastUpdate).count() > 3) {
            ShowWindow(hLabel, SW_HIDE);
        }
        Sleep(300);
    }
}

DWORD WINAPI SpeechThread(LPVOID) {
    ISpRecognizer* recognizer = nullptr;
    ISpRecoContext* context = nullptr;
    ISpRecoGrammar* grammar = nullptr;

    CoInitialize(NULL);

    CoCreateInstance(CLSID_SpInprocRecognizer, NULL, CLSCTX_ALL, IID_ISpRecognizer, (void**)&recognizer);
    recognizer->SetInput(NULL, TRUE);
    recognizer->CreateRecoContext(&context);
    context->CreateGrammar(1, &grammar);
    grammar->LoadDictation(NULL, SPLO_STATIC);
    grammar->SetDictationState(SPRS_ACTIVE);

    CSpEvent event;

    while (true) {
        while (event.GetFrom(context) == S_OK) {
            if (event.eEventId == SPEI_RECOGNITION) {
                ISpRecoResult* result = event.RecoResult();
                LPWSTR text;
                result->GetText(SP_GETWHOLEPHRASE, SP_GETWHOLEPHRASE, FALSE, &text, NULL);
                UpdateSubtitle(text);
                CoTaskMemFree(text);
            }
        }
        Sleep(50);
    }

    return 0;
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE, LPSTR, int) {
    const wchar_t CLASS_NAME[] = L"SubtitleWindow";

    WNDCLASS wc = {};
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = CLASS_NAME;
    RegisterClass(&wc);

    HWND hwnd = CreateWindowEx(
        WS_EX_TOPMOST | WS_EX_TOOLWINDOW,
        CLASS_NAME,
        L"",
        WS_POPUP,
        200, 200, 1200, 120,
        NULL, NULL, hInstance, NULL
    );

    hLabel = CreateWindow(
        L"STATIC",
        L"",
        WS_VISIBLE | WS_CHILD,
        0, 0, 1200, 120,
        hwnd, NULL, hInstance, NULL
    );

    HFONT font = CreateFont(
        50, 0, 0, 0, FW_BOLD,
        FALSE, FALSE, FALSE,
        DEFAULT_CHARSET,
        OUT_OUTLINE_PRECIS,
        CLIP_DEFAULT_PRECIS,
        CLEARTYPE_QUALITY,
        VARIABLE_PITCH,
        L"Microsoft YaHei"
    );
    SendMessage(hLabel, WM_SETFONT, (WPARAM)font, TRUE);

    SetWindowLong(hLabel, GWL_STYLE, SS_CENTER | WS_VISIBLE);

    ShowWindow(hwnd, SW_SHOW);

    lastUpdate = std::chrono::steady_clock::now();
    CreateThread(NULL, 0, HideThread, NULL, 0, NULL);
    CreateThread(NULL, 0, SpeechThread, NULL, 0, NULL);

    MSG msg = {};
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    return 0;
}
