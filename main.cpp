// Copyright (c) 2024 Christopher Antos
// License: http://opensource.org/licenses/MIT

#include "main.h"
#include "darkmode.h"
#include "res.h"
#include <stdlib.h>
#include <shellapi.h>
#include <dwmapi.h>
#include <wbemidl.h>

constexpr UINT c_idTrayIcon = 1;
constexpr UINT WMU_TRAYNOTIFY = WM_USER + 111;

static const UINT c_msgTaskbarCreated = RegisterWindowMessageW(L"TaskbarCreated");
static const WCHAR c_szWndClass[] = L"HyperVTray_hidden_window";

static const WCHAR c_szTip[] = L"Tooltip text";

static HINSTANCE s_hinst = 0;
static HWND s_hwndMain = 0;
static HICON s_hicon = 0;
static bool s_fInTray = false;

static bool TrayMessage(DWORD dwMessage, UINT uID, HICON hIcon, LPCTSTR pszTip)
{
    NOTIFYICONDATA  tnd = { sizeof(tnd) };
    tnd.hWnd                = s_hwndMain;
    tnd.uID                 = uID;
    tnd.uFlags              = NIF_MESSAGE;
    tnd.uCallbackMessage    = WMU_TRAYNOTIFY;

    if (hIcon)
    {
        tnd.uFlags |= NIF_ICON;
        tnd.hIcon = hIcon;
    }

    if (pszTip)
    {
        tnd.uFlags |= NIF_TIP;
        lstrcpyn(tnd.szTip, pszTip, _countof(tnd.szTip) - 1);
    }

    const DWORD tickBegin = GetTickCount();
    while (!Shell_NotifyIcon(dwMessage, &tnd))
    {
        if (GetTickCount() - tickBegin > 30 * 1000)
            return false;
        Sleep(1 * 1000);
    }

    return true;
}

static void AddTrayIcon()
{
    assert(!s_fInTray);

    TrayMessage(NIM_ADD, c_idTrayIcon, s_hicon, c_szTip);
    s_fInTray = true;
}

#if 0
static void UpdateTrayIcon()
{
    assert(s_fInTray);

    TrayMessage(NIM_MODIFY, c_idTrayIcon, s_hicon, c_szTip);
}
#endif

static void DeleteTrayIcon()
{
	if (s_fInTray)
    {
        TrayMessage(NIM_DELETE, c_idTrayIcon, 0, 0);
        s_fInTray = false;
    }
}

static void DoContextMenu()
{
    HMENU hmenu = LoadMenu(s_hinst, MAKEINTRESOURCEW(IDR_CONTEXT_MENU));
    HMENU hmenuSub;
    UINT id;
    POINT pt;

    if (hmenu)
    {
        hmenuSub = GetSubMenu(hmenu, 0);

        // Workaround:  due to a well-known issue in Windows, the menu won't
        // disappear unless the window is set as the foreground window.
        SetForegroundWindow(s_hwndMain);

        GetCursorPos(&pt);
        id = TrackPopupMenu(hmenuSub,
                            TPM_LEFTALIGN|TPM_RIGHTBUTTON|TPM_RETURNCMD,
                            pt.x, pt.y, 0, s_hwndMain, NULL);

        // Workaround:  due to a well-known issue in Windows, the menu won't
        // disappear correctly unless it is sent a message (WM_NULL is a nop).
        SendMessage(s_hwndMain, WM_NULL, 0, 0);

        switch (id)
        {
        case IDM_EXIT:
            DestroyWindow(s_hwndMain);
            PostQuitMessage(0);
            break;

        case 0:
            break;
        }

        DestroyMenu(hmenu);
    }
}

static LRESULT CALLBACK HiddenWndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    switch (uMsg)
    {
    case WMU_TRAYNOTIFY:
        {
            switch (lParam)
            {
            case WM_LBUTTONDOWN:
                break;

            case WM_LBUTTONDBLCLK:
                break;

            case WM_RBUTTONUP:
                DoContextMenu();
                break;

            default:
                break;
            }
        }
        break;

    case WM_DESTROY:
        DeleteTrayIcon();
        s_hwndMain = 0;
        break;

    default:
        if (uMsg == c_msgTaskbarCreated)
        {
            AddTrayIcon();
            break;
        }
        return DefWindowProc(hwnd, uMsg, wParam, lParam);
    }

    return 0;
}

static bool Init()
{
    s_hicon = LoadIcon(s_hinst, MAKEINTRESOURCEW(IDI_MAIN));

    // Initialize the window.

    WNDCLASS wc;
    wc.style = CS_DBLCLKS;
    wc.lpfnWndProc = HiddenWndProc;
    wc.cbClsExtra = 0;
    wc.cbWndExtra = 0;
    wc.hInstance = s_hinst;
    wc.hIcon = s_hicon;
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_3DFACE+1);
    wc.lpszMenuName = NULL;
    wc.lpszClassName = c_szWndClass;

    if (!RegisterClass(&wc))
        return false;

    s_hwndMain = CreateWindow(c_szWndClass, c_szWndClass, 0, 0, 0, 0, 0, nullptr, 0, s_hinst, 0);
    if (!s_hwndMain)
        return false;

    // Initialize COM.

    HRESULT hr = CoInitializeEx(0, COINIT_MULTITHREADED);
    if (SUCCEEDED(hr))
        hr = CoInitializeSecurity(
            NULL,                        // Security descriptor
            -1,                          // COM negotiates authentication service
            NULL,                        // Authentication services
            NULL,                        // Reserved
            RPC_C_AUTHN_LEVEL_DEFAULT,   // Default authentication level for proxies
            RPC_C_IMP_LEVEL_IMPERSONATE, // Default Impersonation level for proxies
            NULL,                        // Authentication info
            EOAC_NONE,                   // Additional capabilities of the client or server
            NULL);                       // Reserved
    if (FAILED(hr))
    {
        // TODO: somehow report the error?
    }

    return true;
}

int PASCAL WinMain(HINSTANCE hinstCurrent, HINSTANCE /*hinstPrevious*/, LPSTR /*lpszCmdLine*/, int /*nCmdShow*/)
{
    s_hinst = hinstCurrent;

    int argc;
    const WCHAR **argv = const_cast<const WCHAR**>(CommandLineToArgvW(GetCommandLine(), &argc));

    if (argc)
    {
        --argc;
        ++argv;
    }

    // Parse options.

    bool fAllowDarkMode = true;

    while (argc)
    {
        if (_wcsicmp(argv[0], L"/?") == 0 ||
            _wcsicmp(argv[0], L"/h") == 0 ||
            _wcsicmp(argv[0], L"/help") == 0 ||
            _wcsicmp(argv[0], L"-?") == 0 ||
            _wcsicmp(argv[0], L"-h") == 0 ||
            _wcsicmp(argv[0], L"--help") == 0)
        {
            // TODO: show message box with help.
            return 0;
        }

        if (_wcsicmp(argv[0], L"/nodarkmode") == 0 ||
            _wcsicmp(argv[0], L"--nodarkmode") == 0)
        {
            fAllowDarkMode = false;
        }
        else
        {
            // TODO: unrecognized flag.
            return 1;
        }

        --argc, ++argv;
    }

    if (fAllowDarkMode)
        AllowDarkMode();

    // Use a mutex to ensure only one instance.

    const HANDLE hMutex = CreateMutex(0, false, L"HyperVTrayMutex_chrisant996");

    const DWORD dwMutexError = GetLastError();
    if (dwMutexError == ERROR_ALREADY_EXISTS)
    {
        ReleaseMutex(hMutex);
        return 0;
    }

    const DWORD dwWaitResult = WaitForSingleObject(hMutex, INFINITE);
    if (dwWaitResult != WAIT_OBJECT_0)
    {
        ReleaseMutex(hMutex);
        return 0;
    }

    // Initialize the app.

    MSG msg = { 0 };

    if (!Init())
    {
        msg.wParam = 1;
        goto LError;
    }

    AddTrayIcon();

    // Main message loop.

    if (s_hwndMain)
    {
        while (GetMessage(&msg, nullptr, 0, 0))
        {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }

    // Cleanup.

    {
        MSG tmp;
        do {} while(PeekMessage(&tmp, 0, WM_QUIT, WM_QUIT, PM_REMOVE));
    }

LError:
    ReleaseMutex(hMutex);

    return int(msg.wParam);
}

