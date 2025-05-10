// Copyright (c) 2024 Christopher Antos
// License: http://opensource.org/licenses/MIT

#include "main.h"
#include "vms.h"
#include "darkmode.h"
#include "res.h"
#include <windowsx.h>
#include <stdlib.h>
#include <shellapi.h>
#include <dwmapi.h>
#include <wbemidl.h>
#include <map>

static const WCHAR c_usage[] =
L"Usage:  HyperVTray [--help --nodarkmode]\n"
L"\n"
L"HyperVTray adds a system tray icon which you can right click on to manage Hyper-V virtual machines."
;

static const UINT c_msgTaskbarCreated = RegisterWindowMessageW(L"TaskbarCreated");
static const WCHAR c_szWndClass[] = L"HyperVTray_hidden_window";

static HINSTANCE s_hinst = 0;
static HWND s_hwndMain = 0;
static HICON s_hicon = 0;

// Tray icon management.

constexpr UINT c_idTrayIcon = 1;
constexpr UINT WMU_TRAYNOTIFY = WM_USER;
static const WCHAR c_szTip[] = L"Hyper-V management";
static bool s_isIconInstalled = false;

static bool TrayMessage(DWORD dwMessage, LPCWSTR title=nullptr, LPCWSTR message=nullptr, DWORD dwInfoFlags=NIIF_INFO)
{
    NOTIFYICONDATA  data = { sizeof(data) };
    data.hWnd               = s_hwndMain;
    data.uID                = c_idTrayIcon;
    data.uFlags             = NIF_MESSAGE;
    data.uCallbackMessage   = WMU_TRAYNOTIFY;

    if (s_hicon)
    {
        data.uFlags |= NIF_ICON;
        data.hIcon = s_hicon;
    }

    data.uFlags |= NIF_TIP;
    wcscpy_s(data.szTip, c_szTip);

    if (title && message)
    {
        data.uFlags = NIF_INFO;
        data.dwInfoFlags = dwInfoFlags;
        wcscpy_s(data.szInfo, message);
        wcscpy_s(data.szInfoTitle, title);
        data.uTimeout = 4 * 1000;
    }

    const DWORD tickBegin = GetTickCount();
    while (!Shell_NotifyIcon(dwMessage, &data))
    {
        if (GetTickCount() - tickBegin > 10 * 1000)
            return false;
        Sleep(500);
    }

    return true;
}

static void AddTrayIcon()
{
    assert(!s_isIconInstalled);

    TrayMessage(NIM_ADD);
    s_isIconInstalled = true;
}

static void UpdateTrayIcon(LPCWSTR title=nullptr, LPCWSTR message=nullptr, DWORD dwInfoFlags=NIIF_INFO)
{
    assert(s_isIconInstalled);

    TrayMessage(NIM_MODIFY, title, message, dwInfoFlags);
}

static void DeleteTrayIcon()
{
	if (s_isIconInstalled)
    {
        TrayMessage(NIM_DELETE);
        s_isIconInstalled = false;
    }
}

// Notifications.

constexpr UINT c_timerId = 99;
constexpr UINT c_timerFirstInterval = 500;
constexpr UINT c_timerRepeatInterval = 2500;

struct WatchForStateChanges
{
    WatchForStateChanges() = default;
    WatchForStateChanges(VmState _original, VmState _target)
        : original(_original)
        , seen(_original)
        , target(_target)
        , changed(false)
    {}

    VmState original = VmState::Unknown;
    VmState seen = VmState::Unknown;
    VmState target = VmState::Unknown;
    bool changed = false;
};

static std::map<std::wstring, WatchForStateChanges> s_watching;
static UINT s_nextInterval = c_timerFirstInterval;

static void DoNotifications()
{
    auto vms = GetVirtualMachines();

    std::wstring message;
    for (const auto& vm : vms)
    {
        auto& watching = s_watching.find(vm.name);
        if (watching == s_watching.end())
            continue;

        ULONG state;
        if (!GetIntegerProp(vm.vm, L"EnabledState", state))
            continue;

        auto& w = watching->second;
        const VmState oldState = w.seen;
        const VmState newState = VmState(state);

        bool doErase = false;

        if (oldState != newState)
        {
            s_nextInterval = c_timerRepeatInterval;
            w.seen = newState;
            w.changed = true;

            doErase = (newState == w.target);

            message = vm.name;
            AppendStateString(message, newState, false/*brackets*/);
            UpdateTrayIcon(L"VM State Changed", message.c_str());
        }
        else if (w.changed && (newState == VmState::Running ||
                               newState == VmState::Stopped ||
                               newState == VmState::Paused ||
                               newState == VmState::Saved))
        {
            doErase = true;
        }

        if (doErase)
        {
            s_watching.erase(vm.name);
            if (s_watching.empty())
                break;
        }
    }
}

// Context menu.

enum class VmOp { Connect, Start, Stop, ShutDown, Save, Pause };
enum class MenuMode { Watching, LDown, Cancelled };

static VirtualMachines s_vms;
static HMENU s_hmenu = 0;
static bool s_inContextMenu = false;
static INT s_menuSelectIndex = -1;
static INT s_menuDownIndex = -1;
static MenuMode s_menuMode = MenuMode::Watching;
static RECT s_menuItemRect;

static DWORD EnableFlags(bool enable)
{
    return enable ? MF_ENABLED : MF_DISABLED;
}

static HMENU BuildContextMenu(const VirtualMachines vms)
{
    HMENU hmenu = CreatePopupMenu();
    if (hmenu)
    {
        std::wstring name;
        ULONG state;
        for (UINT i = 0; i < vms.size(); ++i)
        {
            if (!GetIntegerProp(vms[i].vm, L"EnabledState", state))
                continue;

            VmState vmstate = VmState(state);

            name.clear();
            if (i + 1 <= 9)
            {
                WCHAR prefix[] = { '&', WCHAR('1' + i), ' ', '-', ' ', '\0' };
                name = prefix;
            }
            name += vms[i].name;
            AppendStateString(name, vmstate, true/*brackets*/);

            const UINT idmBase = IDM_FIRSTVM + (i * 10);
            const UINT idmPopup = idmBase + WORD(VmOp::Connect);
            AppendMenuW(hmenu, MF_POPUP, idmPopup, name.c_str());

            bool enableStart = true;
            bool enableStop = true;
            bool enableSave = true;
            bool enablePause = true;
            switch (vmstate)
            {
            case VmState::Running:
                enableStart = false;
                break;
            case VmState::Saved:
            case VmState::ShutDown:
            case VmState::Stopped:
                enableStop = false;
                enableSave = false;
                enablePause = false;
                break;
            case VmState::Paused:
                enablePause = false;
                break;
            }

            HMENU hmenuSub = CreatePopupMenu();
            AppendMenuW(hmenuSub, MF_STRING, idmBase + WORD(VmOp::Connect), L"&Connect");
            AppendMenuW(hmenuSub, MF_STRING|EnableFlags(enableStart), idmBase + WORD(VmOp::Start), L"Sta&rt");
            AppendMenuW(hmenuSub, MF_STRING|EnableFlags(enableStop), idmBase + WORD(VmOp::Stop), L"St&op");
            AppendMenuW(hmenuSub, MF_STRING, idmBase + WORD(VmOp::ShutDown), L"Shut&Down");
            AppendMenuW(hmenuSub, MF_STRING|EnableFlags(enableSave), idmBase + WORD(VmOp::Save), L"Sa&ve State");
            AppendMenuW(hmenuSub, MF_STRING|EnableFlags(enablePause), idmBase + WORD(VmOp::Pause), L"&Pause");

            MENUITEMINFOW mii = { sizeof(mii) };
            mii.fMask = MIIM_SUBMENU;
            mii.hSubMenu = hmenuSub;
            SetMenuItemInfoW(hmenu, idmPopup, false, &mii);
        }
        AppendMenuW(hmenu, MF_SEPARATOR, -1, L"");
        AppendMenuW(hmenu, 0, IDM_EXIT, L"E&xit");
    }

    return hmenu;
}

static void DoCommand(UINT id)
{
    if (id == IDM_EXIT)
    {
        if (s_hwndMain)
            DestroyWindow(s_hwndMain);
        PostQuitMessage(0);
    }
    else if (id >= IDM_FIRSTVM)
    {
        const UINT index = (id - IDM_FIRSTVM) / 10;
        if (index < s_vms.size())
        {
            const auto& vm = s_vms[index];

            const VmOp op = VmOp((id - IDM_FIRSTVM) % 10);
            VmState requestedState = VmState::Unknown;

            switch (VmOp((id - IDM_FIRSTVM) % 10))
            {
            case VmOp::Connect:     VmConnect(vm.vm); break;
            case VmOp::Start:       requestedState = VmState::Running; break;
            case VmOp::Stop:        requestedState = VmState::Stopped; break;
            case VmOp::ShutDown:    requestedState = VmState::ShutDown; break;
            case VmOp::Save:        requestedState = VmState::Saved; break;
            case VmOp::Pause:       requestedState = VmState::Paused; break;
            }

            if (requestedState != VmState::Unknown)
            {
                ULONG state;
                if (GetIntegerProp(vm.vm, L"EnabledState", state))
                {
                    s_watching[vm.name] = { VmState(state), requestedState };
                    s_nextInterval = c_timerFirstInterval;
                    SetTimer(s_hwndMain, c_timerId, s_nextInterval, 0);
                }

                ChangeVmState(vm.vm, requestedState);
            }
        }
    }
}

static void OnMenuSelect(WPARAM wParam, LPARAM lParam)
{
    s_menuSelectIndex = -1;

    const WORD flags = HIWORD(wParam);
    if (flags & MF_POPUP)
    {
        const UINT index = LOWORD(wParam);
        if (index < s_vms.size())
        {
            if (GetMenuItemRect(0, HMENU(lParam), index, &s_menuItemRect))
            {
                s_menuSelectIndex = index;
            }
        }
    }
}

static void OnEnterIdle(WPARAM wParam, LPARAM lParam)
{
    if (wParam == MSGF_MENU)
    {
        POINT pt;
        if (!GetCursorPos(&pt))
        {
LCancel:
            s_menuMode = MenuMode::Cancelled;
            s_menuSelectIndex = -1;
            s_menuDownIndex = -1;
            return;
        }

        const bool fDown = (GetKeyState(VK_LBUTTON) < 0);
        if (fDown)
        {
            if (s_menuMode == MenuMode::Cancelled ||
                (s_menuMode == MenuMode::LDown && s_menuDownIndex != s_menuSelectIndex) ||
                !PtInRect(&s_menuItemRect, pt))
                goto LCancel;
            s_menuDownIndex = s_menuSelectIndex;
            s_menuMode = MenuMode::LDown;
        }
        else if (s_menuSelectIndex < 0 || !PtInRect(&s_menuItemRect, pt))
        {
            s_menuMode = MenuMode::Watching;
            s_menuDownIndex = -1;
        }
        else if (s_menuMode == MenuMode::LDown && s_menuDownIndex == s_menuSelectIndex)
        {
            assert(s_menuDownIndex >= 0);
            SendMessage(s_hwndMain, WM_CANCELMODE, 0, 0);
            if (UINT(s_menuDownIndex) < s_vms.size())
                VmConnect(s_vms[s_menuDownIndex].vm);
            goto LCancel;
        }
    }
}

static void DoContextMenu(HWND hwnd)
{
    s_vms = GetVirtualMachines();

    s_hmenu = BuildContextMenu(s_vms);
    if (!s_hmenu)
        return;

    s_menuSelectIndex = -1;
    s_menuDownIndex = -1;
    s_menuMode = MenuMode::Watching;

    // Workaround:  due to a well-known issue in Windows, the menu won't
    // disappear unless the window is set as the foreground window.
    SetForegroundWindow(hwnd);

    s_inContextMenu = true;

    POINT pt;
    GetCursorPos(&pt);
    const UINT id = TrackPopupMenu(s_hmenu, TPM_LEFTALIGN|TPM_RIGHTBUTTON|TPM_RETURNCMD, pt.x, pt.y, 0, hwnd, NULL);

    s_inContextMenu = false;

    // Workaround:  due to a well-known issue in Windows, the menu won't
    // disappear correctly unless it is sent a message (WM_NULL is a nop).
    SendMessage(hwnd, WM_NULL, 0, 0);

    DestroyMenu(s_hmenu);
    s_hmenu = 0;

    DoCommand(id);

    s_vms.clear();
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
                DoContextMenu(hwnd);
                break;

            default:
                break;
            }
        }
        break;

    case WM_MENUSELECT:
        OnMenuSelect(wParam, lParam);
        break;
    case WM_ENTERIDLE:
        OnEnterIdle(wParam, lParam);
        break;

    case WM_TIMER:
        if (wParam == c_timerId)
        {
            if (!s_inContextMenu)
            {
                DoNotifications();
                SetTimer(hwnd, c_timerId, s_nextInterval, 0);
            }
            if (s_watching.empty())
                KillTimer(hwnd, c_timerId);
        }
        break;

    case WM_DESTROY:
        s_vms.clear();
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

// Hidden main window.

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
            MessageBoxW(0, c_usage, L"HyperVTray", MB_OK|MB_ICONINFORMATION);
            return 0;
        }

        if (_wcsicmp(argv[0], L"/nodarkmode") == 0 ||
            _wcsicmp(argv[0], L"--nodarkmode") == 0)
        {
            fAllowDarkMode = false;
        }
        else
        {
            WCHAR message[1024];
            swprintf_s(message, L"Unrecognized argument \"%s\".\n\n%s", argv[0], c_usage);
            MessageBox(0, message, L"HyperVTray", MB_OK|MB_ICONERROR);
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

