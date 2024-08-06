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

constexpr UINT c_idTrayIcon = 1;
constexpr UINT WMU_TRAYNOTIFY = WM_USER + 111;

static const UINT c_msgTaskbarCreated = RegisterWindowMessageW(L"TaskbarCreated");
static const WCHAR c_szWndClass[] = L"HyperVTray_hidden_window";

static const WCHAR c_szTip[] = L"Tooltip text";

static HINSTANCE s_hinst = 0;
static HWND s_hwndMain = 0;
static HICON s_hicon = 0;
static bool s_fInTray = false;

enum class MenuMode { Watching, LDown, Cancelled };

static HMENU s_hmenu = 0;
static std::vector<SPI<IWbemClassObject>> s_vms;
static INT s_menuSelectIndex = -1;
static INT s_menuDownIndex = -1;
static MenuMode s_menuMode = MenuMode::Watching;
static RECT s_menuItemRect;

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

#if 0
namespace Hyper_V_Manager
{
    /// <inheritdoc />
    /// <summary>
    /// Main Form
    /// </summary>
    public partial class Form1 : Form
    {
        private readonly Timer _timer = new Timer();
        private readonly Dictionary<string, string> _changingVMs = new Dictionary<string, string>();

        /// <summary>
        /// Form load event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form1_Load(object sender, EventArgs e)
        {
            _timer.Elapsed += TimerElapsed;
            _timer.Interval = 4500;
            BuildContextMenu();
        }

        /// <summary>
        /// Time elapsed event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TimerElapsed(object sender, ElapsedEventArgs e)
        {
            if(!contextMenuStrip1.Visible)
                UpdateBalloontip();
        }

        /// <summary>
        /// Update Balloon Tooltip on VM State Change
        /// </summary>
        private void UpdateBalloontip()
        {
            var localVMs = GetVMs();

            foreach (var vm in localVMs)
            {
                if (!_changingVMs.ContainsKey(vm["ElementName"].ToString())) continue;
                var initvmBalloonState = _changingVMs[vm["ElementName"].ToString()];
                var vmState = (VmState) Convert.ToInt32(vm["EnabledState"]);
                var currentBalloonState = vmState.ToString();

                if (initvmBalloonState != currentBalloonState)
                {
                    notifyIcon1.ShowBalloonTip(4000, "VM State Changed", vm["ElementName"] + " " + currentBalloonState, ToolTipIcon.Info);
                    _changingVMs[vm["ElementName"].ToString()] = currentBalloonState;
                }
                else if (vmState == VmState.Running || vmState == VmState.Stopped || vmState == VmState.Paused || vmState == VmState.Saved)
                    _changingVMs.Remove(vm["ElementName"].ToString());
                else if (_changingVMs.Count <= 0)
                    _timer.Enabled = false;
            }
        }
    }
}
#endif

enum class VmOp { Connect, Start, Stop, ShutDown, Save, Pause };

static void AppendStateString(std::wstring& inout, VmState state)
{
    static const struct { VmState state; LPCWSTR text; } c_states[] =
    {
        { VmState::Unknown,     L"Unknown" },
        { VmState::Other,       L"Other" },
        { VmState::Running,     L"Running" },
        { VmState::Stopped,     L"Stopped" },
        { VmState::ShutDown,    L"ShutDown" },
        { VmState::Saved,       L"Saved" },
        { VmState::Paused,      L"Paused" },
        { VmState::Starting,    L"Starting" },
        { VmState::Saving,      L"Saving" },
        { VmState::Stopping,    L"Stopping" },
        { VmState::Pausing,     L"Pausing" },
        { VmState::Resuming,    L"Resuming" },
    };

    for (const auto& s : c_states)
    {
        if (s.state == state)
        {
            inout.append(L"  [");
            inout.append(s.text);
            inout.append(L"]");
            return;
        }
    }

    if (ULONG(state))
    {
        WCHAR unknown[64];
        swprintf_s(unknown, L"  [%d]", state);
        inout.append(unknown);
        return;
    }
}

static DWORD EnableFlags(bool enable)
{
    return enable ? MF_ENABLED : MF_DISABLED;
}

static HMENU BuildContextMenu(const std::vector<SPI<IWbemClassObject>> vms)
{
    HMENU hmenu = CreatePopupMenu();
    if (hmenu)
    {
        std::wstring name;
        ULONG state;
        for (UINT i = 0; i < vms.size(); ++i)
        {
            if (!GetStringProp(vms[i], L"ElementName", name) ||
                !GetIntegerProp(vms[i], L"EnabledState", state))
                continue;

            VmState vmstate = VmState(state);

            AppendStateString(name, vmstate);

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
            AppendMenuW(hmenuSub, MF_STRING, idmBase + WORD(VmOp::Connect), L"Connect");
            AppendMenuW(hmenuSub, MF_STRING|EnableFlags(enableStart), idmBase + WORD(VmOp::Start), L"Start");
            AppendMenuW(hmenuSub, MF_STRING|EnableFlags(enableStop), idmBase + WORD(VmOp::Stop), L"Stop");
            AppendMenuW(hmenuSub, MF_STRING, idmBase + WORD(VmOp::ShutDown), L"ShutDown");
            AppendMenuW(hmenuSub, MF_STRING|EnableFlags(enableSave), idmBase + WORD(VmOp::Save), L"Save State");
            AppendMenuW(hmenuSub, MF_STRING|EnableFlags(enablePause), idmBase + WORD(VmOp::Pause), L"Pause");

            MENUITEMINFOW mii = { sizeof(mii) };
            mii.fMask = MIIM_SUBMENU;
            mii.hSubMenu = hmenuSub;
            SetMenuItemInfoW(hmenu, idmPopup, false, &mii);
        }
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
            IWbemClassObject* pObject = s_vms[index];
            switch (VmOp((id - IDM_FIRSTVM) % 10))
            {
                case VmOp::Connect:
                    VmConnect(pObject);
                    break;
                case VmOp::Start:
                    ChangeVmState(pObject, VmState::Running);
                    break;
                case VmOp::Stop:
                    ChangeVmState(pObject, VmState::Stopped);
                    break;
                case VmOp::ShutDown:
                    ChangeVmState(pObject, VmState::ShutDown);
                    break;
                case VmOp::Save:
                    ChangeVmState(pObject, VmState::Saved);
                    break;
                case VmOp::Pause:
                    ChangeVmState(pObject, VmState::Paused);
                    break;
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
                VmConnect(s_vms[s_menuDownIndex]);
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

    POINT pt;
    GetCursorPos(&pt);
    const UINT id = TrackPopupMenu(s_hmenu, TPM_LEFTALIGN|TPM_RIGHTBUTTON|TPM_RETURNCMD, pt.x, pt.y, 0, hwnd, NULL);

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

