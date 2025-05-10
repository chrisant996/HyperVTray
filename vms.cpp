// Copyright (c) 2024 Christopher Antos
// License: http://opensource.org/licenses/MIT

#include "main.h"
#include "vms.h"
#include <atlcomcli.h>
#include <algorithm>
#include <shellapi.h>

void LaunchManager(HWND hwnd)
{
    WCHAR system32[1024] = { 0 };
    DWORD dw = GetEnvironmentVariableW(L"SYSTEMROOT", system32, _countof(system32));
    if (!dw || dw > _countof(system32))
        return;
    if (wcscat_s(system32, L"\\System32"))
        return;

    WCHAR appname[1024];
    if (swprintf_s(appname, L"%s\\mmc.exe", system32) < 0)
        return;

    WCHAR args[1024];
    if (swprintf_s(args, L"\"%s\\virtmgmt.msc\"", system32) < 0)
        return;

    WCHAR cwd[1024] = { 0 };
    dw = GetEnvironmentVariableW(L"ProgramFiles", cwd, _countof(cwd));
    if (!dw || dw > _countof(cwd))
        return;
    if (wcscat_s(cwd, L"\\Hyper-V"))
        return;

    SHELLEXECUTEINFOW sei = { sizeof(sei) };
    sei.hwnd = hwnd;
    sei.fMask = SEE_MASK_FLAG_DDEWAIT|SEE_MASK_FLAG_NO_UI|SEE_MASK_NOCLOSEPROCESS;
    sei.lpVerb = L"runas";
    sei.lpFile = appname;
    sei.lpParameters = args;
    sei.lpDirectory = cwd;
    sei.nShow = SW_SHOWNORMAL;

    if (!ShellExecuteExW(&sei) || !sei.hProcess)
    {
#ifdef DEBUG
        const DWORD err = GetLastError();
        WCHAR message[1024] = { 0 };
        if (swprintf_s(message, L"Error Code %d (0x%X).", err, err) > 0)
            MessageBox(NULL, message, L"HyperVTray Error", MB_ICONERROR|MB_OK);
#endif
        return;
    }

    CloseHandle(sei.hProcess);
}

void VmConnect(IWbemClassObject* pObject)
{
    WCHAR appname[1024] = { 0 };
    DWORD dw = GetEnvironmentVariableW(L"SYSTEMROOT", appname, _countof(appname));
    if (!dw || dw > _countof(appname))
        return;
    if (wcscat_s(appname, L"\\System32\\vmconnect.exe"))
        return;

    std::wstring name;
    if (!GetStringProp(pObject, L"ElementName", name))
        return;

    WCHAR command[1024];
    if (swprintf_s(command, L"\"%s\" localhost \"%s\"", appname, name.c_str()) < 0)
        return;

    const DWORD dwCreationFlags = 0;
    STARTUPINFO si = { sizeof(si) };
    PROCESS_INFORMATION pi = { 0 };
    if (!CreateProcessW(appname, command, 0, 0, false, dwCreationFlags, 0, 0, &si, &pi))
    {
#ifdef DEBUG
        const DWORD err = GetLastError();
        WCHAR message[1024] = { 0 };
        if (swprintf_s(message, L"Error Code %d (0x%X).", err, err) > 0)
            MessageBox(NULL, message, L"HyperVTray Error", MB_ICONERROR|MB_OK);
#endif
        return;
    }

    CloseHandle(pi.hThread);
    CloseHandle(pi.hProcess);
}

static HRESULT GetShutdownComponent(IWbemServices* pServices, IWbemClassObject* pObject, IWbemClassObject** ppShutdownComponent)
{
    HRESULT hr;

    std::wstring path;
    if (!GetStringProp(pObject, L"__RELPATH", path))
        return E_FAIL;

    WCHAR query[1024];
    swprintf_s(query, L"associators of {%s} where AssocClass=Msvm_SystemDevice ResultClass=Msvm_ShutdownComponent", BSTR(path.c_str()));

    SPI<IEnumWbemClassObject> spEnum;
    const long flags = WBEM_FLAG_FORWARD_ONLY;
    hr = pServices->ExecQuery(BSTR(L"WQL"), BSTR(query), flags, 0, &spEnum);
    if (FAILED(hr))
        return hr;

    ULONG uReturned = 0;
    hr = spEnum->Next(WBEM_INFINITE, 1, ppShutdownComponent, &uReturned);
    if (FAILED(hr))
        return hr;
    if (!uReturned)
        return E_FAIL;

    return S_OK;
}

static HRESULT GetMethodParams(IWbemServices* pServices, LPCWSTR objectPath, LPCWSTR methodName, IWbemClassObject** ppInParams)
{
    HRESULT hr;

    SPI<IWbemClassObject> spClass;
    hr = pServices->GetObject(BSTR(objectPath), 0, 0, &spClass, 0);
    if (FAILED(hr))
        return hr;

    SPI<IWbemClassObject> spInParamsDefinition;
    hr = spClass->GetMethod(methodName, 0, &spInParamsDefinition, NULL);
    if (FAILED(hr))
        return hr;

    hr = spInParamsDefinition->SpawnInstance(0, ppInParams);
    if (FAILED(hr))
        return hr;

    return S_OK;
}

static HRESULT ExecMethod(IWbemServices* pServices, IWbemClassObject* pObject, LPCWSTR methodName, IWbemClassObject* pInParams, IWbemClassObject** ppOutParams)
{
    HRESULT hr;

    std::wstring path;
    if (!GetStringProp(pObject, L"__PATH", path))
        return E_FAIL;

    hr = pServices->ExecMethod(BSTR(path.c_str()), BSTR(methodName), 0, 0, pInParams, ppOutParams, 0);
    if (FAILED(hr))
        return hr;

    return S_OK;
}

void ChangeVmState(IWbemClassObject* pObject, VmState requestedState)
{
    auto vms = GetVirtualMachines();

    SPI<IWbemLocator> spLocator;
    if (FAILED(CoCreateInstance(CLSID_WbemLocator, 0, CLSCTX_INPROC_SERVER, IID_IWbemLocator, (void**)&spLocator)))
        return;

    SPI<IWbemServices> spServices;
    if (FAILED(spLocator->ConnectServer(BSTR(L"ROOT\\Virtualization\\V2"), 0, 0, 0, 0, 0, 0, &spServices)))
        return;

    SPI<IWbemClassObject> spInParams;
    if (requestedState == VmState::Stopped)
    {
        if (FAILED(GetMethodParams(spServices, L"Msvm_ShutdownComponent", L"InitiateShutdown", &spInParams)))
            return;

        if (FAILED(spInParams->Put(L"Force", 0, &CComVariant(true), 0)))
            return;
        if (FAILED(spInParams->Put(L"Reason", 0, &CComVariant(L"Shutdown"), 0)))
            return;

        SPI<IWbemClassObject> spShutdownComponent;
        if (FAILED(GetShutdownComponent(spServices, pObject, &spShutdownComponent)))
            return;
        if (FAILED(ExecMethod(spServices, spShutdownComponent, L"InitiateShutdown", spInParams, 0)))
            return;
    }
    else
    {
        if (FAILED(GetMethodParams(spServices, L"Msvm_ComputerSystem", L"RequestStateChange", &spInParams)))
            return;

        VARIANT vt;
        VariantInit(&vt);
        vt.vt = VT_I4;
        vt.iVal = INT(requestedState);
        if (FAILED(spInParams->Put(L"RequestedState", 0, &vt, 0)))
        {
            VariantClear(&vt);
            return;
        }
        VariantClear(&vt);

        SPI<IWbemClassObject> spOutParams;
        if (FAILED(ExecMethod(spServices, pObject, L"RequestStateChange", spInParams, &spOutParams)))
            return;

        // TODO: handle response from request to change.
        // https://docs.microsoft.com/en-us/windows/desktop/hyperv_v2/requeststatechange-msvm-computersystem
    }
}

VirtualMachines GetVirtualMachines()
{
    HRESULT hr;
    VirtualMachines vms;

    {
        SPI<IWbemLocator> spLocator;
        hr = CoCreateInstance(CLSID_WbemLocator, 0, CLSCTX_INPROC_SERVER, IID_IWbemLocator, (void**)&spLocator);
        if (FAILED(hr))
            goto LOut;

        SPI<IWbemServices> spServices;
        hr = spLocator->ConnectServer(BSTR(L"ROOT\\Virtualization\\V2"), 0, 0, 0, 0, 0, 0, &spServices);
        if (FAILED(hr))
            goto LOut;

        hr = CoSetProxyBlanket(spServices, RPC_C_AUTHN_WINNT, RPC_C_AUTHZ_NONE, 0, RPC_C_AUTHN_LEVEL_CALL, RPC_C_IMP_LEVEL_IMPERSONATE, 0, EOAC_NONE);
        if (FAILED(hr))
            goto LOut;

        // Enumerate the available virtual machines.

        SPI<IEnumWbemClassObject> spEnum;
        const long flags = WBEM_FLAG_FORWARD_ONLY;//WBEM_FLAG_RETURN_IMMEDIATELY|WBEM_FLAG_FORWARD_ONLY;
        hr = spServices->ExecQuery(BSTR(L"WQL"), L"SELECT * FROM Msvm_ComputerSystem WHERE Caption=\"Virtual Machine\"", flags, 0, &spEnum);
        if (FAILED(hr))
            goto LOut;

        while (spEnum)
        {
            ULONG uReturned = 0;
            SPI<IWbemClassObject> spObject;
            hr = spEnum->Next(WBEM_INFINITE, 1, &spObject, &uReturned);
            if (FAILED(hr) || !uReturned)
                break;

            std::wstring name;
            if (!GetStringProp(spObject, L"ElementName", name))
                continue;

            vms.emplace_back(std::move(spObject), std::move(name));
        }
    }

    std::sort(vms.begin(), vms.end(), &VmEntry::less);

LOut:
    return vms;
}

void AppendStateString(std::wstring& inout, VmState state, bool brackets)
{
    static const struct { VmState state; LPCWSTR text; } c_states[] =
    {
        { VmState::Unknown,     L"Unknown" },
        { VmState::Other,       L"Other" },
        { VmState::Running,     L"Running" },
        { VmState::Stopped,     L"Stopped" },
        { VmState::ShutDown,    L"ShutDown" },
        { VmState::Saved,       L"Saved" },
        { VmState::Test,        L"Test" },
        { VmState::Defer,       L"Defer" },
        { VmState::Paused,      L"Paused" },
        { VmState::Starting,    L"Starting" },
        { VmState::Reset,       L"Reset" },
        { VmState::_Starting,   L"Starting" },
        { VmState::Saving,      L"Saving" },
        { VmState::Stopping,    L"Stopping" },
        { VmState::Pausing,     L"Pausing" },
        { VmState::Resuming,    L"Resuming" },
    };

    for (const auto& s : c_states)
    {
        if (s.state == state)
        {
            if (!inout.empty())
                inout.append(brackets ? L"  [" : L" ");
            else if (brackets)
                inout.append(L"[");
            inout.append(s.text);
            if (brackets)
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

bool GetStringProp(IWbemClassObject* pObject, LPCWSTR propName, std::wstring& out)
{
    bool ok = false;

    out.clear();

    VARIANT vt;
    VariantInit(&vt);

    HRESULT hr = pObject->Get(propName, 0, &vt, 0, 0);
    if (SUCCEEDED(hr) && (V_VT(&vt) & VT_TYPEMASK) == VT_BSTR)
    {
        out = V_BSTR(&vt);
        ok = true;
    }

    VariantClear(&vt);
    return ok;
}

bool GetIntegerProp(IWbemClassObject* pObject, LPCWSTR propName, ULONG& out)
{
    bool ok = false;

    out = 0;

    VARIANT vt;
    VariantInit(&vt);

    HRESULT hr = pObject->Get(propName, 0, &vt, 0, 0);
    if (SUCCEEDED(hr))
    {
        switch (V_VT(&vt) & VT_TYPEMASK)
        {
        case VT_I1:
        case VT_UI1:
            out = V_UI1(&vt);
            ok = true;
            break;
        case VT_I2:
        case VT_UI2:
            out = V_UI2(&vt);
            ok = true;
            break;
        case VT_I4:
        case VT_UI4:
        case VT_INT:
        case VT_UINT:
            out = V_UI4(&vt);
            ok = true;
            break;
        }
    }

    VariantClear(&vt);
    return ok;
}
