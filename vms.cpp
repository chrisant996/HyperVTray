// Copyright (c) 2024 Christopher Antos
// License: http://opensource.org/licenses/MIT

#include "main.h"
#include "vms.h"

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
    if (swprintf_s(command, L"\"%s\\System32\\vmconnect.exe\" localhost \"%s\"", appname, name.c_str()) < 0)
        return;

    const DWORD dwCreationFlags = 0;
    STARTUPINFO si = { sizeof(si) };
    PROCESS_INFORMATION pi = { 0 };
    if (!CreateProcessW(appname, command, 0, 0, false, dwCreationFlags, 0, 0, &si, &pi))
        return;

    CloseHandle(pi.hThread);
    CloseHandle(pi.hProcess);
}

static bool GetStringProp(IWbemClassObject* pObject, LPCWSTR propName, std::wstring& out)
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

void ChangeVmState(VmState requestedState, LPCWSTR vmName)
{
    auto vms = GetVirtualMachines();

    // TODO: _timer.Enabled = true;

    std::wstring name;
    for (size_t i = 0; i < vms.size(); ++i)
    {
        if (!GetStringProp(vms[i], L"ElementName", name))
            continue;
        if (wcscmp(vmName, name.c_str()) != 0)
            continue;

        // Set the state to unknown as we request the change.
        // TODO: _changingVMs[vm["ElementName"].ToString()] = VmState.Unknown.ToString();

        // TODO: handle response from request to change.
        // https://docs.microsoft.com/en-us/windows/desktop/hyperv_v2/requeststatechange-msvm-computersystem
        // TODO: vms[i]-> InvokeMethod("RequestStateChange", inParams, null);
    }
}

std::vector<SPI<IWbemClassObject>> GetVirtualMachines()
{
    HRESULT hr;
    std::vector<SPI<IWbemClassObject>> vms;

    {
        SPI<IWbemLocator> spLocator;
        hr = CoCreateInstance(CLSID_WbemLocator, 0, CLSCTX_INPROC_SERVER, IID_IWbemLocator, (void**)&spLocator);
        if (FAILED(hr))
            goto LOut;

        SPI<IWbemServices> spServices;
        hr = spLocator->ConnectServer(BSTR(L"ROOT\\DEFAULT"), 0, 0, 0, 0, 0, 0, &spServices);
        if (FAILED(hr))
            goto LOut;

        hr = CoSetProxyBlanket(spServices, RPC_C_AUTHN_WINNT, RPC_C_AUTHZ_NONE, 0, RPC_C_AUTHN_LEVEL_CALL, RPC_C_IMP_LEVEL_IMPERSONATE, 0, EOAC_NONE);
        if (FAILED(hr))
            goto LOut;

        // Enumerate the available virtual machines.

        SPI<IEnumWbemClassObject> spEnum;
        hr = spServices->ExecQuery(BSTR(L"WQL"), L"Select * From Msvm_ComputerSystem", WBEM_FLAG_RETURN_IMMEDIATELY|WBEM_FLAG_FORWARD_ONLY, 0, &spEnum);
        if (FAILED(hr))
            goto LOut;

        std::wstring caption;
        while (spEnum)
        {
            ULONG uReturned = 0;
            SPI<IWbemClassObject> spObject;
            hr = spEnum->Next(WBEM_INFINITE, 1, &spObject, &uReturned);
            if (FAILED(hr) || !uReturned)
                break;

            if (!GetStringProp(spObject, L"Caption", caption))
                continue;
            if (wcscmp(caption.c_str(), L"Virtual Machine") != 0)
                continue;

            vms.emplace_back(std::move(spObject));
        }
    }

LOut:
    return vms;
}
