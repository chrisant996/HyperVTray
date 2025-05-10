// Copyright (c) 2024 Christopher Antos
// License: http://opensource.org/licenses/MIT

#pragma once

#include "main.h"
#include <wbemidl.h>
#include <vector>

enum class VmState
{
    // https://learn.microsoft.com/en-us/windows/win32/hyperv_v2/msvm-computersystem
    Unknown     = 0,
    Other       = 1,        // Corresponds to CIM_EnabledLogicalElement.EnabledState = Other.
    Running     = 2,        // Enabled.
    Stopped     = 3,        // Disabled.
    ShutDown    = 4,        // Valid in version 1 (V1) of Hyper-V only. The virtual machine is shutting down via the shutdown service. Corresponds to CIM_EnabledLogicalElement.EnabledState = ShuttingDown.
    Saved       = 6,        // Corresponds to CIM_EnabledLogicalElement.EnabledState = Enabled but offline.
    Test        = 7,
    Defer       = 8,
    Paused      = 9,        // Corresponds to CIM_EnabledLogicalElement.EnabledState = Quiesce, Enabled but paused.
    Starting    = 10,
    Reset       = 11,
    _Starting   = 32770,
    Saving      = 32773,
    Stopping    = 32774,
    Pausing     = 32776,
    Resuming    = 32777,
};

void LaunchManager(HWND hwnd);
void VmConnect(IWbemClassObject* pObject);
void ChangeVmState(IWbemClassObject* pObject, VmState requestedState);

struct VmEntry
{
    VmEntry(IWbemClassObject* pObject, LPCWSTR name) : vm(pObject), name(name) {}
    VmEntry(SPI<IWbemClassObject>&& spObject, std::wstring&& name) : vm(std::move(spObject)), name(std::move(name)) {}

    bool operator()(const VmEntry& a, const VmEntry& b) const { return _wcsicmp(a.name.c_str(), b.name.c_str()) < 0; }
    static bool less(const VmEntry& a, const VmEntry& b) { return _wcsicmp(a.name.c_str(), b.name.c_str()) < 0; }

    SPI<IWbemClassObject> vm;
    std::wstring name;
};
typedef std::vector<VmEntry> VirtualMachines;
VirtualMachines GetVirtualMachines();

void AppendStateString(std::wstring& inout, VmState state, bool brackets);
bool GetStringProp(IWbemClassObject* pObject, LPCWSTR propName, std::wstring& out);
bool GetIntegerProp(IWbemClassObject* pObject, LPCWSTR propName, ULONG& out);
