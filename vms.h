// Copyright (c) 2024 Christopher Antos
// License: http://opensource.org/licenses/MIT

#pragma once

#include "main.h"
#include <wbemidl.h>
#include <vector>

enum class VmState
{
    Unknown     = 0,
    Other       = 1,        // Corresponds to CIM_EnabledLogicalElement.EnabledState = Other.
    Running     = 2,        // Enabled.
    Stopped     = 3,        // Disabled.
    ShutDown    = 4,        // Valid in version 1 (V1) of Hyper-V only. The virtual machine is shutting down via the shutdown service. Corresponds to CIM_EnabledLogicalElement.EnabledState = ShuttingDown.
    Saved       = 6,        // Corresponds to CIM_EnabledLogicalElement.EnabledState = Enabled but offline.
    Paused      = 9,        // Corresponds to CIM_EnabledLogicalElement.EnabledState = Quiesce, Enabled but paused.
    Starting    = 32770,
    Saving      = 32773,
    Stopping    = 32774,
    Pausing     = 32776,
    Resuming    = 32777,
};

void VmConnect(IWbemClassObject* pObject);
void ChangeVmState(LPCWSTR requestedState, LPCWSTR vmName);
std::vector<SPI<IWbemClassObject>> GetVirtualMachines();
