// Derived from https://github.com/ysc3839/win32-darkmode (MIT License).
//
// This is just enough Dark Mode support to get the tray icon context menu to
// use Dark Mode when appropriate.

#pragma once

#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <uxtheme.h>
#include "DarkMode.h"

#ifndef LOAD_LIBRARY_SEARCH_SYSTEM32
// From libloaderapi.h -- this flag was introduced in Win7.
#define LOAD_LIBRARY_SEARCH_SYSTEM32        0x00000800
#endif

using fnRtlGetNtVersionNumbers = void (WINAPI *)(LPDWORD major, LPDWORD minor, LPDWORD build);
static DWORD s_buildNumber = 0;

// 1809 17763
using fnAllowDarkModeForApp = bool (WINAPI *)(bool allow); // ordinal 135, in 1809
static fnAllowDarkModeForApp _AllowDarkModeForApp = nullptr;

// 1903 18362
using fnSetPreferredAppMode = PreferredAppMode (WINAPI *)(PreferredAppMode appMode); // ordinal 135, in 1903
static fnSetPreferredAppMode _SetPreferredAppMode = nullptr;

constexpr bool CheckBuildNumber(DWORD buildNumber)
{
#if 0
    return (buildNumber == 17763 || // 1809
        buildNumber == 18362 || // 1903
        buildNumber == 18363); // 1909
#else
    return (buildNumber >= 18362);
#endif
}

void AllowDarkMode()
{
    auto RtlGetNtVersionNumbers = reinterpret_cast<fnRtlGetNtVersionNumbers>(GetProcAddress(GetModuleHandleW(L"ntdll.dll"), "RtlGetNtVersionNumbers"));
    if (!RtlGetNtVersionNumbers)
        return;

    DWORD major, minor;
    RtlGetNtVersionNumbers(&major, &minor, &s_buildNumber);
    s_buildNumber &= ~0xF0000000;
    if (major < 10)
        return;
    if (major == 10 && !CheckBuildNumber(s_buildNumber))
        return;

    HMODULE hUxtheme = LoadLibraryExW(L"uxtheme.dll", nullptr, LOAD_LIBRARY_SEARCH_SYSTEM32);
    if (!hUxtheme)
        return;

    auto ord135 = GetProcAddress(hUxtheme, MAKEINTRESOURCEA(135));
    if (!ord135)
        return;

    if (s_buildNumber < 18362)
    {
        _AllowDarkModeForApp = reinterpret_cast<fnAllowDarkModeForApp>(ord135);
        _AllowDarkModeForApp(true);
    }
    else
    {
        _SetPreferredAppMode = reinterpret_cast<fnSetPreferredAppMode>(ord135);
        _SetPreferredAppMode(PreferredAppMode::AllowDark);
    }
}
