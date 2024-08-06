# HyperVTray

This is written in C++ and based on [Hyper-V Manager](https://github.com/kariudo/Hyper-VManagerTray).

## Why was it created?

I got a new laptop and switched from Windows 10 to Windows 11, and found that Hyper-V Manager flashed a form on startup, and also required higher permissions.  I also wanted some other minor adjustments.

So, I ported it to C++ and made some minor improvements.

## Required permissions

The permissions issue can be solved by **_TBD: fill in the steps_**.

## Building HyperVTray

HyperVTray uses [Premake](http://premake.github.io) to generate Visual Studio solutions. Note that Premake >= 5.0.0-beta1 is required.

1. Cd to your clone of hypervtray.
2. Run <code>premake5.exe <em>toolchain</em></code> (where <em>toolchain</em> is one of Premake's actions - see `premake5.exe --help`).
3. Build scripts will be generated in <code>.build\\<em>toolchain</em></code>. For example `.build\vs2019\hypervtray.sln`.
4. Call your toolchain of choice (Visual Studio, msbuild.exe, etc).

# Credits

https://github.com/kariudo/Hyper-VManagerTray
https://www.codeproject.com/articles/84143/programmed-hyper-v-management
