# HyperVTray

This is written in C++ and based on [Hyper-V Manager](https://github.com/kariudo/Hyper-VManagerTray).

It adds an icon in the system tray which lets you right click and use a context menu to list or control the available Hyper-V VMs.

![image](https://raw.githubusercontent.com/chrisant996/HyperVTray/master/screenshot.png)

## Why was it created?

I got a new laptop and switched from Windows 10 to Windows 11, and found that Hyper-V Manager flashed a form on startup, and also required higher permissions.  I also wanted some other minor adjustments.

So, I ported it to C++ and made some minor improvements.

## Required permissions

If HypervTray isn't able to query and control Hyper-V (e.g. if the list of VMs is empty), then try adding your user account as a member of the "Performance Monitor Users" group.

1. Open the Windows Start menu.
2. Search for "Computer Management" and launch it.
3. Double click on "Local Users and Groups", then double click on "Groups".
4. Double click on "Performance Monitor Users", then click "Add" and add your user account.
5. Click the OK buttons to close the dialog boxes, then close the Computer Management window.
6. Then start HyperVTray, and it should be able to list and control the VMs.

## Building HyperVTray

HyperVTray uses [Premake](http://premake.github.io) to generate Visual Studio solutions. Note that Premake >= 5.0.0-beta4 is required.

1. Cd to your clone of hypervtray.
2. Run <code>premake5.exe <em>toolchain</em></code> (where <em>toolchain</em> is one of Premake's actions - see `premake5.exe --help`).
3. Build scripts will be generated in <code>.build\\<em>toolchain</em></code>. For example `.build\vs2019\hypervtray.sln`.
4. Call your toolchain of choice (Visual Studio, msbuild.exe, etc).

# Credits

- Hunter Horsman (kariudo), https://github.com/kariudo/Hyper-VManagerTray
- Andrew Lyakhov (andrew_lyakhov), https://www.codeproject.com/articles/84143/programmed-hyper-v-management
