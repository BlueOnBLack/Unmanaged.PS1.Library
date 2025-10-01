# PS1.NativeInterop
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)]()
[![Build Status](https://img.shields.io/badge/build-passing-brightgreen.svg)]()
[![PowerShell Gallery](https://img.shields.io/badge/PowerShell%20Module-PS1%20NativeInterop-blue.svg)]()

> Managed COM / Win32 API bridge and utilities for PowerShell (PS1).  
> High-level wrapper for advanced native interoperability and system inspection — intended for legitimate administrative, defensive, and research use.

## Table of contents
- [About](#about)
- [Key capabilities](#key-capabilities)
- [Legitimate use cases](#legitimate-use-cases)
- [Quick start](#quick-start)
- [Examples (high level)](#examples-high-level)
- [Security & Responsible Use](#security--responsible-use)
- [Contributing](#contributing)
- [License](#license)
- [Credits & references](#credits--references)

## About
`PS1.NativeInterop` is a PowerShell-focused library that exposes a set of managed wrappers and helpers to interact with native Windows APIs, COM interfaces, and low-level process information. The project is intended to aid system administrators, security researchers, and developers who need to inspect, diagnose, or automate system-level tasks from PowerShell.

> **Important:** This repository contains tools that may be used for intrusive operations. The maintainers refuse support for malicious usage. See [Security & Responsible Use](#security--responsible-use).

## Key capabilities
> *Listed here at a high level. Implementation details that facilitate exploitation or evasion are intentionally omitted.*

- Enum & dynamically load DLLs and resolve exported functions.
- Managed COM wrappers and simplified access to COM interfaces from PowerShell.
- Win32 API interop helpers with configurable character set handling.
- Token privilege inspection and utilities for querying/enumerating privileges.
- Process enumeration and querying of basic process information (PEB, name, PID).
- Dynamic lookup and invocation of API functions and COM methods.
- Parse and format error results from `HRESULT`, Win32 error codes, and `NTSTATUS`.
- Facilities to obtain handles to processes and services by ID (for legitimate admin tasks).
- Helpers to allocate, initialize and free native `STRING` / `UNICODE_STRING` structures.
- Generic handle/resource cleanup helpers (handles, global memory, heaps).
- Utilities to detect group membership (Administrator) and common system accounts.
- Advanced forensic/research-only features (documented as research): low-level process introspection and syscall metadata extraction.

## Legitimate use cases
- Administrative scripting that needs to call a specific native API not exposed directly by PowerShell.
- Forensic and incident response tooling that inspects process internals and gathers diagnostic data.
- Interop test harnesses for native libraries and COM objects when building Windows integrations.
- Security research and red-team/blue-team lab work, conducted in controlled environments with explicit permission.

## Quick start
> These are non-actionable installation hints. See the module manifest in `./src` for real install steps.

## Code samples
Below are quick, high‑level samples showing the module's call patterns. These examples are non‑destructive and intended for documentation/demo use only.

```powershell

Clear-Host
write-host

# COM: show product key UI (no parameters)
Use-ComInterface `
    -CLSID "17CCA47D-DAE5-4E4A-AC42-CC54E28F334A" `
    -IID  "f2dcb80d-0670-44bc-9002-cd18688730af" `
    -Index 3 `
    -Name  "ShowProductKeyUI" `
    -Return "void"

# Unmanaged DLL: Beep (kernel32.dll)
Invoke-UnmanagedMethod `
    -Dll      "kernel32.dll" `
    -Function "Beep" `
    -Return   "bool" `
    -Params   "uint dwFreq, uint dwDuration" `
    -Values   @(750, 300)  # 750 Hz beep for 300 ms

Write-Host
$hProc = [IntPtr]::Zero
$hProcNext = [IntPtr]::Zero
$ret = Invoke-UnmanagedMethod `
    -Dll NTDLL `
    -Function ZwGetNextProcess `
    -Values @($hProc, [UInt32]0x02000000, [UInt32]0x00, [UInt32]0x00, ([ref]$hProcNext)) `
    -Mode Allocate -SysCall
write-host "NtGetNextProcess Test: $ret"
write-host "hProcNext Value :$hProcNext"

Write-Host
$hThread = [IntPtr]::Zero
$hThreadNext = [IntPtr]::Zero
$ret = Invoke-UnmanagedMethod `
    -Dll NTDLL `
    -Function ZwGetNextThread `
    -Values @([IntPtr]::new(-1), $hThread, 0x0040, 0x00, 0x00, ([ref]$hThreadNext)) `
    -Mode Protect -SysCall
write-host "NtGetNextThread Test: $ret"
write-host "hThreadNext Value :$hThreadNext"

Write-Host
$ret = Invoke-UnmanagedMethod `
    -Dll NTDLL `
    -Function NtClose `
    -Values @([IntPtr]$hProcNext) `
    -Mode Allocate -SysCall
write-host "NtClose Test: $ret"

Write-Host
$FileHandle = [IntPtr]::Zero
$IoStatusBlock    = New-IntPtr -Size 16
$ObjectAttributes = New-IntPtr -Size 48 -WriteSizeAtZero
$filePath = ("\??\{0}\test.txt" -f [Environment]::GetFolderPath('Desktop'))
$ObjectName = Init-NativeString -Encoding Unicode -Value $filePath
[Marshal]::WriteIntPtr($ObjectAttributes, 0x10, $ObjectName)
[Marshal]::WriteInt32($ObjectAttributes,  0x18, 0x40)
$ret = Invoke-UnmanagedMethod `
    -Dll NTDLL `
    -Function NtCreateFile `
    -Values @(
        ([ref]$FileHandle),   # OUT HANDLE
        0x40100080,           # DesiredAccess (GENERIC_WRITE | SYNCHRONIZE | FILE_WRITE_DATA)
        $ObjectAttributes,    # POBJECT_ATTRIBUTES
        $IoStatusBlock,       # PIO_STATUS_BLOCK
        [IntPtr]::Zero,       # AllocationSize
        0x80,                 # FileAttributes (FILE_ATTRIBUTE_NORMAL)
        0x07,                 # ShareAccess (read|write|delete)
        0x5,                  # CreateDisposition (FILE_OVERWRITE_IF)
        0x20,                 # CreateOptions (FILE_NON_DIRECTORY_FILE)
        [IntPtr]::Zero,       # EaBuffer
        0x00                  # EaLength
    ) `
    -Mode Protect -SysCall
Free-NativeString -StringPtr $ObjectName
write-host "NtCreateFile Test: $ret"

