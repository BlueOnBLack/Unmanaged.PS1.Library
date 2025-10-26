# PS1.NativeInterop
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)]()
[![Build Status](https://img.shields.io/badge/build-passing-brightgreen.svg)]()
[![PowerShell Gallery](https://img.shields.io/badge/PowerShell%20Module-PS1%20NativeInterop-blue.svg)]()

> Managed COM / Win32 API bridge and utilities for PowerShell (PS1).  
> High-level wrapper for advanced native interoperability and system inspection — intended for legitimate administrative, defensive, and research use.

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
>
> # PowerShell Low-Level Utilities

A collection of PowerShell helper functions for working with Windows low-level APIs, pointers, memory, COM interfaces, tokens, syscalls and more.

## Big label -> Function List

````
<#
View Error Message Info, using system32 Build In DLL
For varios types Win32, Ntstatus, hResult, Activation, etc
#>
Function Parse-ErrorMessage

<#
Get Error facility info for specific error Message
#>
Function Parse-ErrorFacility

<#
Join multiple flags, with ease
#>
Function Bor

<#
Extract Flags From `Flag` value
#>
Function Get-EnumFlags

<#
Dump memory for view it later, with hex editior
#>
Function Dump-MemoryAddress

<#
Initilize & Free, new pointer,
Using Byte[] or Size, 
And Also Set First Value For Offset 0x0
Or make Ref Type [Allocate Handle Size, And Copy Value Of IntPtr]
Instead ByRef, who sometimes fail. !
#>
Function New-IntPtr

<#
Check if pointer is a valid pointer,
Return True Or False
#>
Function IsValid-IntPtr

<#
Free pointer from diffrent type's:
"HGlobal", "Handle", "NtHandle",
"ServiceHandle", "Heap", "STRING",
"UNICODE_STRING", "BSTR", "VARIANT",
"Local", "Auto", "Desktop", "WindowStation",
"License", "LSA"
#>
Function Free-IntPtr

<#
Register NTDLL! or Any Win32! Api,
with ease, Using Reflection
#>
Function Register-NativeMethods

<#
Convert Number to Word
will get diffrent result each time's
#>
Function Get-Base26Name

<#
Get Function ASM Byte's Code
using Local file [system32] folder
in case of security software hook
#>
Function Get-SysCallData

<#
Call a specifc Function of IID of CLSID
using delegete & Low level call's
Also, can do it manual, Using
Initialize-ComObject, Receive-ComObject, .. etc
#>
Function Use-ComInterface

<#
Helper to Create Full Interface, using Reflection
and call, the function later.
Alternative way for Use-ComInterface
#>
Function New-ComInterface
Function Invoke-ComInterface

<#
Helper to Create Full Structure, using Reflection
Also with GetSize And Casting Support
[Type]::GetSize(), [Type]$handle
#>
Function New-Struct

<#
Helper to Call any Win32 / Low level Api call,
Including syscall with ease
Minimal code,Parameters, Dll Name, Function Name
Using delegete, And ASM, And Protect/Allocate Both way's supported
#>
Function Invoke-UnmanagedMethod

<#
Helper to Create Managed UNICODE_STRING / STRING Struct
for x64 x68, and read the info, or free it later
or Create With specifc parameters [length, max, etc]
#>
Function Init-NativeString
Function Parse-NativeString
Function Free-NativeString

<#
Helper to Create Managed VARIANT Struct
for x64 x68, and read the info, or free it later
#>
Function New-Variant
Function Parse-Variant
Function Free-Variant

<#
Helper to Managed GUID, From / TO:
GUID <> IntPtr <> Byte's
#>
Function Guid-Handler

<#
Helper to Managed Privileges,
For specifc hToken, [Process]
Can Query Privileges, Adjust Privileges
Or Adjust All Privileges, and Also
give Account A specifc Privileges, like SeAssignPrimaryTokenPrivilege
which normal user / Admin doesn't have 
#>
Function Adjust-TokenPrivileges

<#
Check Current Process Token, 
if Belong to Administrators Group or is a System User
or is it an elavated Process, using Sid, who is Build manully
#>
Function Check-AccountType

<#
NtCurrentTeb implantation in PS'1 code
using 5 diffrent Way's, By ref, return, etc etc.
and using 3 Types  option low level API [Proctect, Allocate, AllocateEx]
And lot of ASM , And extra suppprt for specifc Special Location's:
NtCurrentTeb, NtCurrentTeb->ClientID,Peb,Ldr,ProcessHeap,Parameters
#>
Function NtCurrentTeb

<#
Low level Api, for loading DLL, 
As Data, Or `Not` Data,
Also, implantate Get-DllHandle Function
who is high level, and return pointer for loaded DLL only
#>
Function Ldr-LoadDll
Function Get-DllHandle

<#
Read Loader api module, using low level way
Read Process TEB-> LDR, And parse them in 3 way's
"Load", "Memory", "Init"
#>
Function Get-LoadedModules

<#
Query Process using Low level Api call's
And get basic info, PebBaseAddress, UniqueProcessId,
InheritedFromUniqueProcessId, ImageFileName
#>
Function Query-Process

<#
Get Token Of user, who is logged in OR Not
Using 2 Api, LogonUserExExW, LsaLogonUser
with option to modify parametes like LogonType, TokenType
LogonType->0x03, TokenType ->0x02 for example
#>
Function Obtain-UserToken

<#
Get Token From Service/Process/Process who is also Service
With -Impersonat option in case of Duplicate Token
#>
Function Get-ProcessHelper
Function Get-ProcessHandle

<#
Process a Token of diffrent user,
to allow it to Run interactive window on Current User Desktop
Using new Desktop\Vista or Current
#>
Function Process-UserToken

<#
Create Basic Process or advanced Process,
Using User Token, Process Token, using High level & low level APi
Also Support Duplicate / As Parent Method's [RunAsTi]
Send-CsrClientCall, is a low level helper, to allow run Notepad for example
on current desktop, otherwise, some Process will fail
#>
Function Invoke-Process
Function Invoke-ProcessAsUser
Function Invoke-NativeProcess
Function Send-CsrClientCall
````

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

