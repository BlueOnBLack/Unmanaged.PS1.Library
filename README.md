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
