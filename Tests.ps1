# Use-ComInterface Example

write-host "`n`nIEnumerator & Params\values [Test]`nDCB00C01-570F-4A9B-8D69-199FDBA5723B->DCB00000-570F-4A9B-8D69-199FDBA5723B->GetNetwork`n"
[intPtr]$ppEnumNetwork = [intPtr]::Zero
Use-ComInterface `
    -CLSID "DCB00C01-570F-4A9B-8D69-199FDBA5723B" `
    -IID "DCB00000-570F-4A9B-8D69-199FDBA5723B" `
    -Index 1 `
    -Name "GetNetwork" `
    -Return "uint" `
    -Params 'system.UINT32 Flags, out INTPTR ppEnumNetwork' `
    -Values @(1, [ref]$ppEnumNetwork)

if ($ppEnumNetwork -ne [IntPtr]::Zero) {
    $networkList = $ppEnumNetwork | Receive-ComObject
    foreach ($network in $networkList) {
        "Name: $($network.GetName()), IsConnected: $($network.IsConnected())"
    }
    $networkList | Release-ComObject
}

Write-Host
Write-Host

# Invoke-UnmanagedMethod Example

# ZwQuerySystemInformation
# https://www.geoffchappell.com/studies/windows/km/ntoskrnl/api/ex/sysinfo/query.htm?tx=61&ts=0,1677

# SYSTEM_PROCESS_INFORMATION structure
# https://www.geoffchappell.com/studies/windows/km/ntoskrnl/api/ex/sysinfo/process.htm

# ZwQuerySystemInformation
# https://www.geoffchappell.com/studies/windows/km/ntoskrnl/api/ex/sysinfo/query.htm?tx=61&ts=0,1677

# SYSTEM_BASIC_INFORMATION structure
# https://www.geoffchappell.com/studies/windows/km/ntoskrnl/inc/api/ntexapi/system_basic_information.htm

# Step 1: Get required buffer size
Write-Host
$ReturnLength = 0
$dllResult = Invoke-UnmanagedMethod `
  -Dll "ntdll.dll" `
  -Function "ZwQuerySystemInformation" `
  -Return "uint32" `
  -Params "int SystemInformationClass, IntPtr SystemInformation, uint SystemInformationLength, ref uint ReturnLength" `
  -Values @(0, [IntPtr]::Zero, 0, [ref]$ReturnLength)

# Allocate buffer (add some extra room just in case)
$infoBuffer = New-IntPtr -Size $ReturnLength

# Step 2: Actual call with allocated buffer
$result = Invoke-UnmanagedMethod `
  -Dll "ntdll.dll" `
  -Function "ZwQuerySystemInformation" `
  -Return "uint32" `
  -Params "int SystemInformationClass, IntPtr SystemInformation, uint SystemInformationLength, ref uint ReturnLength" `
  -Values @(0, $infoBuffer, $ReturnLength, [ref]$ReturnLength)

if ($result -ne 0) {
    Write-Host "NtQuerySystemInformation failed: 0x$("{0:X}" -f $result)"
    Parse-ErrorMessage -MessageId $result
    New-IntPtr -hHandle $infoBuffer -Release
    return
}

# Parse values from the structure
$sysBasicInfo = [PSCustomObject]@{
    PageSize                     = [Marshal]::ReadInt32($infoBuffer,  0x08)
    NumberOfPhysicalPages        = [Marshal]::ReadInt32($infoBuffer,  0x0C)
    LowestPhysicalPageNumber     = [Marshal]::ReadInt32($infoBuffer,  0x10)
    HighestPhysicalPageNumber    = [Marshal]::ReadInt32($infoBuffer,  0x14)
    AllocationGranularity        = [Marshal]::ReadInt32($infoBuffer,  0x18)
    MinimumUserModeAddress       = [Marshal]::ReadIntPtr($infoBuffer, 0x20)
    MaximumUserModeAddress       = [Marshal]::ReadIntPtr($infoBuffer, 0x28)
    ActiveProcessorsAffinityMask = [Marshal]::ReadIntPtr($infoBuffer, 0x30)
    NumberOfProcessors           = [Marshal]::ReadByte($infoBuffer,   0x38)
}

$sysBasicInfo | Format-List

[intPtr]$ppEnumNetwork = [intPtr]::Zero
Use-ComInterface `
    -CLSID "DCB00C01-570F-4A9B-8D69-199FDBA5723B" `
    -IID "DCB00000-570F-4A9B-8D69-199FDBA5723B" `
    -Index 1 `
    -Name "GetNetwork" `
    -Return "uint" `
    -Params 'system.UINT32 Flags, out INTPTR ppEnumNetwork' `
    -Values @(1, [ref]$ppEnumNetwork)

if ($ppEnumNetwork -ne [IntPtr]::Zero) {
    $networkList = $ppEnumNetwork | Receive-ComObject
    foreach ($network in $networkList) {
        "Name: $($network.GetName()), IsConnected: $($network.IsConnected())"
    }
    $networkList | Release-ComObject
}

Write-Host
<#
https://github.com/winsiderss/phnt/blob/master/ntexapi.h
https://doxygen.reactos.org/d6/d0e/ndk_2iotypes_8h_source.html
https://processhacker.sourceforge.io/doc/ntexapi_8h_source.html
https://github.com/mic101/windows/blob/master/WRK-v1.2/public/sdk/inc/ntexapi.h
https://www.geoffchappell.com/studies/windows/km/ntoskrnl/source/inc/ntexapi.htm
https://github.com/xmoezzz/SiglusExtract/blob/master/source/SiglusExtract/NativeLib/ntexapi.h

// private
typedef struct _BOOT_ENTRY_LIST
{
    ULONG NextEntryOffset;
    BOOT_ENTRY BootEntry;
} BOOT_ENTRY_LIST, *PBOOT_ENTRY_LIST;

typedef struct _BOOT_ENTRY
{
    ULONG Version;
    ULONG Length;
    ULONG Id;
    ULONG Attributes;
    ULONG FriendlyNameOffset;
    ULONG BootFilePathOffset;
    ULONG OsOptionsLength;
    _Field_size_bytes_(OsOptionsLength) UCHAR OsOptions[1];
} BOOT_ENTRY, *PBOOT_ENTRY;

typedef struct _BOOT_ENTRY {
    ULONG Version;
    ULONG Length;
    ULONG Id;
    ULONG Attributes;
    ULONG FriendlyNameOffset;
    ULONG BootFilePathOffset;
    ULONG OsOptionsLength;
    UCHAR OsOptions[ANYSIZE_ARRAY];
    //WCHAR FriendlyName[ANYSIZE_ARRAY];
    //FILE_PATH BootFilePath;
} BOOT_ENTRY, *PBOOT_ENTRY;

// private
typedef struct _EFI_DRIVER_ENTRY
{
    ULONG Version;
    ULONG Length;
    ULONG Id;
    ULONG FriendlyNameOffset;
    ULONG DriverFilePathOffset;
} EFI_DRIVER_ENTRY, *PEFI_DRIVER_ENTRY;

// private
typedef struct _EFI_DRIVER_ENTRY_LIST
{
    ULONG NextEntryOffset;
    EFI_DRIVER_ENTRY DriverEntry;
} EFI_DRIVER_ENTRY_LIST, *PEFI_DRIVER_ENTRY_LIST;

// private
typedef struct _FILE_PATH
{
    ULONG Version;
    ULONG Length;
    ULONG Type;
    _Field_size_bytes_(Length) UCHAR FilePath[1];
} FILE_PATH, *PFILE_PATH;


#define FILE_PATH_VERSION 1
#define FILE_PATH_TYPE_ARC           1
#define FILE_PATH_TYPE_ARC_SIGNATURE 2
#define FILE_PATH_TYPE_NT            3
#define FILE_PATH_TYPE_EFI           4
#define FILE_PATH_TYPE_MIN FILE_PATH_TYPE_ARC
#define FILE_PATH_TYPE_MAX FILE_PATH_TYPE_EFI

#define WINDOWS_OS_OPTIONS_SIGNATURE "WINDOWS"
#define WINDOWS_OS_OPTIONS_VERSION 1

#define BOOT_ENTRY_ATTRIBUTE_ACTIVE             0x00000001
#define BOOT_ENTRY_ATTRIBUTE_DEFAULT            0x00000002
#define BOOT_ENTRY_ATTRIBUTE_WINDOWS            0x00000004
#define BOOT_ENTRY_ATTRIBUTE_REMOVABLE_MEDIA    0x00000008

typedef struct _WINDOWS_OS_OPTIONS {
    UCHAR Signature[8];
    ULONG Version;
    ULONG Length;
    ULONG OsLoadPathOffset;
    WCHAR OsLoadOptions[ANYSIZE_ARRAY];
    //FILE_PATH OsLoadPath;
} WINDOWS_OS_OPTIONS, *PWINDOWS_OS_OPTIONS;
#>

Write-Host
$length = [Uint32]1
$Ptr    = [IntPtr]::Zero
$lastErr = Invoke-UnmanagedMethod `
    -Dll "ntdll.dll" `
    -Function "NtEnumerateBootEntries" `
    -Return "Int64" `
    -Params "IntPtr Ptr, ref uint length" `
    -Values @($Ptr, [ref]$length)
$ntStatus = Parse-ErrorMessage `
    -MessageId $lastErr

if ([Int64]$lastErr -eq 3221225507) {
    $Ptr = New-IntPtr -Size $length
    $lastErr = Invoke-UnmanagedMethod `
        -Dll "ntdll.dll" `
        -Function "NtEnumerateBootEntries" `
        -Return "Int64" `
        -Params "IntPtr Ptr, ref uint length" `
        -Values @($Ptr, [ref]$length)
}

function Show-BootEntryDebugLog {
    param (
        [IntPtr]$ptr
    )

    $offsetNames = @{
        0x00 = "Offset Of NextEntryOffset"
        0x04 = "Version"
        0x08 = "Length"
        0x0C = "Id"
        0x10 = "Attributes"
        0x14 = "Offset Of FriendlyName"
        0x18 = "Offset Of BootFilePath"
        0x1C = "OsOptionsLength"
    }

    Write-Host
    Write-Host "Boot Entry Fields:"
    for ($offset = 0; $offset -le 0x1C; $offset += 4) {
        $value = [Marshal]::ReadInt32($ptr, $offset)
        $name = if ($offsetNames.ContainsKey($offset)) { $offsetNames[$offset] } else { "Unknown" }
        Write-Host ("Offset 0x{0:X2}  = {1}  > {2}" -f $offset, $value, $name)
    }

    $bootFilePathOffset = [Marshal]::ReadInt32($ptr, 0x18)
    $additionalOffset = 62
    $finalOffset = $bootFilePathOffset + $additionalOffset
    $bootFilePathPtr = [IntPtr]::Add($ptr, $finalOffset)

    # Read the string
    $bootFilePathStr = [Marshal]::PtrToStringUni($bootFilePathPtr)
    Write-Host ("Offset 0x{0:X2} = {1}  > EFI unicode string:" -f $finalOffset, $finalOffset)
}
function Dump-MemoryBlock {
    param (
        [IntPtr]$BaseAddress,
        [int]$Offset,
        [int]$Length,
        [string]$Label,
        [string]$Id
    )

    $ptr = [IntPtr]::Add($BaseAddress, $Offset)
    $buffer = New-Object byte[] $Length
    [Marshal]::Copy($ptr, $buffer, 0, $Length)

    # Optional hex dump
    Write-Host
    Write-Host "$Label Hex Dump" -ForegroundColor Red
    ($buffer | ForEach-Object { "{0:X2}" -f $_ }) -join " "

    # Save to Desktop
    $desktop = [Environment]::GetFolderPath("Desktop")
    $filePath = Join-Path $desktop "${Id}_${Label}.bin"
    [System.IO.File]::WriteAllBytes($filePath, $buffer)

    Write-Host "Saved to: $filePath" -ForegroundColor Green
}
function Parse-MixedAnsiUnicode {
    param (
        [Parameter(Mandatory=$true)][IntPtr] $BasePtr,
        [Parameter(Mandatory=$true)][int] $Offset,
        [Parameter(Mandatory=$true)][int] $Length
    )

    function IsAnsiPrintable([byte]$b) {
        return ($b -ge 32 -and $b -le 126)
    }

    function IsUnicodePrintable([char]$c) {
        return ($c -ge [char]32 -and $c -le [char]126)
    }

    # Copy bytes from memory to managed array
    $rawBytes = New-Object byte[] $Length
    [Marshal]::Copy([IntPtr]::Add($BasePtr, $Offset), $rawBytes, 0, $Length)

    $pos = 0
    $results = @()

    while ($pos -lt $rawBytes.Length) {
        # Skip junk bytes (non-printable and not zero)
        if (-not (IsAnsiPrintable $rawBytes[$pos]) -and $rawBytes[$pos] -ne 0) {
            $pos++
            continue
        }

        # Detect UTF-16LE start: printable byte followed by 0x00
        if (($pos + 1) -lt $rawBytes.Length -and $rawBytes[$pos + 1] -eq 0 -and (IsAnsiPrintable $rawBytes[$pos])) {
            $utf16Chars = New-Object System.Collections.Generic.List[char]

            while (($pos + 1) -lt $rawBytes.Length) {
                $pair = [byte[]]@($rawBytes[$pos], $rawBytes[$pos + 1])
                $char = [System.Text.Encoding]::Unicode.GetString($pair)[0]

                # Check for the specific byte pair 20,00 (little-endian representation of ' ')
                if ($pair[0] -eq 0x20 -and $pair[1] -eq 0x00) {
                    $pos += 2
                    break
                }

                if ($char -eq [char]0) { 
                    $pos += 2
                    break
                }
                if (-not (IsUnicodePrintable $char)) {
                    $pos += 2
                    break
                }
                

                $utf16Chars.Add($char)
                $pos += 2
            }

            if ($utf16Chars.Count -gt 2) {
                $results += (-join $utf16Chars).Trim()
            }
        }
        else {
            # Read ANSI bytes until non-printable or zero
            $ansiChars = New-Object System.Collections.Generic.List[char]

            while ($pos -lt $rawBytes.Length) {
                $b = $rawBytes[$pos]

                if (-not (IsAnsiPrintable $b) -or $b -eq 0) {
                    break
                }

                $ansiChars.Add([char]$b)
                $pos++
            }

            if ($ansiChars.Count -gt 2) {
                $results += (-join $ansiChars).Trim()
            }
            else {
                $pos++
            }
        }
    }

    return $results
}

if ([Int64]$lastErr -eq 0) {
    $BassAddress = $Ptr
    $bootEntries = @()
    do {
        # Initialize base address and index
        $index = 1

        # Read values from the BootEntry structure
        $Version = [Uint32]([Marshal]::ReadInt32($BassAddress, 4 * $index++))
        $Length = [Uint32]([Marshal]::ReadInt32($BassAddress, 4 * $index++))
        $Id = [Uint32]([Marshal]::ReadInt32($BassAddress, 4 * $index++))
        $Attributes = [Uint32]([Marshal]::ReadInt32($BassAddress, 4 * $index++))

        # Read FriendlyName
        $FriendlyNameOffset = [Uint32]([Marshal]::ReadInt32($BassAddress, 4 * $index++))
        $FriendlyName = [Marshal]::PtrToStringUni([IntPtr]::Add($BassAddress, [IntPtr]::Add($FriendlyNameOffset,4)))

        # Read BootFilePathOffset from structure
        $BootFilePathOffset = [UInt32]([Marshal]::ReadInt32($BassAddress, 4 * $index++))
        $BootFilePtr = [IntPtr]::Add($BassAddress, $BootFilePathOffset)

        # Read FilePath.Type (ULONG at offset 8)
        $EfiPtr = [IntPtr]::Add($BootFilePtr, $FilePathOffset)
        $FilePath = $null
        $FilePathOffset = 0x3E
        $ToRead = ($Length - ($BootFilePathOffset + $FilePathOffset)) / 2
        if ($ToRead -gt 0) {
            $FilePath = [Marshal]::PtrToStringUni($EfiPtr, $ToRead)
        }

        # Read OsOptions
        $OsOptions = $null
        $OsOptionsLength = [Uint32]([Marshal]::ReadInt32($BassAddress, 4 * $index++))

        if ($OsOptionsLength -gt 0) {
            $parsedStrings = Parse-MixedAnsiUnicode -BasePtr $BassAddress -Offset 32 -Length ($OsOptionsLength)
            $OsOptions = $parsedStrings | ForEach-Object { Write-Output $_ }
        }

        # Define the attribute flags
        $ATTRIBUTE_ACTIVE  = 0x00000001
        $ATTRIBUTE_DEFAULT = 0x00000002
        $ATTRIBUTE_WINDOWS = 0x00000004
        $ATTRIBUTE_REMOVABLE_MEDIA = 0x00000008

        # Check each attribute flag using bitwise AND (-band)
        $Flags = @()
        if ($Attributes -band $ATTRIBUTE_ACTIVE) {
            $Flags += "ACTIVE"
        }
        if ($Attributes -band $ATTRIBUTE_DEFAULT) {
            $Flags += "DEFAULT"
        }
        if ($Attributes -band $ATTRIBUTE_WINDOWS) {
            $Flags += "WINDOWS"
        }
        if ($Attributes -band $ATTRIBUTE_REMOVABLE_MEDIA) {
            $Flags += "REMOVABLE_MEDIA"
        }
        Write-Host
        write-host "ID              = $Id"
        write-host "Version         = $Version"
        write-host "Attributes      = $($Flags -join ', ')"
        write-host "OsOptions       = $OsOptions"
        write-host "FriendlyName    = $FriendlyName"
        write-host "FilePath        = $FilePath"
        Show-BootEntryDebugLog -ptr $BassAddress
        Write-Host
        Write-Host

        #Dump-MemoryBlock -BaseAddress $BassAddress -Offset 28 -Length ($FriendlyNameOffset - 28) -Label "OsOptions" -Id $Id
        #Dump-MemoryBlock -BaseAddress $BassAddress -Offset $BootFilePathOffset -Length ($Length - $BootFilePathOffset) -Label "FilePath" -Id $Id
        #Dump-MemoryBlock -BaseAddress $BassAddress -Offset 0 -Length $Length -Label "BassAddress" -Id $Id

        $index = $Version = $Length = $null
        $OsOptionsLength = $OsOptions = $null
        $Id = $Attributes = $FriendlyNameOffset = $null
        $FriendlyName = $BootFilePathOffset = $BootFilePtr = $null

        # Next Item
        $nextOffset = [Uint32]([Marshal]::ReadInt32($BassAddress))
        $BassAddress = [IntPtr]::Add($BassAddress, $nextOffset)
    } while ($nextOffset -ne 0)
    [Marshal]::FreeHGlobal($Ptr)
}

# Create Process AS System, using TrustedInstaller service
# sometimes it fail on first time, since, service not loaded well.
try {
    Invoke-Process `
        -CommandLine "cmd /k echo Hello From TrustedInstaller && whoami" `
        -ServiceName TrustedInstaller `
        -RunAsConsole `
        -RunAsTI
}
# Create Process AS System, using winlogon Process
# its still better to use TrustedInstaller service.
catch {
    Invoke-Process `
        -CommandLine "cmd /k echo Hello From winlogon && whoami" `
        -ProcessName winlogon `
        -RunAsConsole `
        -RunAsTI
}

try {
    $hHandle = Get-ProcessHandle `
        -ProcessName 'TrustedInstaller.exe'
}
catch {
    $hHandle = Get-ProcessHandle `
        -ServiceName 'TrustedInstaller'
}

Invoke-NativeProcess `
    -ImageFile "cmd.exe" `
    -commandLine "/k whoami" `
    -hProc $hHandle

# Start notepad & kill it later

Invoke-NativeProcess 'notepad' -Register
$procID = Query-Process -ProcessName 'notepad' | select -Last 1 -ExpandProperty UniqueProcessId
if (-not $procID) {
    Write-Error "Notepad process not found.!"
    return
}

$hProc = Get-ProcessHandle -ProcessId $procID
$ret = Invoke-UnmanagedMethod `
    -Dll "ntdll.dll" `
    -Function "ZwTerminateProcess" `
    -Return "NTSTATUS" `
    -Params "HANDLE ProcessHandle, NTSTATUS ExitStatus" `
    -Values @($hProc, 1)
if ($ret -ne 0) {
    Write-Error "ZwTerminateProcess fail with $lastError,`n$(Parse-ErrorMessage -MessageId $ret -Flags NTSTATUS)"
    return
}
Free-IntPtr $hProc -Method NtHandle
Query-Process -ProcessName 'notepad'

# Another stupid demo, with some crap properties like Size_T.
$ptr = [Marshal]::AllocHGlobal(200)
0..20| %{[Marshal]::WriteInt64($ptr, $_*8, 10)}
Dump-MemoryAddress -Pointer $ptr -Length 200 -FileName "BEfore"
$lastErr = Invoke-UnmanagedMethod `
    -Dll "NtosKrnl.exe" `
    -Function "RtlZeroMemory" `
    -Return "Void" `
    -Params "void* Destination, size_t Length" `
    -Values @($Ptr, [UintPtr]::new(200))
Dump-MemoryAddress -Pointer $ptr -Length 200 -FileName "After"
[Marshal]::FreeHGlobal($ptr)

# Test Charset <> Ansi
$Func = Register-NativeMethods @(
    @{ 
        Name       = "MessageBoxA"
        Dll        = "user32.dll"
        ReturnType = [int]
        CharSet    = 'Ansi'
        Parameters = [Type[]]@(
            [IntPtr],    # hWnd
            [string],    # lpText
            [string],    # lpCaption
            [uint32]     # uType
        )
    })
$Func::MessageBoxA(
    [IntPtr]::Zero, "Hello from ANSI!", "MessageBoxA", 0)

# Test Charset <> Ansi
Invoke-UnmanagedMethod `
    -Dll "user32.dll" `
    -Function "MessageBoxA" `
    -Return "int32" `
    -Params "HWND hWnd, LPCSTR lpText, LPCSTR lpCaption, UINT uType" `
    -Values @(0, "Hello from ANSI!", "MessageBoxA", 0) `
    -CharSet Ansi

# Test Charset <> Ansi
Invoke-UnmanagedMethod `
    -Dll "User32.dll" `
    -Function "MessageBoxA" `
    -Values @(
        [IntPtr]::Zero,
        "Some Text",
        "Some title",
        0)

<#
MIDL_INTERFACE("85713fa1-7796-4fa2-be3b-e2d6124dd373")
    IWindowsUpdateAgentInfo : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetInfo( 
            /* [in] */ VARIANT varInfoIdentifier,
            /* [retval][out] */ __RPC__out VARIANT *retval) = 0;
        
    };

// {C2E88C2F-6F5B-4AAA-894B-55C847AD3A2D}
DEFINE_GUID(CLSID_WindowsUpdateAgentInfo,0xC2E88C2F,0x6F5B,0x4AAA,0x89,0x4B,0x55,0xC8,0x47,0xAD,0x3A,0x2D);

* IWindowsUpdateAgentInfo interface (wuapi.h)
* https://learn.microsoft.com/en-us/windows/win32/api/wuapi/nf-wuapi-iwindowsupdateagentinfo-getinfo

HRESULT GetInfo(
  [in]  VARIANT varInfoIdentifier,
  [out] VARIANT *retval
);
#>

# Read Info from interface using `Variant`struct
"ApiMajorVersion", "ApiMinorVersion", "ProductVersionString" | ForEach-Object {
    $name = $_
    $outVarPtr = New-Variant -Type VT_EMPTY
    $inVarPtr  = New-Variant -Type VT_BSTR -Value $name
    try {
        $ret = Use-ComInterface `
            -CLSID "C2E88C2F-6F5B-4AAA-894B-55C847AD3A2D" `
            -IID "85713fa1-7796-4fa2-be3b-e2d6124dd373" `
            -Index 1 -Name "GetInfo" `
            -Values @($inVarPtr, $outVarPtr)

        if ($ret -eq 0) {
            $value = Parse-Variant -variantPtr $outVarPtr
            Write-Host "$name -> $value"
        }

    } finally {
        Free-IntPtr -handle $inVarPtr  -Method VARIANT
        Free-IntPtr -handle $outVarPtr -Method VARIANT
    }
}

Write-Host
write-host 'NtGetNextProcess Test'
$hProc = [IntPtr]::Zero
$hProcNext = [IntPtr]::Zero
Invoke-UnmanagedMethod `
    -Dll NTDLL `
    -Function ZwGetNextProcess `
    -Values @($hProc, [UInt32]0x02000000, [UInt32]0x00, [UInt32]0x00, ([ref]$hProcNext)) `
    -SysCall
$hProcNext

Write-Host
write-host 'NtGetNextThread Test'
$hThread = [IntPtr]::Zero
$hThreadNext = [IntPtr]::Zero
Invoke-UnmanagedMethod `
    -Dll NTDLL `
    -Function ZwGetNextThread `
    -Values @([IntPtr]::new(-1), $hThread, 0x0040, 0x00, 0x00, ([ref]$hThreadNext)) `
    -SysCall
$hThreadNext

Write-Host
write-host 'NtClose Test'
Invoke-UnmanagedMethod `
    -Dll NTDLL `
    -Function NtClose `
    -Values @([IntPtr]$hProcNext) `
    -SysCall

Write-Host
write-host 'NtCreateFile Test'

$FileHandle = [IntPtr]::Zero
$IoStatusBlock    = New-IntPtr -Size 16
$ObjectAttributes = New-IntPtr -Size 48 -WriteSizeAtZero

$filePath = ("\??\{0}\test.txt" -f [Environment]::GetFolderPath('Desktop'))
$ObjectName = Init-NativeString -Encoding Unicode -Value $filePath
[Marshal]::WriteIntPtr($ObjectAttributes, 0x10, $ObjectName)
[Marshal]::WriteInt32($ObjectAttributes,  0x18, 0x40)

# Call syscall
Invoke-UnmanagedMethod `
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
    -SysCall

Free-NativeString -StringPtr $ObjectName


# Fail from system/TI
Write-Host 'Invoke-ProcessAsUser, As Logon' -ForegroundColor Green
Invoke-ProcessAsUser `
    -Application cmd `
    -CommandLine "/k whoami" `
    -UserName user `
    -Password 0444 `
    -Mode Logon `
    -RunAsConsole

# Work From both Normal/Admin/System/TI Account
Write-Host 'Invoke-ProcessAsUser, As Token' -ForegroundColor Green
Invoke-ProcessAsUser `
    -Application cmd `
    -CommandLine "/k whoami" `
    -UserName user `
    -Password 0444 `
    -Mode Token `
    -RunAsConsole

# only work from system/TI
Write-Host 'Invoke-ProcessAsUser, As User' -ForegroundColor Green
Invoke-ProcessAsUser `
    -Application cmd `
    -CommandLine "/k whoami" `
    -UserName user `
    -Password 0444 `
    -Mode User `
    -RunAsConsole

# only work from system/TI
Write-Host 'Invoke-NativeProcess, with hToken' -ForegroundColor Green
$hToken = Obtain-UserToken `
    -UserName user `
    -Password 0444 `
    -loadProfile
Invoke-NativeProcess `
    -ImageFile cmd `
    -commandLine "/k whoami" `
    -hToken $hToken

Free-IntPtr $hToken -Method NtHandle
Free-IntPtr $hProc  -Method NtHandle