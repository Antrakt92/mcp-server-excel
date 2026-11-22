#Requires -Version 7.4
<#
.SYNOPSIS
    Forcibly stops CLI background services owned by the current account in this worktree and Windows session.
.DESCRIPTION
    Development builds do not save workbooks or terminate Excel. Unsaved work
    in stopped services may be lost. An optional pipe limits test-run cleanup.
#>
[CmdletBinding()]
param([string]$PipeName)

$ErrorActionPreference = 'Stop'
$buildProcess = [Diagnostics.Process]::GetCurrentProcess()
try { $buildSession = $buildProcess.SessionId }
finally { $buildProcess.Dispose() }
$buildIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
try { $buildOwnerSid = $buildIdentity.User.Value }
finally { $buildIdentity.Dispose() }
if ([string]::IsNullOrWhiteSpace($buildOwnerSid)) { throw 'Cannot verify the current Windows account.' }
$root = Split-Path -Parent $PSScriptRoot
$cliPaths = @(
    (Join-Path $root 'src\ExcelMcp.CLI\bin\Release\net10.0-windows\excelcli.exe'),
    (Join-Path $root 'src\ExcelMcp.CLI\bin\Debug\net10.0-windows\excelcli.exe'),
    (Join-Path $root 'tests\ExcelMcp.CLI.Tests\bin\Release\net10.0-windows\excelcli.exe'),
    (Join-Path $root 'tests\ExcelMcp.CLI.Tests\bin\Debug\net10.0-windows\excelcli.exe')
)

foreach ($candidate in Get-CimInstance Win32_Process -Filter "Name = 'excelcli.exe' AND SessionId = $buildSession") {
    if ($null -eq $candidate.SessionId -or $candidate.SessionId -ne $buildSession -or
        $cliPaths -notcontains $candidate.ExecutablePath -or
        $candidate.CommandLine -notmatch '^(?:"[^"]+"|\S+)\s+service\s+run(?:\s|$)' -or
        $null -eq $candidate.CreationDate) {
        continue
    }
    if (-not [string]::IsNullOrWhiteSpace($PipeName)) {
        $pipe = [regex]::Match($candidate.CommandLine,
            '(?:^|\s)--pipe-name\s+(?:"(?<pipe>[^"]+)"|(?<pipe>\S+))(?=\s|$)')
        if (-not $pipe.Success -or $pipe.Groups['pipe'].Value -ine $PipeName) {
            continue
        }
    }

    $process = $null
    $processFilter = "ProcessId = $($candidate.ProcessId) AND SessionId = $buildSession"
    try {
        $process = Get-Process -Id $candidate.ProcessId -ErrorAction Stop
        # Pin the handle; call its getter directly so PowerShell propagates errors.
        $null = $process.get_Handle()
        if ($process.HasExited -or $process.SessionId -ne $buildSession) {
            continue
        }
        $started = $process.StartTime.ToUniversalTime().Ticks
        if (($started - ($started % 10)) -ne $candidate.CreationDate.ToUniversalTime().Ticks) {
            continue
        }
        $current = Get-CimInstance Win32_Process -Filter $processFilter
        if ($process.HasExited -or $null -eq $current -or $current.ProcessId -ne $candidate.ProcessId -or
            $current.SessionId -ne $buildSession -or $current.CreationDate -ne $candidate.CreationDate -or
            $current.ExecutablePath -ne $candidate.ExecutablePath -or
            $current.CommandLine -ne $candidate.CommandLine) {
            continue
        }
        # An alternate Windows account can share the desktop session and worktree.
        $owner = Invoke-CimMethod -InputObject $current -MethodName GetOwnerSid
        if ($null -eq $owner -or $owner.ReturnValue -ne 0 -or $owner.Sid -cne $buildOwnerSid -or $process.HasExited) {
            continue
        }
        $process.Kill()
        Write-Host "Stopped development CLI service PID $($candidate.ProcessId)."
    }
    catch {
        # A service can exit between identity validation and termination.
        if (Get-CimInstance Win32_Process -Filter $processFilter) {
            throw
        }
    }
    finally {
        if ($null -ne $process) { $process.Dispose() }
    }
}
$global:LASTEXITCODE = 0
