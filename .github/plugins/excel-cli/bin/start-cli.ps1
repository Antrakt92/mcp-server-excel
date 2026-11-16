#Requires -Version 7.4
[CmdletBinding()]
param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$PassthroughArgs
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

if ($null -eq $PassthroughArgs) {
    $PassthroughArgs = @()
}

$npxCommand = Get-Command "npx" -CommandType Application -All -ErrorAction SilentlyContinue |
    Where-Object Name -in "npx.exe", "npx.cmd" |
    Select-Object -First 1
$nativeArguments = @("-y", "@sbroenne/excelcli@latest") + $PassthroughArgs
if ($null -eq $npxCommand) {
    throw "excel-cli requires Node.js 18 or later with npm/npx available on PATH."
}

$binaryPath = $npxCommand.Source
if ($npxCommand.Name -eq "npx.cmd") {
    # Run npm's JS entry point directly to avoid corrupting arguments in cmd.exe.
    $nodeCommand = Get-Command "node.exe" -CommandType Application -ErrorAction SilentlyContinue |
        Select-Object -First 1
    $npxDirectory = Split-Path -Parent $npxCommand.Source
    $npxCliPath = Join-Path $npxDirectory "node_modules\npm\bin\npx-cli.js"
    if ($null -ne $nodeCommand -and (Test-Path -LiteralPath $npxCliPath -PathType Leaf)) {
        $binaryPath = $nodeCommand.Source
        $nativeArguments = @($npxCliPath) + $nativeArguments
    } else {
        # PATHEXT can rank a forwarding CMD shim ahead of its own executable.
        $binaryPath = Join-Path $npxDirectory "npx.exe"
        if (-not (Test-Path -LiteralPath $binaryPath -PathType Leaf)) {
            throw "excel-cli requires Node.js 18 or later with npm/npx available on PATH."
        }
    }
}

$startInfo = [Diagnostics.ProcessStartInfo]::new()
$startInfo.FileName = $binaryPath
foreach ($argument in $nativeArguments) { $startInfo.ArgumentList.Add($argument) }
$startInfo.UseShellExecute = $false
$startInfo.CreateNoWindow = $true
$startInfo.RedirectStandardOutput = $true
$startInfo.RedirectStandardError = $true

$process = [System.Diagnostics.Process]::Start($startInfo)
try {
    $stdoutTask = $process.StandardOutput.ReadToEndAsync()
    $stderrTask = $process.StandardError.ReadToEndAsync()
    $process.WaitForExit()
    $stdout = $stdoutTask.GetAwaiter().GetResult()
    $stderr = $stderrTask.GetAwaiter().GetResult()
    if (-not [string]::IsNullOrEmpty($stderr)) { [Console]::Error.Write($stderr) }
    if (-not [string]::IsNullOrEmpty($stdout)) { Write-Output -NoEnumerate $stdout }
    $exitCode = $process.ExitCode
}
finally { $process.Dispose() }
exit $exitCode
