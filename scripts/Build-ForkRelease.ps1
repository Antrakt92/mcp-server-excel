#Requires -Version 7.4
<#
.SYNOPSIS
Builds local portable Antrakt fork MCP and CLI release packages; never publishes.
.DESCRIPTION
Requires Windows, the pinned SDK and a clean Git checkout by default. Produces
ZIPs, SHA256SUMS and RELEASE-MANIFEST.json in a new output directory. This command
builds packages only; it does not run Excel or claim runtime/test acceptance.
.PARAMETER AllowDirtySource
For development packages only. Records dirty source explicitly in the manifest.
.PARAMETER NoRestore
Passes --no-restore to dotnet publish when the win-x64 assets are already restored.
.EXAMPLE
pwsh -NoProfile -File scripts/Build-ForkRelease.ps1 -OutputDirectory artifacts/fork/3.0.0
#>
[CmdletBinding()]
param(
    [string]$Version = '3.0.0',
    [string]$OutputDirectory,
    [string]$DotNetExecutable = 'dotnet',
    [switch]$AllowDirtySource,
    [switch]$NoRestore
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
. (Join-Path $PSScriptRoot 'PackageHelpers.ps1')
$root = [IO.Path]::GetFullPath((Split-Path $PSScriptRoot -Parent))
if (-not $IsWindows) { throw 'Portable fork builds require Windows.' }
$Version = Get-PackageVersion $Version
if ($Version -cnotmatch '^\d+\.\d+\.\d+$') { throw 'Fork release version must be a stable three-part numeric version.' }
$assemblyVersion = "$Version.0"
$null = [Version]::Parse($assemblyVersion)
if (-not $OutputDirectory) { $OutputDirectory = "artifacts/fork/$Version" }
$output = [IO.Path]::GetFullPath($OutputDirectory, $root)
Assert-PackageOutputPath $output $root
if ([IO.Path]::Exists($output)) { throw "Use a new release output directory: $output" }

function Read-ForkGit {
    param([string[]]$GitArguments)
    (Invoke-PackageCommand $root 'git' $GitArguments -TimeoutSeconds 30).Trim()
}
$commit = Read-ForkGit -GitArguments @('rev-parse', 'HEAD')
$tree = Read-ForkGit -GitArguments @('rev-parse', 'HEAD^{tree}')
$sourceStatus = Read-ForkGit -GitArguments @('status', '--porcelain=v1', '--untracked-files=normal')
$dirty = -not [string]::IsNullOrWhiteSpace($sourceStatus)
if ($dirty -and -not $AllowDirtySource) { throw 'Release source is dirty. Commit intended changes first, or use -AllowDirtySource for a development package.' }
$epoch = [long](Read-ForkGit -GitArguments @('show', '-s', '--format=%ct', 'HEAD'))
$stamp = [DateTimeOffset]::FromUnixTimeSeconds($epoch)
if ($stamp.Year -lt 1980 -or $stamp.Year -gt 2107) { throw 'Source commit timestamp is outside the ZIP timestamp range.' }
$sdk = (Invoke-PackageCommand $root $DotNetExecutable @('--version') -TimeoutSeconds 30).Trim()
$expectedSdk = (Read-PackageJson (Join-Path $root 'global.json')).sdk.version
if ($sdk -cne $expectedSdk) { throw "Use the pinned SDK $expectedSdk; selected SDK is $sdk." }
if ($root.IndexOfAny([char[]]@(',', ';')) -ge 0) { throw 'The source path contains characters that cannot be mapped safely by this release builder.' }
$repository = 'https://github.com/Antrakt92/mcp-server-excel'

function Assert-ForkSourceUnchanged {
    $currentCommit = Read-ForkGit -GitArguments @('rev-parse', 'HEAD')
    $currentTree = Read-ForkGit -GitArguments @('rev-parse', 'HEAD^{tree}')
    $currentStatus = Read-ForkGit -GitArguments @('status', '--porcelain=v1', '--untracked-files=normal')
    if ($currentCommit -cne $commit -or $currentTree -cne $tree -or $currentStatus -cne $sourceStatus) {
        throw 'Source changed during packaging; release output was not installed.'
    }
}

function New-ForkZip {
    param([string]$Source, [string]$Destination, [DateTimeOffset]$Timestamp)
    $file = [IO.File]::Open($Destination, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write)
    try {
        $zip = [IO.Compression.ZipArchive]::new($file, [IO.Compression.ZipArchiveMode]::Create, $true)
        try {
            foreach ($item in Get-ChildItem -LiteralPath $Source -File | Sort-Object Name) {
                $entry = $zip.CreateEntry($item.Name, [IO.Compression.CompressionLevel]::Optimal)
                $entry.LastWriteTime = $Timestamp
                $target = $entry.Open()
                try {
                    $inputFile = [IO.File]::OpenRead($item.FullName)
                    try { $inputFile.CopyTo($target) } finally { $inputFile.Dispose() }
                } finally { $target.Dispose() }
            }
        } finally { $zip.Dispose() }
    } finally { $file.Dispose() }
}

Invoke-PackageStage 'ForkRelease' {
    param($stage)
    $release = Join-Path $stage 'release'
    [IO.Directory]::CreateDirectory($release) | Out-Null
    $packages = @()
    foreach ($component in @(
        @{ Name = 'MCP'; Project = 'McpServer'; Executable = 'Sbroenne.ExcelMcp.McpServer.exe' },
        @{ Name = 'CLI'; Project = 'CLI'; Executable = 'excelcli.exe' }
    )) {
        $published = Join-Path $stage ("publish-" + $component.Name)
        $project = Join-Path $root "src/ExcelMcp.$($component.Project)/ExcelMcp.$($component.Project).csproj"
        $arguments = @(
            'publish', $project, '-c', 'Release', '-r', 'win-x64', '--self-contained', 'true',
            '-p:PublishSingleFile=true', '-p:IncludeNativeLibrariesForSelfExtract=true',
            '-p:PublishTrimmed=false', '-p:PublishReadyToRun=false', '-p:PackAsTool=false',
            "-p:Version=$Version", "-p:AssemblyVersion=$assemblyVersion", "-p:FileVersion=$assemblyVersion",
            "-p:InformationalVersion=$Version+$commit", '-p:IncludeSourceRevisionInInformationalVersion=false',
            "-p:RepositoryUrl=$repository", "-p:PackageProjectUrl=$repository",
            "-p:RepositoryCommit=$commit", '-p:AppInsightsConnectionString=', '-p:APPINSIGHTS_CONNECTION_STRING=',
            "-p:PathMap=$root=/_/src", '-p:DebugType=none', '-p:DebugSymbols=false',
            '-p:NuGetAudit=true', '-p:NuGetAuditMode=all', '-p:Deterministic=true',
            '-o', $published, '--verbosity', 'minimal'
        )
        if ($NoRestore) { $arguments += '--no-restore' }
        Invoke-PackageCommand $root $DotNetExecutable $arguments | Out-Host
        $executable = Join-Path $published $component.Executable
        Assert-PackageRuntimeArchitecture $executable 'x64'
        $metadata = [Diagnostics.FileVersionInfo]::GetVersionInfo($executable)
        if ($metadata.FileVersion -cne $assemblyVersion -or $metadata.ProductVersion.Split('+')[0] -cne $Version) {
            throw "The $($component.Name) executable has incorrect release version metadata."
        }
        $package = Join-Path $stage ("package-" + $component.Name)
        [IO.Directory]::CreateDirectory($package) | Out-Null
        Copy-PackageItem $executable (Join-Path $package $component.Executable)
        foreach ($notice in @('LICENSE', 'FORK-NOTICE.md')) {
            Copy-PackageItem (Join-Path $root $notice) (Join-Path $package $notice)
        }
        $usage = if ($component.Name -eq 'MCP') {
            'Configure your MCP client to launch Sbroenne.ExcelMcp.McpServer.exe as a local stdio server. Reconnect after switching versions.'
        } else { 'Run excelcli.exe --help for current CLI commands and service controls.' }
        $readme = @"
ExcelMCP - Antrakt Fork $Version ($($component.Name), Windows x64)
Based on upstream ExcelMcp 2.6.0 by Stefan Broenner and contributors.
Independent fork: $repository
MIT license; see LICENSE and FORK-NOTICE.md. No vendor endorsement is claimed.

Requires installed Microsoft Excel 2016 or later and an interactive Windows desktop.
The .NET runtime is included. Extract the complete ZIP into its own directory.
$usage
Set EXCELMCP_TELEMETRY_MODE=off to explicitly disable telemetry.

Verify this ZIP against SHA256SUMS from the same release. See RELEASE-MANIFEST.json
for source provenance and verification status. Build completion is not runtime acceptance.
Normal shutdown preserves unsaved and attached workbook state under upstream 2.6.0
rules. Cancellation does not authorize saving, closing, discarding or rollback.
Recovery: $repository/blob/v$Version/docs/fork/RECOVERY.md
"@
        [IO.File]::WriteAllText((Join-Path $package 'README.txt'), $readme.Replace("`r`n", "`n"))
        $filename = "ExcelMcp-Antrakt-$($component.Name)-$Version-win-x64.zip"
        $archive = Join-Path $release $filename
        New-ForkZip $package $archive $stamp
        $packages += [ordered]@{
            file = $filename
            bytes = ([IO.FileInfo]$archive).Length
            sha256 = (Get-FileHash -LiteralPath $archive -Algorithm SHA256).Hash.ToLowerInvariant()
            executable = $component.Executable
            executableSHA256 = (Get-FileHash -LiteralPath $executable -Algorithm SHA256).Hash.ToLowerInvariant()
        }
    }
    Assert-ForkSourceUnchanged
    $manifest = [ordered]@{
        schemaVersion = 1
        name = 'ExcelMCP - Antrakt Fork'
        version = $Version
        upstreamVersion = '2.6.0'
        upstreamRepository = 'https://github.com/sbroenne/mcp-server-excel'
        repository = $repository
        sourceCommit = $commit
        sourceTree = $tree
        sourceDirty = $dirty
        sdkVersion = $sdk
        runtimeIdentifier = 'win-x64'
        selfContained = $true
        singleFile = $true
        telemetryConnectionStringEmbedded = $false
        verification = [ordered]@{
            nativePEArchitecture = 'passed'
            executableVersionMetadata = 'passed'
            build = 'passed'
            tests = 'not run by this packaging command'
            protocolAndLifecycle = 'not run by this packaging command'
            desktopExcel = 'not run by this packaging command'
        }
        packages = $packages
    }
    $manifestPath = Join-Path $release 'RELEASE-MANIFEST.json'
    Write-PackageJson $manifestPath $manifest
    $checksums = @($packages | ForEach-Object { "$($_.sha256)  $($_.file)" })
    $checksums += "$((Get-FileHash -LiteralPath $manifestPath -Algorithm SHA256).Hash.ToLowerInvariant())  RELEASE-MANIFEST.json"
    [IO.File]::WriteAllText((Join-Path $release 'SHA256SUMS'), ($checksums -join "`n") + "`n")
    Assert-PackageOutputPath $output $root
    if ([IO.Path]::Exists($output)) { throw 'The final output appeared during the build; refusing to replace it.' }
    [IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($output)) | Out-Null
    $pending = "$output.$([Guid]::NewGuid().ToString('N')).staging"
    try {
        Copy-PackageItem $release $pending
        # Directory.Move fails if the final output already exists; never replace it.
        Assert-ForkSourceUnchanged
        [IO.Directory]::Move($pending, $output)
    } finally {
        if ([IO.Directory]::Exists($pending)) { Remove-PackageStaging $pending }
    }
    Write-Output "Fork release packages created in $output. No publication or runtime tests were performed."
}
