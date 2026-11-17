# Create a sample workbook

Use the Windows x64 CLI ZIP from the [fork releases](https://github.com/Antrakt92/mcp-server-excel/releases/latest). Run this in PowerShell from the extracted CLI directory on a Windows desktop with Excel installed.

The script creates a new workbook in your temporary directory, adds a worksheet, and saves the workbook. Each command uses the same open Excel session.

```powershell
$path = Join-Path $env:TEMP ("ExcelMcp-demo-{0:N}.xlsx" -f [guid]::NewGuid())
$created = .\excelcli.exe -q session create $path | ConvertFrom-Json
if (-not $created.success) { throw $created.error.message }

$session = $created.data.sessionId
$save = $false
try {
    .\excelcli.exe -q sheet create --session $session --sheet-name Sales
    if ($LASTEXITCODE -ne 0) { throw 'Adding the Sales sheet failed.' }

    .\excelcli.exe -q sheet list --session $session
    if ($LASTEXITCODE -ne 0) { throw 'Sheet listing failed.' }

    $save = $true
}
finally {
    if ($save) {
        .\excelcli.exe -q session close --session $session --save
    }
    else {
        .\excelcli.exe -q session close --session $session
    }
    $closeExit = $LASTEXITCODE
    if ($closeExit -ne 0 -and -not $save) {
        Write-Warning "Closing failed. Check session $session before retrying."
    }
}
if ($closeExit -ne 0) { throw 'Closing the sample workbook failed.' }
$path
```

Use the printed path in the [read-only audit example](../README.md#try-a-read-only-audit) to list the saved worksheets and check the workbook hash. You can delete the sample workbook after closing the audit session.
