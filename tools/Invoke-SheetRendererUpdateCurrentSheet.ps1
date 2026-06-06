[CmdletBinding()]
param(
    [string]$CommandName = "SheetRenderer_UpdateCurrentSheet"
)

$ErrorActionPreference = "Stop"

try {
    $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
}
catch {
    Write-Error "Running Excel instance was not found."
    exit 1
}

try {
    $null = $excel.Run($CommandName)
}
catch {
    Write-Error "Failed to run '$CommandName'. Make sure the SheetRenderer Excel-DNA add-in is loaded. $($_.Exception.Message)"
    exit 1
}
