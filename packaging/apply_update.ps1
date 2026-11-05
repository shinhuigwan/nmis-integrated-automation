param([Parameter(Mandatory=$true)][string]$AppRoot,
      [Parameter(Mandatory=$true)][string]$PayloadRoot,
      [int]$ParentPid = 0, [switch]$NoRestart)
$ErrorActionPreference = 'Stop'
$AppRoot = [IO.Path]::GetFullPath($AppRoot).TrimEnd('\')
$PayloadRoot = [IO.Path]::GetFullPath($PayloadRoot).TrimEnd('\')
$StageRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$BackupRoot = Join-Path $StageRoot 'backup'
$FailedRoot = Join-Path $StageRoot 'failed'
$Names = @('NMIS-Automation.exe','_internal','installers','Setup.bat','Setup.ps1','Start.vbs','run_ui.vbs','apply_update.ps1',
    ([string]::Concat([char]0xC0AC,[char]0xC6A9,[char]0xBC29,[char]0xBC95,'.txt')),'THIRD_PARTY_LICENSES.txt','BUILD_INFO.json')
$Installed = [Collections.Generic.List[string]]::new()
$Saved = [Collections.Generic.List[string]]::new()
function Safe-Child([string]$Root, [string]$Name) {
    $Path = [IO.Path]::GetFullPath((Join-Path $Root $Name))
    if (-not $Path.StartsWith($Root.TrimEnd('\')+'\', [StringComparison]::OrdinalIgnoreCase)) { throw 'Unsafe update path.' }
    return $Path
}
function Move-Retry([string]$From, [string]$To) {
    for ($Attempt=0; $Attempt -lt 20; $Attempt++) {
        try { Move-Item -LiteralPath $From -Destination $To -ErrorAction Stop; return } catch {
            if ($Attempt -eq 19) { throw }
            Start-Sleep -Milliseconds 500
        }
    }
}
try {
    if (-not (Test-Path -LiteralPath $AppRoot -PathType Container) -or $AppRoot -eq [IO.Path]::GetPathRoot($AppRoot).TrimEnd('\')) { throw 'Invalid app directory.' }
    if ($PayloadRoot -eq $AppRoot -or $PayloadRoot.StartsWith($AppRoot+'\',[StringComparison]::OrdinalIgnoreCase) -or $AppRoot.StartsWith($PayloadRoot+'\',[StringComparison]::OrdinalIgnoreCase)) { throw 'Payload must be outside the application directory.' }
    if (-not (Test-Path -LiteralPath (Safe-Child $PayloadRoot 'NMIS-Automation.exe') -PathType Leaf) -or
        -not (Test-Path -LiteralPath (Safe-Child $PayloadRoot '_internal') -PathType Container)) { throw 'Incomplete update.' }
    if ($ParentPid -gt 0) {
        $Parent = Get-Process -Id $ParentPid -ErrorAction SilentlyContinue
        if ($Parent -and -not $Parent.WaitForExit(120000)) { throw 'The app did not close. Update was cancelled.' }
    }
    New-Item -ItemType Directory -Path $BackupRoot,$FailedRoot -Force | Out-Null
    foreach ($Name in $Names) {
        $Source = Safe-Child $PayloadRoot $Name
        if (-not (Test-Path -LiteralPath $Source)) { continue }
        $Target = Safe-Child $AppRoot $Name
        if (Test-Path -LiteralPath $Target) {
            Move-Retry $Target (Safe-Child $BackupRoot $Name)
            $Saved.Add($Name)
        }
        Move-Retry $Source $Target
        $Installed.Add($Name)
    }
    'Update completed. User settings and data were preserved.' | Set-Content -LiteralPath (Join-Path $AppRoot 'update.log')
    if (-not $NoRestart) {
        # Refresh both shortcuts used by prior source and portable editions.
        $Shell = New-Object -ComObject WScript.Shell
        $KoreanName = [string]::Concat([char]0xD1B5,[char]0xD569,' ',[char]0xC790,[char]0xB3D9,[char]0xD654,' ',[char]0xC2DC,[char]0xC2A4,[char]0xD15C,'.lnk')
        foreach ($Name in @('NMIS Integrated Automation.lnk',$KoreanName)) {
            $Shortcut = $Shell.CreateShortcut((Join-Path ([Environment]::GetFolderPath('Desktop')) $Name))
            $Shortcut.TargetPath = Join-Path $AppRoot 'NMIS-Automation.exe'
            $Shortcut.WorkingDirectory = $AppRoot
            $Shortcut.IconLocation = $Shortcut.TargetPath+',0'
            $Shortcut.Save()
        }
        Start-Process -FilePath (Join-Path $AppRoot 'NMIS-Automation.exe') -WorkingDirectory $AppRoot -WindowStyle Normal
    }
} catch {
    $Failure = $_.Exception.Message
    foreach ($Name in $Installed) {
        $Target = Safe-Child $AppRoot $Name
        if (Test-Path -LiteralPath $Target) { Move-Retry $Target (Safe-Child $FailedRoot $Name) }
    }
    foreach ($Name in $Saved) { Move-Retry (Safe-Child $BackupRoot $Name) (Safe-Child $AppRoot $Name) }
    $Failure | Set-Content -LiteralPath (Join-Path $StageRoot 'update-error.log')
    if ($NoRestart) { Write-Error $Failure; exit 1 }
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show('Update failed. Previous files were restored. '+$Failure,'NMIS Update') | Out-Null
    if (Test-Path -LiteralPath (Join-Path $AppRoot 'NMIS-Automation.exe')) {
        Start-Process -FilePath (Join-Path $AppRoot 'NMIS-Automation.exe') -WorkingDirectory $AppRoot -WindowStyle Normal
    }
    exit 1
}
