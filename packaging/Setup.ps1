param([switch]$VerifyOnly)
$ErrorActionPreference = 'Stop'
$AppRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not [Environment]::Is64BitOperatingSystem) { throw '64-bit Windows is required.' }
function Find-Chrome {
    foreach ($Candidate in @("$env:ProgramFiles\Google\Chrome\Application\chrome.exe", "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe", "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe")) {
        if (Test-Path -LiteralPath $Candidate) { return $Candidate }
    }
    return $null
}
$Chrome = Find-Chrome
Write-Host ('Google Chrome: ' + $(if ($Chrome) { $Chrome } else { 'Not installed' }))
foreach ($ProgId in @('Excel.Application', 'HWPFrame.HwpObject')) {
    Write-Host ($ProgId + ': ' + [bool][Type]::GetTypeFromProgID($ProgId))
}
if ($VerifyOnly) { exit 0 }
if (-not $Chrome) {
    $Installer = Join-Path $AppRoot 'installers\ChromeSetup.exe'
    $Signature = Get-AuthenticodeSignature -LiteralPath $Installer
    if ($Signature.Status -ne 'Valid' -or $Signature.SignerCertificate.Subject -notmatch 'Google LLC') {
        throw 'Google installer signature verification failed.'
    }
    Write-Host 'Installing Google Chrome. Internet access is required.'
    $Process = Start-Process -FilePath $Installer -ArgumentList '/silent /install' -WindowStyle Hidden -Wait -PassThru
    if ($Process.ExitCode -ne 0) { throw ('Chrome installer failed: ' + $Process.ExitCode) }
    $Chrome = Find-Chrome
    if (-not $Chrome) { throw 'Chrome was not found after installation. Rerun Setup.bat after installation completes.' }
}
$Exe = Join-Path $AppRoot 'NMIS-Automation.exe'
if (-not (Test-Path -LiteralPath $Exe)) { throw 'Extract the entire ZIP before running Setup.bat.' }
$Shell = New-Object -ComObject WScript.Shell
$KoreanName = [string]::Concat([char]0xD1B5,[char]0xD569,' ',[char]0xC790,[char]0xB3D9,[char]0xD654,' ',[char]0xC2DC,[char]0xC2A4,[char]0xD15C,'.lnk')
foreach ($Name in @($KoreanName,'NMIS Integrated Automation.lnk')) {
    $Shortcut = $Shell.CreateShortcut((Join-Path ([Environment]::GetFolderPath('Desktop')) $Name))
    $Shortcut.TargetPath = $Exe
    $Shortcut.WorkingDirectory = $AppRoot
    $Shortcut.IconLocation = "$Exe,0"
    $Shortcut.Save()
}
Write-Host 'Setup completed. Start NMIS-Automation.exe or the desktop shortcut.'
Write-Host 'Licensed Microsoft Excel and Hancom Hangul must be installed separately for their respective features.'
