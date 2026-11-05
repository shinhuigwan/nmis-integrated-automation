param(
    [string]$Python = 'python',
    [string]$OutputDirectory = (Join-Path $PSScriptRoot 'dist')
)
$ErrorActionPreference = 'Stop'
Push-Location $PSScriptRoot
try {
    $ArgsList = @('-m','PyInstaller','--noconfirm','--onedir','--windowed',
        '--name','NMIS-Automation','--icon',(Join-Path $PSScriptRoot 'assets\nmis_automation.ico'),
        '--add-data',((Join-Path $PSScriptRoot 'assets') + ';assets'),
        '--add-data',((Join-Path $PSScriptRoot 'hancom_security') + ';hancom_security'),
        '--collect-all','customtkinter','--collect-all','playwright','--collect-all','reportlab',
        '--collect-submodules','win32com','--hidden-import','win32timezone','--hidden-import','win32crypt',
        '--hidden-import','pythoncom','--hidden-import','pywintypes',
        '--distpath',$OutputDirectory,'--workpath',(Join-Path $OutputDirectory '_build'),
        '--specpath',(Join-Path $OutputDirectory '_spec'),(Join-Path $PSScriptRoot 'portable_launcher.py'))
    & $Python @ArgsList
    if ($LASTEXITCODE -ne 0) { throw 'PyInstaller build failed.' }
    $Bundle = Join-Path $OutputDirectory 'NMIS-Automation'
    Copy-Item -Path (Join-Path $PSScriptRoot 'packaging\*') -Destination $Bundle -Force
    & $Python (Join-Path $PSScriptRoot 'tools\collect_licenses.py') (Join-Path $Bundle 'THIRD_PARTY_LICENSES.txt')
    if ($LASTEXITCODE -ne 0) { throw 'License collection failed.' }
    $InstallerDirectory = Join-Path $Bundle 'installers'
    New-Item -ItemType Directory -Path $InstallerDirectory -Force | Out-Null
    $Installer = Join-Path $InstallerDirectory 'ChromeSetup.exe'
    Invoke-WebRequest -Uri 'https://dl.google.com/chrome/install/latest/chrome_installer.exe' -OutFile $Installer
    $Signature = Get-AuthenticodeSignature -LiteralPath $Installer
    if ($Signature.Status -ne 'Valid' -or $Signature.SignerCertificate.Subject -notmatch 'Google LLC') {
        throw 'Chrome installer publisher verification failed.'
    }
    Write-Host "Portable bundle: $Bundle"
} finally {
    Pop-Location
}
