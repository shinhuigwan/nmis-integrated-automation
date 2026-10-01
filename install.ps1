param(
    [switch]$VerifyOnly
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$AppRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$VenvRoot = Join-Path $AppRoot ".venv"
$VenvPython = Join-Path $VenvRoot "Scripts\python.exe"
$Requirements = Join-Path $AppRoot "requirements.txt"
$Launcher = Join-Path $AppRoot "run_ui.vbs"
$IconPath = Join-Path $AppRoot "assets\nmis_automation.ico"
$Desktop = [Environment]::GetFolderPath("Desktop")
$ShortcutPath = Join-Path $Desktop "통합 자동화 시스템.lnk"

function Write-Step([string]$Message) {
    Write-Host "`n>> $Message" -ForegroundColor Cyan
}

function Test-Python([string]$Path) {
    if (-not $Path -or -not (Test-Path -LiteralPath $Path)) {
        return $false
    }
    try {
        & $Path -c "import sys; raise SystemExit(0 if sys.version_info >= (3, 10) else 1)" 2>$null
        return $LASTEXITCODE -eq 0
    } catch {
        return $false
    }
}

function Find-Python {
    $Candidates = [System.Collections.Generic.List[string]]::new()
    foreach ($Name in @("python.exe", "python3.exe")) {
        $Command = Get-Command $Name -ErrorAction SilentlyContinue
        if ($Command -and $Command.Source) {
            $Candidates.Add($Command.Source)
        }
    }
    foreach ($Pattern in @(
        "$env:LOCALAPPDATA\Programs\Python\Python*\python.exe",
        "$env:ProgramFiles\Python*\python.exe",
        "${env:ProgramFiles(x86)}\Python*\python.exe"
    )) {
        Get-ChildItem -Path $Pattern -File -ErrorAction SilentlyContinue |
            Sort-Object FullName -Descending |
            ForEach-Object { $Candidates.Add($_.FullName) }
    }
    foreach ($Candidate in ($Candidates | Select-Object -Unique)) {
        if (Test-Python $Candidate) {
            return $Candidate
        }
    }
    return $null
}

function Install-Python {
    Write-Step "Installing Python 3.12"
    $Winget = Get-Command winget.exe -ErrorAction SilentlyContinue
    if ($Winget) {
        & $Winget.Source install --id Python.Python.3.12 --exact --source winget `
            --scope user --accept-package-agreements --accept-source-agreements --silent
        if ($LASTEXITCODE -ne 0) {
            throw "Python installation through winget failed."
        }
        return
    }

    $Installer = Join-Path $env:TEMP "python-3.12-amd64.exe"
    Invoke-WebRequest -Uri "https://www.python.org/ftp/python/3.12.9/python-3.12.9-amd64.exe" -OutFile $Installer
    $Process = Start-Process -FilePath $Installer -ArgumentList "/quiet InstallAllUsers=0 PrependPath=1 Include_test=0" -Wait -PassThru
    if ($Process.ExitCode -ne 0) {
        throw "Python installer returned exit code $($Process.ExitCode)."
    }
}

function Find-Chrome {
    foreach ($Path in @(
        "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
        "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
        "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
    )) {
        if ($Path -and (Test-Path -LiteralPath $Path)) {
            return $Path
        }
    }
    return $null
}

function Install-Chrome {
    Write-Step "Installing Google Chrome"
    $Winget = Get-Command winget.exe -ErrorAction SilentlyContinue
    if ($Winget) {
        & $Winget.Source install --id Google.Chrome --exact --source winget `
            --accept-package-agreements --accept-source-agreements --silent
        if ($LASTEXITCODE -ne 0) {
            throw "Chrome installation through winget failed."
        }
        return
    }

    $Installer = Join-Path $env:TEMP "chrome_installer.exe"
    Invoke-WebRequest -Uri "https://dl.google.com/chrome/install/latest/chrome_installer.exe" -OutFile $Installer
    $Process = Start-Process -FilePath $Installer -ArgumentList "/silent /install" -Wait -PassThru
    if ($Process.ExitCode -ne 0) {
        throw "Chrome installer returned exit code $($Process.ExitCode)."
    }
}

function Test-ComProgram([string]$ProgId) {
    try {
        return $null -ne [Type]::GetTypeFromProgID($ProgId)
    } catch {
        return $false
    }
}

if (-not (Test-Path -LiteralPath $Requirements)) {
    throw "requirements.txt was not found: $Requirements"
}

Write-Step "Checking prerequisites"
$Python = Find-Python
$Chrome = Find-Chrome
Write-Host ("Python 3.10 or newer: " + $(if ($Python) { $Python } else { "Not found" }))
Write-Host ("Google Chrome: " + $(if ($Chrome) { $Chrome } else { "Not found" }))

if ($VerifyOnly) {
    Write-Host "Verification-only run completed." -ForegroundColor Green
    exit 0
}

if (-not $Python) {
    Install-Python
    $Python = Find-Python
    if (-not $Python) {
        throw "python.exe was not found after installation. Sign in to Windows again and rerun the installer."
    }
}

if (-not $Chrome) {
    Install-Chrome
    $Chrome = Find-Chrome
    if (-not $Chrome) {
        throw "chrome.exe was not found after installation."
    }
}

Write-Step "Creating the project Python virtual environment"
if (-not (Test-Path -LiteralPath $VenvPython)) {
    & $Python -m venv $VenvRoot
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to create the Python virtual environment."
    }
}

Write-Step "Installing required Python packages"
& $VenvPython -m pip install --disable-pip-version-check --upgrade pip setuptools wheel
if ($LASTEXITCODE -ne 0) {
    throw "Failed to update pip."
}
& $VenvPython -m pip install --disable-pip-version-check -r $Requirements
if ($LASTEXITCODE -ne 0) {
    throw "Failed to install the required Python packages."
}

Write-Step "Verifying the installation"
& $VenvPython -c "import customtkinter, openpyxl, pandas, playwright, pypdf, reportlab, olefile, win32com.client, xlrd, xlutils"
if ($LASTEXITCODE -ne 0) {
    throw "Required Python package verification failed."
}

Write-Step "Creating a desktop shortcut"
$Shell = New-Object -ComObject WScript.Shell
$Shortcut = $Shell.CreateShortcut($ShortcutPath)
$Shortcut.TargetPath = "$env:WINDIR\System32\wscript.exe"
$Shortcut.Arguments = '"' + $Launcher + '"'
$Shortcut.WorkingDirectory = $AppRoot
$Shortcut.IconLocation = $(if (Test-Path -LiteralPath $IconPath) { "$IconPath,0" } else { "$env:WINDIR\System32\shell32.dll,0" })
$Shortcut.Save()

Write-Host "`nInstallation completed." -ForegroundColor Green
Write-Host "Desktop shortcut: $ShortcutPath"
Write-Host "Chrome: $Chrome"

if (-not (Test-ComProgram "Excel.Application")) {
    Write-Warning "Microsoft Excel was not detected. Monthly-report and Excel COM features require licensed Microsoft Excel."
}
if (-not (Test-ComProgram "HWPFrame.HwpObject")) {
    Write-Warning "Hancom Office Hangul was not detected. HWP receipt-number entry requires licensed Hancom Office."
}

Write-Host "Naver Mail requires IMAP and an application password to be configured once on each PC." -ForegroundColor Yellow
exit 0
