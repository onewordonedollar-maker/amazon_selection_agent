param(
    [switch]$Repair
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$RuntimeRoot = Join-Path $Root ".runtime"
$PythonDir = Join-Path $RuntimeRoot "python"
$PythonExe = Join-Path $PythonDir "python.exe"
$Requirements = Join-Path $Root "requirements.txt"
$PythonVersion = "3.13.14"
$InstallerUrl = "https://www.python.org/ftp/python/$PythonVersion/python-$PythonVersion-amd64.exe"
$Installer = Join-Path $env:TEMP "amazon-selection-agent-python-$PythonVersion.exe"

try {
    New-Item -ItemType Directory -Force -Path $RuntimeRoot | Out-Null

    if (-not (Test-Path $PythonExe)) {
        Write-Host "Downloading project-local Python $PythonVersion..."
        Invoke-WebRequest -UseBasicParsing -Uri $InstallerUrl -OutFile $Installer

        Write-Host "Installing Python into $PythonDir..."
        $Arguments = @(
            "/quiet",
            "InstallAllUsers=0",
            "TargetDir=$PythonDir",
            "PrependPath=0",
            "Include_launcher=0",
            "Include_test=0",
            "Include_pip=1",
            "Shortcuts=0",
            "AssociateFiles=0"
        )
        $Process = Start-Process -FilePath $Installer -ArgumentList $Arguments -Wait -PassThru
        if ($Process.ExitCode -ne 0 -or -not (Test-Path $PythonExe)) {
            throw "Python installer exited with code $($Process.ExitCode)."
        }
    }

    if ($Repair) {
        Write-Host "Repairing project dependencies..."
    } else {
        Write-Host "Installing project dependencies..."
    }
    & $PythonExe -m pip install --disable-pip-version-check -r $Requirements
    if ($LASTEXITCODE -ne 0) {
        throw "pip install failed with exit code $LASTEXITCODE."
    }

    & $PythonExe -c "import streamlit, openpyxl; print('Project runtime ready.')"
    if ($LASTEXITCODE -ne 0) {
        throw "The project runtime verification failed."
    }
}
catch {
    Write-Error $_
    exit 1
}
finally {
    if (Test-Path $Installer) {
        Remove-Item -LiteralPath $Installer -Force -ErrorAction SilentlyContinue
    }
}
