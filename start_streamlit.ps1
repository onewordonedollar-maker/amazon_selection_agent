param(
    [Parameter(Mandatory = $true)]
    [string]$PythonExe,

    [Parameter(Mandatory = $true)]
    [string]$Root
)

$ErrorActionPreference = "Stop"
$App = Join-Path $Root "streamlit_app.py"

function Test-StreamlitReady {
    try {
        $Response = Invoke-WebRequest `
            -UseBasicParsing `
            -Uri "http://127.0.0.1:8501/" `
            -TimeoutSec 1
        return $Response.StatusCode -eq 200
    }
    catch {
        return $false
    }
}

try {
    if (-not (Test-StreamlitReady)) {
        $Info = New-Object System.Diagnostics.ProcessStartInfo
        $Info.FileName = $PythonExe
        $Info.WorkingDirectory = $Root
        $Info.Arguments = (
            '-m streamlit run "{0}" --server.port 8501 --server.headless true' -f $App
        )
        $Info.UseShellExecute = $true
        $Info.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden

        $Process = [System.Diagnostics.Process]::Start($Info)
        if ($null -eq $Process) {
            throw "Windows did not create the Streamlit process."
        }
    }

    $Deadline = (Get-Date).AddSeconds(30)
    while ((Get-Date) -lt $Deadline) {
        if (Test-StreamlitReady) {
            exit 0
        }
        Start-Sleep -Seconds 1
    }
    throw "Streamlit did not become ready within 30 seconds."
}
catch {
    Write-Error $_
    exit 1
}
