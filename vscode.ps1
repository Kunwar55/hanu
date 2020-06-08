[CmdletBinding()]
param(
    [Parameter()]
    [ValidateNotNull()]
    [string[]]$AdditionalExtensions = @(),

    [switch]$LaunchWhenDone
)

if (!($IsLinux -or $IsOSX)) {

    $codeCmdPath = "C:\Program Files (x86)\Microsoft VS Code\bin\code.cmd"

    try {
        $ProgressPreference = 'SilentlyContinue'

        if (!(Test-Path $codeCmdPath)) {
            Write-Host "`nDownloading latest stable Visual Studio Code..." -ForegroundColor Yellow
            Remove-Item -Force $env:TEMP\vscode-stable.exe -ErrorAction SilentlyContinue
            Invoke-WebRequest -Uri https://vscode-update.azurewebsites.net/latest/win32/stable -OutFile $env:TEMP\vscode-stable.exe

            Write-Host "`nInstalling Visual Studio Code..." -ForegroundColor Yellow
            Start-Process -Wait $env:TEMP\vscode-stable.exe -ArgumentList /silent, /mergetasks=!runcode
        }
        else {
            Write-Host "`nVisual Studio Code is already installed." -ForegroundColor Yellow
        }

        $extensions = @("ms-vscode.PowerShell") + $AdditionalExtensions
        foreach ($extension in $extensions) {
            Write-Host "`nInstalling extension $extension..." -ForegroundColor Yellow
            & $codeCmdPath --install-extension $extension
        }

        if ($LaunchWhenDone) {
            Write-Host "`nInstallation complete, starting Visual Studio Code...`n`n" -ForegroundColor Green
            & $codeCmdPath
        }
        else {
            Write-Host "`nInstallation complete!`n`n" -ForegroundColor Green
        }
    }
    finally {
        $ProgressPreference = 'Continue'
    }
}
else {
    Write-Error "This script is currently only supported on the Windows operating system."
}
