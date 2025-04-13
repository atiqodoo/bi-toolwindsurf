$DesktopPath = [Environment]::GetFolderPath("Desktop")
$ShortcutPath = Join-Path $DesktopPath "Paint Analytics.lnk"
$TargetPath = Join-Path $PSScriptRoot "run_app.bat"
$IconPath = Join-Path $env:SystemRoot "System32\imageres.dll"

$WScriptShell = New-Object -ComObject WScript.Shell
$Shortcut = $WScriptShell.CreateShortcut($ShortcutPath)
$Shortcut.TargetPath = $TargetPath
$Shortcut.WorkingDirectory = $PSScriptRoot
$Shortcut.IconLocation = "$IconPath,63"  # Paint-like icon
$Shortcut.Description = "Paint Retail Analytics Application"
$Shortcut.Save()

Write-Host "Shortcut created on desktop: 'Paint Analytics'"
