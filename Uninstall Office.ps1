# Uninstall Microsoft 365
$Microsoft365Versions = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*Microsoft 365*").DisplayVersion
$Microsoft365ProductCodes = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*Microsoft 365*").PSChildName

if ($Microsoft365Versions) {
    foreach ($ProductCode in $Microsoft365ProductCodes) {
        Start-Process "msiexec.exe" -ArgumentList "/x $ProductCode /passive /norestart" -Wait
    }
}

# Uninstall OneNote
$OneNoteVersions = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*OneNote*").DisplayVersion
$OneNoteProductCodes = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*OneNote*").PSChildName

if ($OneNoteVersions) {
    foreach ($ProductCode in $OneNoteProductCodes) {
        Start-Process "msiexec.exe" -ArgumentList "/x $ProductCode /passive /norestart" -Wait
    }
}

# Uninstall Office
$OfficeVersions = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*Office*").DisplayVersion
$OfficeProductCodes = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*Office*").PSChildName

if ($OfficeVersions) {
    if ($OfficeVersions -eq "16.0.13628.20274" -or $OfficeVersions -eq "16.0.14131.20278") {
        # Office 365 / Office 2019 / Office 2021
        foreach ($ProductCode in $OfficeProductCodes) {
            Start-Process "msiexec.exe" -ArgumentList "/x $ProductCode /passive /norestart" -Wait
        }
    } elseif ($OfficeVersions -eq "16.0.4266.1001") {
        # Office 2016
        Start-Process "cscript.exe" -ArgumentList "$env:windir\system32\scrrun.dll //i //s `"$env:ProgramFiles\Common Files\Microsoft Shared\OFFICE16\Office Setup Controller\setup.xml`"" -Wait
    } else {
        Write-Host "Unsupported Office version: $OfficeVersions"
    }
}

# Uninstall Microsoft 365 apps
Get-AppxPackage Microsoft.365.* | Remove-AppxPackage

# Uninstall OneNote app
Get-AppxPackage Microsoft.Office.OneNote* | Remove-AppxPackage