Remove-Item $env:TEMP\*.* -ErrorAction SilentlyContinue
Remove-Item "C:\Windows\Prefetch\*.*" -ErrorAction SilentlyContinue
Remove-Item "C:\Windows\Temp\*.*" -ErrorAction SilentlyContinue
Remove-Item "$env:LOCALAPPDATA\Microsoft\Windows\History\*.*" -ErrorAction SilentlyContinue
Remove-Item "$env:LOCALAPPDATA\Microsoft\Windows\WebCache\*.*" -ErrorAction SilentlyContinue
Remove-Item "$env:LOCALAPPDATA\Microsoft\Windows\Temporary Internet Files\*.*" -ErrorAction SilentlyContinue
Remove-Item "$env:APPDATA\Microsoft\Windows\Cookies\Low\*.*" -ErrorAction SilentlyContinue
Remove-Item "$env:APPDATA\Microsoft\Windows\Cookies\*.*" -ErrorAction SilentlyContinue

$ExcludedUsers = "Public","Administrator","ADMINI~1"

$LocalProfiles = $(Get-ChildItem -Path "C:\Users" | Where {($_.LastAccessTime) -lt (Get-Date).AddDays(-2)}).Name

$Folders = (Get-ChildItem -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList").PSChildName



foreach ($LocalProfile in $LocalProfiles) {
    if(!($ExcludedUsers -like $LocalProfile)) {
        Write-Host "Deleting profile $LocalProfile"
        Remove-Item -Path "C:\Users\$LocalProfile" -Force -WhatIf
        foreach ($Key in $Folders) {
            $Values = $(Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$Key").ProfileImagePath
            if ($Values -match $LocalProfile) {
                Remove-Item -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$Key" -WhatIf
            }
        }
    }
}

