
function New-Intro()
{
    Write-Host -ForegroundColor Yellow "Before running this troubleshooter, attempt to click the button on the login prompt that says 'If your connectivity is slow, select this option'"
    Write-Host -ForegroundColor Yellow "Before proceeding to the next level, try and have the user login after each step"
    Write-Host -ForegroundColor Yellow "Please choose carefully from the following menu"
    Write-Host "1. Delete local profile"
    Write-Host "2. Rebuild profile on the server"
    Write-Host "3. Exit"
}

function Stop-Dragon() {
    Stop-Process -Name "natspeak" -Confirm -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
}

$dragonPath = "\\MHC-MSASNDVNMS1\Profiles\McLaren\DMNEv2"

do {
    Clear-Host
    New-Intro
    $response = Read-Host "Enter selection"
} until ($response -eq 1 -or $response -eq 2 -or $response -eq 3 -or $response -eq 4)

if ($response -eq 1) {

    Clear-Host
    Stop-Dragon
    $users = $(Get-ChildItem -Path "C:\Users").Name
    Write-Host -ForegroundColor Red "Warning: This will delete the users locally cached profile from the computer under all user accounts"
    
    do {
        $username = Read-Host "Enter username"
        Write-Host -NoNewline "You entered $username, is this correct?"
        $answer = Read-Host " [Y/N]"
    } until ($answer -eq "Y" -or $answer -eq "y")

    Clear-Host

    # If this folder exists, it means that the profile has logged in to Dragon
    foreach ($user in $users) {
        if (Test-Path "C:\Users\$user\AppData\Roaming\Nuance") {
            Get-ChildItem -Path "C:\Users\$user\AppData\Roaming\Nuance\NaturallySpeaking12\Cache\$username*" | % {Write-Host $_.FullName}
        }
    }

    Write-Host -ForegroundColor Green "Found cache profiles"

    do {
        $confirmDeletion = Read-Host "Continue with deletion of the user profiles [Y/N]"
    } until ($confirmDeletion -eq "Y" -or $confirmDeletion -eq "y" -or $confirmDeletion -eq "N" -or $confirmDeletion -eq "n")

    if ($confirmDeletion -eq "Y" -or $confirmDeletion -eq "y") {

        Clear-Host
        Write-Host -ForegroundColor Red "Deleting local profiles"

        foreach ($user in $users) {
            if (Test-Path "C:\Users\$user\AppData\Roaming\Nuance") {
                Remove-Item -Path "C:\Users\$user\AppData\Roaming\Nuance\NaturallySpeaking12\Cache\$username*" -WhatIf
            }
        }

        Clear-Host
        Write-Host -ForegroundColor Green "Please have the user attempt to login"
    }

    if ($confirmDeletion -eq "N" -or $confirmDeletion -eq "n") {
        exit
    }
}

if ($response -eq 2) {

    Clear-Host
    Stop-Dragon
    $users = $(Get-ChildItem -Path "C:\Users").Name
    Write-Host -ForegroundColor Red "Warning: This will delete the users profile from the server and must be rebuilt"

    do {
        $username = Read-Host "Enter username"
        Write-Host -NoNewline "You entered $username, is this correct?"
        $answer = Read-Host " [Y/N]"
    } until ($answer -eq "Y" -or $answer -eq "y")

    Clear-Host

    Write-Host -ForegroundColor Green "Backing up macros to C:\Users\$env:username\Desktop\$username.dat"

    $directory = $(Get-ChildItem -Path "\\MHC-MSASNDVNMS1\Profiles\McLaren\DMNEv2\$username*\$username\current").FullName
    New-PSDrive -Name "A" -Root $directory -PSProvider "FileSystem"
    Clear-Host
    Push-Location -Path "A:"
    Copy-Item -Path "mycmds.dat" -Destination "C:\Users\$env:username\Desktop\$username.dat"
    Pop-Location
    Remove-PSDrive -Name "A"

    if (Test-Path "C:\Users\$env:username\Desktop\$username.dat") {
        Write-Host -ForegroundColor Green "Successfully backed up macros, proceeding with server profile deletion. This could take some time."
        $directoryToRemove = $(Get-ChildItem -Path "\\MHC-MSASNDVNMS1\Profiles\McLaren\DMNEv2\$username*").FullName
        Remove-Item -Path $directoryToRemove -WhatIf
    }

    Write-Host -ForegroundColor Green "Please have the user attempt login"
}

if ($response -eq 3) {
    exit
}
