# An attempt to logically place troubleshooting steps into a progressive script for fixing Dragon issues

$users = $(Get-ChildItem -Path "C:\Users").Name

function Stop-Dragon() {
    Stop-Process -Name "natspeak" -Confirm -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
}

function Step-Two() {
    Clear-Host
    Write-Host -ForegroundColor Green "Step 2. Attempting to delete local cache files"
    Stop-Dragon

    do {
        $username = Read-Host "Enter user's username"
        Write-Host -NoNewline "You entered $username, is this correct?"
        $answer = Read-Host " [Y/N]"
    } until ($answer -eq "Y" -or $answer -eq "y")

    Clear-Host

    [array]$directories = $null

    if ($answer -eq "Y" -or $answer -eq "y") {
        # If this folder exists, it means that the profile has logged in to Dragon
        foreach ($user in $users) {
            if (Test-Path "C:\Users\$user\AppData\Roaming\Nuance") {
                $directories += $(Get-ChildItem -Path "C:\Users\$user\AppData\Roaming\Nuance\NaturallySpeaking12\Cache\$username*").FullName
            }
        }

        if ($directories.Count -eq 0) {
            Write-Host -ForegroundColor Yellow "No locally cached profiles found, proceeding to next step."
            Start-Sleep 2
            Step-Three
        }

        if ($directories.Count -gt 0) {
            Write-Host -ForegroundColor Red "Directories to be deleted"

            # Creates a custom PSObject to print directories in a grid
            $data = foreach ($directory in $directories) {
                if ($directory) {
                    [PSCustomObject]@{
                        "Directories" = $directory
                    }
                }
            }

            Write-Host ($data | Format-Table | Out-String)

            do {
                $confirmDeletion = Read-Host "Continue with deletion of the user profiles [Y/N]"
            } until ($confirmDeletion -eq "Y" -or $confirmDeletion -eq "y" -or $confirmDeletion -eq "N" -or $confirmDeletion -eq "n")
    
            if ($confirmDeletion -eq "Y" -or $confirmDeletion -eq "y") {
    
                Clear-Host
                Write-Host -ForegroundColor Red "Deleting local profiles"
        
                # We can probably use our object to delete, look into removing this line
                foreach ($user in $users) {
                    if (Test-Path "C:\Users\$user\AppData\Roaming\Nuance") {
                        Remove-Item -Path "C:\Users\$user\AppData\Roaming\Nuance\NaturallySpeaking12\Cache\$username*" -WhatIf
                    }
                }

                do {
                    $confirmAnswer = Read-Host -Prompt "Have user attempt to login. Did this solve the issue? [y/n]"
                } until ($confirmAnswer -eq "Y" -or $confirmAnswer -eq "y" -or $confirmAnswer -eq "N" -or $confirmAnswer -eq "n")

                if ($confirmAnswer -eq "Y" -or $confirmAnswer -eq "y") {
                    exit
                }

                if ($confirmAnswer -eq "N" -or $confirmAnswer -eq "n") {
                    Step-Three
                }
            }
        }
    }
}

function Step-Three() {
    Clear-Host
    Write-Host "Step 3. Attempting to restore from backup"
    Stop-Dragon

    # Remove this as $username variable should still be in memory
    #do {
    #    $username = Read-Host "Enter username"
    #    Write-Host -NoNewline "You entered $username, is this correct?"
    #    $answer = Read-Host " [Y/N]"
    #} until ($answer -eq "Y" -or $answer -eq "y")

    #Clear-Host

    $lastKnownGood = $(Get-ChildItem -Path "\\MHC-MSASNDVNMS1\Profiles\McLaren\DMNEv2\$username*\last_known_good").FullName
    $directoryToReplace = $(Get-ChildItem -Path "\\MHC-MSASNDVNMS1\Profiles\McLaren\DMNEv2\$username*\last_known_good\$username").FullName
    $directoryToDelete = $(Get-ChildItem -Path "\\MHC-MSASNDVNMS1\Profiles\McLaren\DMNEv2\$username*\$username").FullName
    $profileFolder = $(Get-ChildItem -Path "\\MHC-MSASNDVNMS1\Profiles\McLaren\DMNEv2\$username*").FullName

    if (Test-Path $lastKnownGood) {
        Write-Host -ForegroundColor Green "Found a profile backup, attempting to restore"
        New-PSDrive -Name "A" -Root $lastKnownGood -PSProvider "FileSystem" >$null
        Clear-Host
        Push-Location -Path "A:"
        Remove-Item -Path $directoryToDelete -WhatIf
        Copy-Item -Path $directoryToReplace -Destination $profileFolder -WhatIf
        Pop-Location
        Remove-PSDrive -Name "A"

        do {
            $confirmAnswer = Read-Host -Prompt "Have user attempt to login. Did this solve the issue? [y/n]"
        } until ($confirmAnswer -eq "Y" -or $confirmAnswer -eq "y" -or $confirmAnswer -eq "N" -or $confirmAnswer -eq "n")

        if ($confirmAnswer -eq "Y" -or $confirmAnswer -eq "y") {
            exit
        }

        if ($confirmAnswer -eq "N" -or $confirmAnswer -eq "n") {
            Step-Four
        }
    }
}

function Step-Four() {
    Stop-Dragon
    Clear-Host
    Write-Host -ForegroundColor Red "Step 4. Deleting master roaming user profile"
    $macroDirectory = $(Get-ChildItem -Path "\\MHC-MSASNDVNMS1\Profiles\McLaren\DMNEv2\$username*\$username\current").FullName
    Write-Host -ForegroundColor Green "Backing up macros to C:\Users\$env:username\Desktop\$username.dat"

    New-PSDrive -Name "B" -Root $macroDirectory -PSProvider "FileSystem" >$null
    Push-Location -Path "B:"
    Copy-Item -Path "mycmds.dat" -Destination "C:\Users\$env:username\Desktop\$username.dat"
    Pop-Location
    Remove-PSDrive -Name "B"

    if (Test-Path "C:\Users\$env:username\Desktop\$username.dat") {
        Write-Host -ForegroundColor Green "Successfully backed up macros, proceeding with server profile deletion. This could take some time."
        $directoryToRemove = $(Get-ChildItem -Path "\\MHC-MSASNDVNMS1\Profiles\McLaren\DMNEv2\$username*").FullName
        Remove-Item -Path $directoryToRemove -WhatIf

        do {
            $confirmAnswer = Read-Host -Prompt "Have user attempt to login. Did this solve the issue? [y/n]"
        } until ($confirmAnswer -eq "Y" -or $confirmAnswer -eq "y" -or $confirmAnswer -eq "N" -or $confirmAnswer -eq "n")
    
        if ($confirmAnswer -eq "Y" -or $confirmAnswer -eq "y") {
            exit
        }
    
        if ($confirmAnswer -eq "N" -or $confirmAnswer -eq "n") {
            Clear-Host
            Write-Host -ForegroundColor Red "It appears something went wrong. Continue manual troubleshooting."
        }
    }

    Write-Host -ForegroundColor Red "Unable to verify that the macros were backed up, exiting."
    exit
}

# Beginning of the program
Clear-Host
Write-Host -ForegroundColor Yellow "Welcome to the interactive troubleshooter, press enter to continue."
Read-Host

# Beginning of Step 1
# Attempt to kill Dragon
# TODO: Verify that the process is stopped
Clear-Host
Write-Host -NoNewline -ForegroundColor Yellow "Attempting to exit Dragon"
Start-Sleep 1
Write-Host -NoNewline -ForegroundColor Yellow "."
Start-Sleep 1
Write-Host -NoNewline -ForegroundColor Yellow "."
Start-Sleep 1
Write-Host -ForegroundColor Yellow "."
Write-Host -ForegroundColor Green "Step 1. Attempt to have the user login to Dragon by clicking on the 'If you know your network connectivity is slow' checkbox at the login screen"
Stop-Dragon

do {
    $answer = Read-Host -Prompt "Did this solve the issue? [Y/N]"
} until ($answer -eq "Y" -or $answer -eq "y" -or $answer -eq "N" -or $answer -eq "n")

if ($answer -eq "Y" -or $answer -eq "y") {
    exit
}

if ($answer -eq "N" -or $answer -eq "n") {
    Step-Two
}
