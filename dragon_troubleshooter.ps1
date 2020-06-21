function New-InteractiveTroubleshooter() {
    # An attempt to logically place troubleshooting steps into a progressive script for fixing Dragon issues

    # This sets the users variable so we can recursively search for cached profiles
    # since some genius decided to store them on each individual user's account
    $users = $(Get-ChildItem -Path "C:\Users").Name

    # This function helps to kill the Dragon process "natspeak.exe" automagically so brainlets don't have to think too hard
    function Stop-Dragon() {
        Stop-Process -Name "natspeak" -Confirm -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    }

    # Step two is for deleting the local profile, it does this in every. Single. User. Folder.
    # It takes no arguments and returns no values since the administrator will dictate if this resolves the issue or not.
    function Step-Two() {
        Clear-Host
        Write-Host -ForegroundColor Green "Step 2. Attempting to delete local cache files"
        Stop-Dragon

        # Small logical loop to wait until proper username is entered
        do {
            $username = Read-Host "Enter user's username"
            Write-Host -NoNewline "You entered $username, is this correct?"
            $answer = Read-Host " [Y/N]"
        } until ($answer -eq "Y" -or $answer -eq "y")

        Clear-Host

        # Initialize an empty array so that we can push a list of directories onto the stack for manipulation later
        [array]$directories = $null

        # If someone answers yes, then we can proceed to loop through every user on the system and store that cached directory profile in the array
        if ($answer -eq "Y" -or $answer -eq "y") {
            # If this folder exists, it means that the profile has logged in to Dragon
            foreach ($user in $users) {
                if (Test-Path "C:\Users\$user\AppData\Roaming\Nuance") {
                    $directories += $(Get-ChildItem -Path "C:\Users\$user\AppData\Roaming\Nuance\NaturallySpeaking12\Cache\$username*").FullName
                }
            }

            # If for some reason there is no cache, skip the rest of the steps and proceed directly to Step 3
            if ($directories.Count -eq 0) {
                Write-Host -ForegroundColor Yellow "No locally cached profiles found, proceeding to next step."
                Start-Sleep 2
                Step-Three
            }

            # If there is at least 1 directory, lets try and delete it to force Dragon to redownload the master profile
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

                # Final warning prompt before deletion
                do {
                    $confirmDeletion = Read-Host "Continue with deletion of the user profiles [Y/N]"
                } until ($confirmDeletion -eq "Y" -or $confirmDeletion -eq "y" -or $confirmDeletion -eq "N" -or $confirmDeletion -eq "n")
    
                # Once accepted, deletes the cached directories from all of the user directories stored in the path
                if ($confirmDeletion -eq "Y" -or $confirmDeletion -eq "y") {
    
                    Clear-Host
                    Write-Host -ForegroundColor Red "Deleting local profiles"
        
                    # We can probably use our object to delete, look into removing this line
                    # TODO: Research PSObject loops so we do not have to test AGAIN if the folder exists for deletion
                    foreach ($user in $users) {
                        if (Test-Path "C:\Users\$user\AppData\Roaming\Nuance") {
                            Remove-Item -Path "C:\Users\$user\AppData\Roaming\Nuance\NaturallySpeaking12\Cache\$username*" -WhatIf
                        }
                    }

                    # The rest is just conditional logic to control program execution flow
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

    # Step three normally involves see if there is a backup of the master profile
    function Step-Three() {

        Clear-Host
        Write-Host "Step 3. Attempting to restore from backup"
        Stop-Dragon

        # These are variables of the folder structure on the master server
        # TODO: Possibly consider renaming these better?
        $lastKnownGood = $(Get-ChildItem -Path "\\MHC-MSASNDVNMS1\Profiles\McLaren\DMNEv2\$username*\last_known_good").FullName
        $directoryToReplace = $(Get-ChildItem -Path "\\MHC-MSASNDVNMS1\Profiles\McLaren\DMNEv2\$username*\last_known_good\$username").FullName
        $directoryToDelete = $(Get-ChildItem -Path "\\MHC-MSASNDVNMS1\Profiles\McLaren\DMNEv2\$username*\$username").FullName
        $profileFolder = $(Get-ChildItem -Path "\\MHC-MSASNDVNMS1\Profiles\McLaren\DMNEv2\$username*").FullName

        # First, we should test and see if there is a folder "last_known_good" since that is where the backup profile is stored
        # If not, we should just continue to Step 4 since it's a waste of time.
        if (Test-Path $lastKnownGood) {
            Write-Host -ForegroundColor Green "Found a profile backup, attempting to restore"

            # Creates a immutable drive so we can push that onto the stack as a "directory" and then just grab the folders we need
            # It then removes that drive so we don't accidentally mess with the server
            # TODO: Possibly look into making sure that the drive was removed before proceeding, and if not reboot since these are not persistant
            New-PSDrive -Name "A" -Root $lastKnownGood -PSProvider "FileSystem" >$null
            Clear-Host
            Push-Location -Path "A:"
            Remove-Item -Path $directoryToDelete -WhatIf
            Copy-Item -Path $directoryToReplace -Destination $profileFolder -WhatIf
            Pop-Location
            Remove-PSDrive -Name "A"

            # Just some more control logic to direct flow
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

    # Step 4 is basically just blowing the profile from the server away
    # First though, we should backup the users macros so they do not get angry
    # No one likes an angry person
    function Step-Four() {
        Stop-Dragon
        Clear-Host
        Write-Host -ForegroundColor Red "Step 4. Deleting master roaming user profile"

        # This is the directory where macros are stored. 
        $macroDirectory = $(Get-ChildItem -Path "\\MHC-MSASNDVNMS1\Profiles\McLaren\DMNEv2\$username*\$username\current").FullName
        Write-Host -ForegroundColor Green "Backing up macros to C:\Users\$env:username\Desktop\$username.dat"

        # Again, this is just copying the .dat file to the user's desktop so that after deletion, we can import all their macros.
        # TODO: Look into finding the vocabulary file as well for backup.
        New-PSDrive -Name "B" -Root $macroDirectory -PSProvider "FileSystem" >$null
        Push-Location -Path "B:"
        Copy-Item -Path "mycmds.dat" -Destination "C:\Users\$env:username\Desktop\$username.dat"
        Pop-Location
        Remove-PSDrive -Name "B"

        # This is to double check and MAKE SURE we have the macros backed up, otherwise we just exit. No angry user's, remember?
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
}
function New-Intro() {
    Write-Host -ForegroundColor Yellow "Before running this troubleshooter, attempt to click the button on the login prompt that says 'If your connectivity is slow, select this option'"
    Write-Host -ForegroundColor Yellow "Before proceeding to the next level, try and have the user login after each step"
    Write-Host -ForegroundColor Yellow "Please choose carefully from the following menu"
    Write-Host "1. Delete local profile"
    Write-Host "2. Rebuild profile on the server"
    Write-Host "3. Interactive Troubleshooter"
    Write-Host "4. Exit"
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
            Get-ChildItem -Path "C:\Users\$user\AppData\Roaming\Nuance\NaturallySpeaking12\Cache\$username*" | % { Write-Host $_.FullName }
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
    New-InteractiveTroubleshooter
}

if ($response -eq 4) {
    exit
}
