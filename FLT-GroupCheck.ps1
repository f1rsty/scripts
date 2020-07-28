# Checks to see if a list of computers are not in a group

$Group = "FLT-NightlyRebootForcedLogon"

$GroupObject = Get-ADGroup $Group | Select -ExpandProperty distinguishedName

$Results = ForEach ($Computer in $Computers)
{   Try {
        $ComputerObject = Get-ADComputer $Computer -Properties MemberOf,Created -ErrorAction Stop
    }
    Catch {
        Write-Warning "Unable to locate $Computer"
        Continue
    }
    If ($ComputerObject.MemberOf -notcontains $GroupObject)
    {   
        $ComputerObject | Select Name,Enabled,Created
    }
}

$Results | Export-Csv C:\Files\Computers_Not_in_Group_$Group.csv -NoTypeInformation
