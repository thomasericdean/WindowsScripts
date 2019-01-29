foreach($Computer in $Computers)
{
    # If the computer is on
    if (Test-Connection -ComputerName $Computer -Quiet -Count 1)
    {
        if(PsExec64.exe \\$Computer net localgroup Administrators $GroupName /add){
            Write-Output "Added $GroupName to local admin on $Computer"
        }
    }
}
