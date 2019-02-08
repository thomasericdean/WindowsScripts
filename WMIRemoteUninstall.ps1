# Get all Windows Computers in an OU
$Computers =("192.168.0.1")

# For each computer
foreach($Computer in $Computers)
{
    # If the computer is on
    if (Test-Connection -ComputerName $Computer -Quiet -Count 1)
    {
        wmic /node:$Computer product where "name like'%%Google%%'" call uninstall
    }
}



