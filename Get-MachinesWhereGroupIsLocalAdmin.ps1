$Computers = Get-ADComputer -Filter { OperatingSystem -like '*Windows*' }
function get-localadmin {  
    param ($strcomputer)  
    
    $admins = Gwmi win32_groupuser –computer $strcomputer   
    $admins = $admins |? {$_.groupcomponent –like '*"Administrators"'}  
    
    $admins |% {  
        $_.partcomponent –match “.+Domain\=(.+)\,Name\=(.+)$” > $nul
        if( ($matches[1].trim('"') + “\” + $matches[2].trim('"')) -eq "GESA\Domain Admins"){
            Write-Output "Admin on $strcomputer"
        }

    }  
}

foreach($computer in $Computers)
{
    get-localadmin($computer)
}
