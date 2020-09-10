<#
    .Synopsis
    List all installed software and display to user and write to CSV file.

    

#>

# wmic qfe list

Function GetSoftware{
    $computername=$env:computername
    $Software = @()
    $UninstallKey="SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall" 
    $reg=[microsoft.win32.registrykey]::OpenRemoteBaseKey("LocalMachine",$computername) 
    $regkey=$reg.OpenSubKey($UninstallKey) 
    $subkeys=$regkey.GetSubKeyNames() 
    foreach($key in $subkeys){
        $thisKey=$UninstallKey+ "\\" +$key 
        $thisSubKey=$reg.OpenSubKey($thisKey) 
        $obj = New-Object PSObject
        $obj | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $computername
        $obj | Add-Member -MemberType NoteProperty -Name "Publisher" -Value $($thisSubKey.GetValue("Publisher"))
        $obj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $($thisSubKey.GetValue("DisplayName"))
        $obj | Add-Member -MemberType NoteProperty -Name "DisplayVersion" -Value $($thisSubKey.GetValue("DisplayVersion"))
        $obj | Add-Member -MemberType NoteProperty -Name "Size" -Value $($thisSubKey.GetValue("Size"))
        $obj | Add-Member -MemberType NoteProperty -Name "InstallDate" -Value $($thisSubKey.GetValue("InstallDate"))
        $Software += $obj
    } 
    # Sort by Publisher
    # $Software | Where-Object { $_.DisplayName } | Select-Object ComputerName, Publisher, DisplayName, DisplayVersion, Size, InstallDate 
    $Software | Sort-Object -Property Publisher |  Where-Object { $_.DisplayName } | Format-Table -AutoSize
    $Software | Sort-Object -Property Publisher |  Where-Object { $_.DisplayName } | Export-Csv -Path .\InstalledSoftware.csv -NoTypeInformation
}

GetSoftware
