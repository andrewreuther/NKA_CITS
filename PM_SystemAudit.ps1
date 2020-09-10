<#
    Author: Rosalia Hernandez & Justine Woo

    Modified by: Andrew Reuther 9/4/2019
        -- Removed Aliases

    

#>



function findSerialNumber 
{
    try
    {
        return Get-WmiObject win32_bios | Select-Object -Expand serialnumber
    }
    catch { return $_.Exception.Message }
}
 
 
function getOS
{
    try
    {
        return Get-WmiObject Win32_OperatingSystem
    }
    catch
    {
        return ""
    }
}
 
function findInstallDate
{
    param ([WMI] $OS)
 
    try
    {
        if (-Not($OS)) { return "OS Not Found" }
        else { return ([WMI]'').ConvertToDateTime(($OS).InstallDate) }
    }
    catch { return $_.Exception.Message }
}
 
function findOSVersion
{
    param ([WMI] $OS)
 
    try
    {
        if (-Not ($OS)) {return "OS Not Found"}
        else
        {
            return $OS.version
        }
    }
    catch { return $_.Exception.Message }
}
 
function findLastBoot
{
    param ([WMI] $OS)
 
    try
    {
        return ([WMI]'').ConvertToDateTime(($OS).lastBootUpTime)
    }
    catch { return $_.Exception.Message }
}
 
function findServicePack
{
    param ([WMI] $OS)
 
    try 
    {
      
      return $OS.ServicePackMajorVersion
    }
    catch { return $_.Exception.Message }
}
 
function findIP
{
    param ([Array] $ipInfo)
 
    try
    {
        $ipList = ""
        $output = $ipInfo | where-object{($_ -like "*IPv4*")} | ForEach-Object{($_ = $_.split(":")[1]) -and ($ipList += $_ + ",")}
        write-host $output
        return $ipList
    }
    catch { return $_.Exception.Message }
}
 
function findHD
{
    try
    {
        foreach ($_ in Get-WmiObject Win32_LogicalDisk)
        {
            if ($_.DeviceID -eq "C:")
            {
                return @($_.FreeSpace, $_.Size)
                break
            }
        }
    }
    catch { return $_.Exception.Message }
}
 
function findMacAddress
{
 
    param ([Array] $ipInfo)
 
    $MacList = ""
    try
    {
        $count = 1
        while ($count -ne $ipInfo.Count)
        {
            if ($ipInfo[$count] -like "*Physical Address*")
            {
                $temp = $count - 1
                while ($True)
                {
                    if (-Not (($ipInfo[$temp] -like " *") -or ($ipInfo[$temp] -eq "")))
                    {
                        $mac = $ipInfo[$count].split(":")[1]
                        $MacList += "(" + $ipInfo[$temp] + ", " + $mac + ")" + "`n"
                        break
                    }
                    else {$temp--}
                }
            }
            $count++
       }
       return $MacList
    }
    catch { return $_.Exception.Message }
}
 
function findPerfMon
{
    param ([string]$HospitalName, [string] $HospitalState, [string] $HospitalCity, [string]$serialNum)
 
    try
    {
       
        $path = "C:\Users\nksupport\" + $HospitalName + " " + $HospitalState + " " + $HospitalCity + " " + $serialNum + " perfMon.blg"
 
        Get-counter -Counter "\LogicalDisk(_Total)\% Free Space", "\PhysicalDisk(_Total)\% Idle Time", "\PhysicalDisk(_Total)\Avg. Disk Sec/Read", "\PhysicalDisk(_Total)\Avg. Disk Sec/Write", "\Memory\Cache Bytes", "\Memory\% Committed Bytes In Use", "\Memory\Available Mbytes", "\Memory\Pool NonPaged Bytes", "\Memory\Pool Paged Bytes", "\Memory\Pages/Sec", "\Processor(_Total)\% User Time", "\System\Processor Queue Length","\Processor(_Total)\% Interrupt Time", "\Processor(_Total)\% Processor Time"  -SampleInterval 5 -MaxSamples 12 | export-counter -Path $path -Force
    }
    catch { return $_.Exception.Message }   
}
 
function findSystemModel
{
 
    try
    {
        $systemInfo = systemInfo
        $output = $systemInfo | where-Object{($_ -like "System Model:*")} | ForEach-Object{$_ -replace ":\s+"," = "}
        $systemModel = ConvertFrom-StringData($output | Out-String)
        return $systemModel."System Model"
    }
    catch { return $_.Exception.Message }
}
 
function findPatches
{
    param ([string] $HospitalName, [string] $HospitalState, [string] $HospitalCity, [string] $serialNum)
 
    try
    {
        $HotFixList = ""
        $session = New-Object -ComObject "Microsoft.Update.Session"
        $searcher = $session.CreateUpdateSearcher()
        $historyCount = $searcher.GetTotalHistoryCount()
        if ($historyCount -ne 0) 
        {
            foreach($_ in $searcher.QueryHistory(0, $historyCount)) {$HotFixList += $_.Title}
        }
        else
         {
            foreach ($_ in (Get-HotFix).HotFixID) {$HotFixList += $_ + ", "}
        }
 
        $content = @{
            serialNum = $serialNum
            HotFix = $HotFixList
        }
 
        $result += New-Object PSObject -Property $content
        $path =  "C:\Users\nksupport\" + $HospitalName + " " + $HospitalState + " " + $HospitalCity + " " + $serialNum + " HotFix.csv"
        $result |Select-Object-Object -Property "serialNum", "HotFix" | export-csv -Path $path -NoTypeInformation
    }
    catch { return $_.Exception.Message }
}
 
function findWireshark
{
    try
    {
        Set-Location "C:\"
        $path = (Get-ChildItem -Path C:\ -Recurse "wireshark.exe").fullname
        start-process $path  
    }
    catch { return $_.Exception.Message }
}
 
function findSoftware
{
    param ([string] $model)
 
    try
    {
        $softwareList = ""
        if ($model -eq "L")
        {
            $type = Read-Host -Prompt "What is the Model of the server? `n 1.HL7 `n 2.Pager `n 3.Old ECG `n 4.ECG `n 5.Aware `n 6.Net Konnect `n 7.ViTrac `n 8.Neuro `n"
            if ($type -eq 1) {$type = "HL7"}
            elseif ($type -eq 2) {$type = "Pager"}
            elseif ($type -eq 3) {$type = "Old ECG"}
            elseif ($type -eq 4) {$type = "ECG"}
            elseif ($type -eq 5) {$type = "Aware"}
            elseif ($type -eq 6) {$type = "NetKonnect"}
            elseif ($type -eq 7) {$type = "ViTrac"}
            elseif ($type -eq 8) {$type = "Neuro"}

            switch ($type){
                1 {}
                2 {}
                3 {}
                4 {}
                5 {}
                6 {}
                7 {}
                8 {}

            }





 
            if ($type -eq "Old ECG")
            {
                $path = (Get-ChildItem -Path C:\ -Recurse "VERSION" | Sort-Object -Property @{Expression = {$_.LastWriteTime}; Descending = $True} | Select-Object-Object -Index 0).fullname
                $string = ((Get-Content $path) | Out-String)
                $softwareList = @($type, $string.substring($string.length - 7, 7))
            }
            else
            {
                $pathHelp = $type + "gwy.exe"
                $path = (Get-ChildItem -Path C:\ -Recurse $pathHelp | Sort-Object -Property @{Expression = {$_.LastWriteTime}; Descending = $True} | Select-Object-Object -Index 0).Directory.fullname
 
                Set-Location $path
                $command = $type + "gwy -version"
                $version = [string](cmd /c $command)
                $version = $version.susbstring(0, $version.IndexOf(","))
                $softwareList =  @($type, $version)
                Set-Location "C:\Users\nksupport"
            }
        }
        else
        {
            Set-Location "C:\Program Files\Nihon Kohden\Unified Gateway"
            foreach ($_ in (Get-ChildItem).name)
            {
                if ($_ -like "*.exe")
                {
                    if ($_ -like "UnifiedGateway.exe")
                    {
                        start-process -FilePath "C:\Users\nksupport\AppData\Local\Chromium\Application\chrome.exe" -ArgumentList "localhost" 
                    }
                    else
                    {
                        start-process $_
                    }
                    $version = Read-Host -Prompt ("Enter the version of {0}" -f $_)
                    $softwareList += "(" + $_ + ", " + $version + ")" + "`n"
                }
            }
        } 
        return $softwareList
    }
    catch { return $_.Exception.Message }
}
 
function findSubnetMask
{
    param([Array] $ipInfo)
    try
    {
        $subnet = ""
        $output = $ipInfo | where-object{($_ -like "*Subnet Mask*")} | ForEach-Object{($_ = $_.split(":")[1]) -and ($subnet += $_)}
        return $subnet
    }
    catch { return $_.Exception.Message }
}
 
function findLicensedBeds
{
 
    param([string] $model)
 
    try
    {
        if ($model -eq "E")
        {
            start-process -FilePath "C:\Users\nksupport\AppData\Local\Chromium\Application\chrome.exe" -ArgumentList "localhost" 
            $LicensedBeds = Read-Host "Number of Beds Licensed"
            return $LicensedBeds
        }
    }
    catch {$_.Exception.Message}
}
 
function Main
{ 
    $HDSpace = findHD
    $HospitalName = Read-Host -Prompt "Enter Hospital Name"
    $HospitalState = Read-Host -Prompt "Enter Hospital State"
    $HospitalCity = Read-Host -Prom pt "Enter Hospital City"
    $model = Read-Host -Prompt "Is the server Legacy or Enterprise ('L' for Legacy or 'E' for Enterprise)?"
    write-host The following process will take up to 3 minutes
    $serialNum = findSerialNumber
    $ipInfo = ipconfig /all
    if ($model -eq "E")
    {
        # $serverType = "EUG"
        $softwareInfo = findSoftware($model)
    }
    else 
    {
        $softwareInfo = findSoftware($model)
        $softwareType = $softwareInfo[0]
        $softwareInfo = $softwareInfo[1]
    }
 
    $content = @{
        HospitalName = $HospitalName
        HospitalState = $HospitalState
        HospitalCity = $HospitalCity
        serialNum = $serialNum
        installDate = findInstallDate(getOS)
        OSVersion = findOSVersion(getOS)
        lastBoot = findLastBoot(getOS)
        servicePack = findServicePack(getOS)
        SystemModel = findSystemModel
        SubnetMask = findSubnetMask($ipInfo)
        IP = findIP($ipInfo)
        MacAddress = findMacAddress($ipInfo)
        HDFree = [string]([math]::Round($HDSpace[0]/1gb, 2)) + "GB"
        HDSize = [string]([math]::Round($HDSpace[1]/1gb, 2)) + "GB"
        HDPercentage = [string]([math]::Round(($HDSpace[0]/1gb)/($HDSpace[1]/1gb) * 100, 2)) + "%"
        PerfMon = findPerfMon -HospitalName $HospitalName -HospitalState $HospitalState -HospitalCity $HospitalCity -serialNum $serialNum
        Patches = findPatches -HospitalName $HospitalName -HospitalState $HospitalState -HospitalCity $HospitalCity -serialNum $serialNum
        SoftwareType = $softwareType
        SoftwareVersions = $softwareInfo
        LicensedBeds = findLicensedBeds($model)
    }
 
    $result += New-Object PSObject -Property $content
 
   $path = ".\"+ $HospitalName + " " + $HospitalState + " " + $HospitalCity + " " + $serialNum + ".csv"
   
   $result | Select-Object-Object -Property "HospitalName", "HospitalState", "HospitalCity", "serialNum", "SystemModel", "OSVersion", "lastBoot", "servicePack", "installDate", "SubnetMask", "HDFree", "HDSize", "HDPercentage", "SoftwareType", "SoftwareVersions", "LicensedBeds", "IP", "MacAddress" | export-csv -Path $path -NoTypeInformation
 
   findWireshark
}

Main
