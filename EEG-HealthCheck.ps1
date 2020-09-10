<#
 .Synopsis
  EEG HealthCheck

 .Description
  Blah 

 .Parameter 
  # TBD.

 .Example
   # TBD

   
 .Outputs 
 
 
#>

# Vars
$reportName = "C:\Software\NK Audit\permissions-$(Get-Date -format MMyy).csv"
If (Test-Path  $reportName){
    Remove-Item -Path $reportName
}

$filePaths=(
    "c:\DVC",
    "C:\NFX",
    "C:\NFX11",
    "C:\NKT",
    "C:\Program Files\nkstart",
    "C:\Program Files (x86)\Nihon Kohden",
    "D:\NKT"
)

$Keys=(
    "HKLM:\SOFTWARE\WOW6432Node\Nihon Kohden",
    "HKLM:\SOFTWARE\WOW6432Node\NK Neuro"
)


function Get-Perms($Path){
    $obj = 

}

Set-Location "C:\Software\NK Audit"
Foreach ($filePath in $filePaths){
    If (test-path $filePath){
        $testPaths = Get-ChildItem -Path $filePath -Recurse -Force
        $Results = @()
        Foreach ($testPath in $testPaths) {    
            $Acl = Get-Acl -Path $testPath.FullName    
            foreach ($Access in $acl.Access) {       
                if ($Access.IdentityReference -notlike "BUILTIN\Administrators" -and $Access.IdentityReference -notlike "domain\Domain Admins" -and $Access.IdentityReference -notlike "CREATOR OWNER" -and $access.IdentityReference -notlike "NT AUTHORITY\SYSTEM") {            
                    $Properties = [ordered]@{'FolderName'=$Folder.FullName;
                    'AD Group'=$Access.IdentityReference;'Permissions'=$Access.FileSystemRights;'Inherited'=$Access.IsInherited}            
                    $Results += New-Object -TypeName PSObject -Property $Properties        
                }    
            }
        }
    }Else{
        $_ | Write-Host
    }
}

$Results | Export-Csv -path $reportName
    