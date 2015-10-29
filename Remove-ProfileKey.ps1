<#
    .NOTES
    Script requires a CSV file with the headers of the ComputerName, target samAccount name as TargetUserName, and domain
    name for the computer as SourceDomain. The script also requires the presence of the ActiveDirectory
    PowerShell module in order to query the SID of the users.

    Written by Ned Bellavance

    .PARAMETERS ProfileRemovalCSV
    Required parameter of the path to the CSV file with the ComputerName, TargetUserName, and SourceDomain.

    .PARAMETERS targetDomainController
    Optional parameter of the targe domain controller to query SID from


#>
[CmdletBinding()]

param(
    [Parameter(Mandatory=$true)]
    [string] $ProfileRemovalCSV,
    [Parameter(Mandatory=$false)]
    [string] $targetDomainController=""
)

Import-Module ActiveDirectory
#Create Log file for the run
$logFile = ".\Remove-ProfileKey_LogFile_$(Get-Date -f yyyy-MM-dd-hh-mm-ss).txt"

Add-Content -Path $logFile -Value "$(get-date -f s) Log file started" -PassThru | `
            Write-Output | Write-Verbose

Function RemoveProfileKey{
[CmdletBinding()]

    param(
        [Parameter(Mandatory=$true)]
        [string] $computerName,
        [Parameter(Mandatory=$true)]
        [string] $DomainName,
        [Parameter(Mandatory=$true)]
        [string] $userName,
        [Parameter(Mandatory=$true)]
        [string] $targetDC
    )

    #Convert computer name to FQDN
    $computerName = "$computerName.$DomainName"

    Add-Content -Path $logFile -Value "$(get-date -f s) Computername set to $computerName" -PassThru | `
            Write-Output | Write-Verbose
    
    #Verify computer is pingable
    if(Test-Connection -ComputerName $computerName -Count 2 -Quiet){
        
        #Check to see if the RemoteRegistry service is started
        if((Get-Service -ComputerName $computerName RemoteRegistry).Status -eq "Stopped"){
            Add-Content -Path $logFile -Value "$(get-date -f s) Remote Registry on $computerName is not started" -PassThru | `
                Write-Output | Write-Verbose
            try{
                #Try to start the service
                Get-Service -ComputerName $computerName RemoteRegistry | Start-Service
                Add-Content -Path $logFile -Value "$(get-date -f s) Started Remote Registry on $computerName" -PassThru | `
                    Write-Output | Write-Verbose

            }
            catch{
                Add-Content -Path $logFile -Value "$(get-date -f s) Starting Remote Registry on $computerName failed: $_" -PassThru | `
                    Write-Output | Write-Host -ForegroundColor Red
                break
            }
        }

        #Try to get the sid for the username
        try{
            $sid = (Get-ADUser -Server $targetDC -Identity $userName).sid
            Add-Content -Path $logFile -Value "$(get-date -f s) SID for $userName found in Active Directory: $sid" -PassThru | `
                Write-Output | Write-Verbose
        }
        catch{
            Add-Content -Path $logFile -Value "$(get-date -f s) SID not found for $userName with error: $_" -PassThru | `
                Write-Output | Write-Host -ForegroundColor Red
            break
        }
        try{
            #Try to get the registry key associated with the user's sid
            $reg = reg query "\\$computerName\HKLM\Software\Microsoft\Windows NT\CurrentVersion\ProfileList" /f $sid /k
            Add-Content -Path $logFile -Value "$(get-date -f s) Registry entries queried successfully: $reg" -PassThru | `
                Write-Output | Write-Verbose
        }
        catch{
            Add-Content -Path $logFile -Value "$(get-date -f s) Reg query command failed with error: $_" -PassThru | `
                Write-Output | Write-Host -ForegroundColor Red
            break
        }

        #If any matches were found
        if($reg[($reg.count)-1] -notlike ("*0 match*")){
            
            for($i=1;$i -lt ($reg.count)-1;$i++){
                try{
                    #Get the profile path for the registry key
                    $pathValue = reg query "\\$computerName\HKLM\Software\Microsoft\Windows NT\CurrentVersion\ProfileList\$sid" /v ProfileImagePath
                    $pathValue = $pathValue[2] -split '\s+' -match '\S'
                    $folder = $pathValue[2].split("\")
                    $folder = $folder | select -Last 1

                    #Try to delete the registry key
                    Add-Content -Path $logFile -Value "$(get-date -f s) Trying to remove registry entry $($reg[$i])" -PassThru | `
                        Write-Output | Write-Verbose
			        reg delete "\\$computerName\HKLM\Software\Microsoft\Windows NT\CurrentVersion\ProfileList\$sid" /f
                    Add-Content -Path $logFile -Value "$(get-date -f s) Successfully removed entry $($reg[$i])" -PassThru | `
                        Write-Output | Write-Verbose
                }
                catch{
                    Add-Content -Path $logFile -Value "$(get-date -f s) Reg delete command failed with error: $_" -PassThru | `
                        Write-Output | Write-Host -ForegroundColor Red
                }
            }

            #Try to get the User folders associated with the profile
            try{
                Add-Content -Path $logFile -Value "$(get-date -f s) Trying to get user folder on $computerName" -PassThru | `
                    Write-Output | Write-Verbose
                $userFolder = Get-Item "\\$computername\c$\Users\$folder"
                #Rename the folder to append .old
                Add-Content -Path $logFile -Value "$(get-date -f s) Renaming folder $($userFolder.Name)" -PassThru | `
                    Write-Output | Write-Verbose
                Rename-Item $userFolder.FullName "$($userFolder.Name).old"

            }
            catch{
                Add-Content -Path $logFile -Value "$(get-date -f s) Retrieval of user folders failed on $computerName with error: $_" -PassThru | `
                    Write-Output | Write-Host -ForegroundColor Red
            }

        }
        else{
            #If there is no registry entry, log it
            Add-Content -path $logFile -Value "$(get-date -f s) SID $sid not found in registry for computer $computerName" -PassThru | `
                Write-Output | Write-Verbose
        }
    }
    else{
        Add-Content -Path $logFile -Value "$(get-date -f s) Computer $computerName not available" -PassThru | `
            Write-Output | Write-Host -ForegroundColor Red
    }

}

#Import the entries from the CSV file
$entries = Import-Csv -Path $ProfileRemovalCSV

foreach($entry in $entries){
    RemoveProfileKey -computerName $entry.ComputerName -DomainName $entry.SourceDomain -userName $entry.TargetUserName -targetDC $targetDomainController
}

Add-Content -Path $logFile -Value "$(get-date -f s) Script complete" -PassThru | `
        Write-Output | Write-Verbose