<#
    .NOTES
    Script to update the extensionAttribute6 on source domain to the value Migration, and
    copy the msExchMailboxGUID value from the source user to the target user. Script is to
    be used in conjuction with the ContactMerge batch file, which has the header columns
    including SourceUserName and TargetUserName

    Written by Ned Bellavance

    .PARAMETER ContactMergeCSVFile
    String containing the path to the tab delimited migration batch file.

    .PARAMETER targetDomainController
    String containing the target domain controller to query and update user objects.

#>

[CmdletBinding()]

param(
    [Parameter(Mandatory=$true)]
    [string] $ContactMergeCSVFile,
    [Parameter(Mandatory=$true)]
    [string] $targetDomainController = ""
)

#Start log file
$logFile = ".\Set-ExtensionAttribute_LogFile_$(Get-Date -f yyyy-MM-dd-hh-mm-ss).txt"

$extValue = "Migration"

#Preferred Source DC list
$sourceDCList = @{}

#Import ActiveDirectory Module
Import-Module ActiveDirectory

#Get the matching user list imported
Add-Content -Path $logFile -Value ("$(get-date -f s) Loading content from $ContactMergeCSVFile") -PassThru | `
            Write-Output | Write-Verbose
if(Test-Path $ContactMergeCSVFile){
    $userList = Import-Csv $ContactMergeCSVFile
}
else{
    Add-Content -Path $logFile -Value ("$(get-date -f s) Invalid Path for UserMigrationCSVFile.  Path submitted is $ContactMergeCSVFile.  Script will now exit") -PassThru | `
            Write-Output | Write-Host -ForegroundColor Red
    exit
}


#Iterate through list
foreach($entry in $userList){
    #Get the first element splitting on the tab
    $sourceUser = $entry.SourceUserName
    $targetUser = $entry.TargetUserName
    $sourceDomain = $entry.SourceDomain
    $sourceDomainController = $sourceDCList.$sourceDomain

    try{
        #Use source name to update extensionAttribute6
        $ADSourceUser = Get-ADUser -Identity $sourceUser -Properties extensionAttribute6,msExchMailboxGuid -Server $sourceDomainController
        if($ADSourceUser.extensionAttribute6 -ne $null){
            Add-Content -Path $logFile -Value ("$(get-date -f s) $sourceUser has value $($ADSourceUser.extensionAttribute6). Clearing current value.") -PassThru | `
                Write-Output | Write-Verbose
            Set-ADUser -Identity $ADSourceUser -Clear extensionAttribute6 -Server $sourceDomainController
        }
        Set-ADUser -Identity $ADSourceUser -Add @{extensionAttribute6=$extValue} -Server $sourceDomainController
        Add-Content -Path $logFile -Value ("$(get-date -f s) $sourceUser attribute set to $extValue.") -PassThru | `
            Write-Output | Write-Verbose
    }
    catch{
        Add-Content -Path $logFile -Value ("$(get-date -f s) Error trying to update $sourceUser.  Error received: $_") -PassThru | `
            Write-Output | Write-Host -ForegroundColor Red
    }
    try{
        #Attempt to update target user with mailbox GUID
        [byte[]]$sourceExchGuid = $ADSourceUser.msExchMailboxGuid
        Add-Content -Path $logFile -Value ("$(get-date -f s) $sourceUser msExchMailboxGUID is set to $sourceExchGuid.") -PassThru | `
            Write-Output | Write-Verbose
        Add-Content -Path $logFile -Value ("$(get-date -f s) Trying msExchMailboxGUID update for $targetUser") -PassThru | `
            Write-Output | Write-Verbose
        Set-ADUser -Server $targetDomainController -Identity $targetUser -Add @{msExchMailboxGUID=$sourceExchGuid}
        Add-Content -Path $logFile -Value ("$(get-date -f s) Set msExchMailboxGUID on target account $targetUser to $sourceExchGuid") -PassThru | `
            Write-Output | Write-Verbose
    }
    catch{
        Add-Content -Path $logFile -Value ("$(get-date -f s) Error trying to update msExchMailboxGUID on $targetUser.  Error received: $_") -PassThru | `
            Write-Output | Write-Host -ForegroundColor Red
    }

}

Add-Content -Path $logFile -Value ("$(get-date -f s) Script complete") -PassThru | `
            Write-Output | Write-Verbose