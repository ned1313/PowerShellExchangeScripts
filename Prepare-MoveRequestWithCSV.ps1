<#
    Prepare-MoveRequest from CSV file
    Must be run from the Scripts folder on the Exchange server
    Typically \Program Files\Microsoft\Exchange Server\V15\Scripts
#>
[CmdletBinding()]
param(
    [string] $CSVFile,
    [string] $RemoteDomainController,
    $RemoteCredential,
    [string] $LocalDomainController,
    $LocalCredential,
    [string] $TargetOU
)

$logFile = ".\Prepare-MoveRequestWithCSV_LogFile_$(Get-Date -f yyyy-MM-dd-hh-mm-ss).txt"

#Import the values from the CSV
$CSV = Import-Csv $CSVFile

#Run Prepare-MoveRequest with the values in the CSV
foreach($entry in $CSV){
    $user = $entry.SourceUserName
    try{
        Add-Content -Path $logFile -Value ("$(get-date -f s) Running Prepare-MoveRequest for user $user") -PassThru | `
            Write-Output | Write-Verbose
        .\Prepare-MoveRequest.ps1 -Identity $user `
            -RemoteForestDomainController $RemoteDomainController -RemoteForestCredential $RemoteCredential `
            -LocalForestDomainController $LocalDomainController -LocalForestCredential $LocalCredential `
            -TargetMailUserOU $TargetOU –UseLocalObject -OverwriteLocalObject
    }
    catch{
        Add-Content -Path $logFile -Value ("$(get-date -f s) Prepare-MoveRequest failed for user $user with error: $_") -PassThru | `
            Write-Output | Write-Host -ForegroundColor Red
    }
}