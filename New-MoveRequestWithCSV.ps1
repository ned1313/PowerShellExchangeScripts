<#
    New-MoveRequest with CSV
#>

[CmdletBinding()]
param(
    [string] $CSVFile,
    [string] $RemoteHost,
    [string] $RemoteGlobalCatalog,
    $RemoteCredential,
    [string] $LocalDomainController,
    $LocalCredential,
    [string] $TargetDatabase,
    [string] $TargetDeliveryDomain
)

$logFile = ".\New-MoveRequestWithCSV_LogFile_$(Get-Date -f yyyy-MM-dd-hh-mm-ss).txt"

#Import the values from the CSV
$CSV = Import-Csv $CSVFile

foreach($entry in $CSV){

    $user = $entry.SourcePrimaryEmail
    try{
        Add-Content -Path $logFile -Value ("$(get-date -f s) Running New-MoveRequest for user $user") -PassThru | `
            Write-Output | Write-Verbose
        New-MoveRequest -Remote -Identity $user -RemoteCredential $RemoteCredential -RemoteGlobalCatalog $RemoteGlobalCatalog `
            -RemoteHostName $RemoteHost -TargetDeliveryDomain $TargetDeliveryDomain -TargetDatabase $TargetDatabase `
            -AllowLargeItems:$true -SuspendWhenReadyToComplete -BadItemLimit 1000 -AcceptLargeDataLoss
    }
    catch{
        Add-Content -Path $logFile -Value ("$(get-date -f s) New-MoveRequest failed for user $user with error: $_") -PassThru | `
            Write-Output | Write-Host -ForegroundColor Red
    }

}