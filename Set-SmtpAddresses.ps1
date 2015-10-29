<#
    .NOTES
    Script to fix the SMTP address entries for the mail user that is being migrated
    by the Dell Migration Manager tool.  The script takes a CSV file from the migration
    process containing the TargetContactAlias and TargetUserName fields.  The alias in the 
    TargetContactAlias will be added as the ExternalEmailAddress with @source.domain.com.  
    The current ExternalEmailAddress will be set as the primary SMTP address.

    Written by Ned Bellavance

    .PARAMETER CSVFile
    String containing the path to the tab delimited migration batch file.

    .PARAMETER targetDomainController
    String containing the target source controller to query and update user objects.

    .PARAMETERS exchangeServer
    Optional parameter of the FQDN of the Exchange server establish a remote PSSession

#>

[CmdletBinding()]

param(
    [Parameter(Mandatory=$true)]
    [string] $csvFile,
    [Parameter(Mandatory=$false)]
    [string] $exchangeServer="",
    [Parameter(Mandatory=$false)]
    [string] $targetDomainController=""
)

#Start log file
$logFile = ".\Set-SmtpAddresses_LogFile_$(Get-Date -f yyyy-MM-dd-hh-mm-ss).txt"

#Build path for address file and check for presence
$addressFile = ".\$((Get-Item $pwd).Name)_AddressFile.csv"
if(-not (Test-Path $addressFile)){
    Add-Content $logFile -Value "$(get-date -f s) The Address File $addressFile is not in the current folder, exiting script" -PassThru | `
            Write-Output | Write-Host -ForegroundColor Red
    Exit
}

#Try to create a remote PSSession with Exchange server
try{
    Add-Content $logFile -Value "$(get-date -f s) Establishing Exchange PSSession with server $exchangeServer" -PassThru | `
            Write-Output | Write-Verbose
    $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$exchangeServer/powershell"
    Import-PSSession $session
    Set-ADServerSettings -PreferredServer $targetDomainController
}
catch{
    Add-Content $logFile -Value "$(get-date -f s) Could not establish remote PSSession with Exchange server $exchangeServer with error: $_" -PassThru | `
            Write-Output | Write-Host -ForegroundColor Red
    break
}

Function SetSmtpAddresses {
    
    [CmdletBinding()]

    param(
        [string] $contactName,
        [string] $mailboxUserName,
        $addressList
    )

    #try to get the mailuser
    try{
        Add-Content $logFile -Value "$(get-date -f s) Attempting to get mailuser $mailboxUserName" -PassThru | `
            Write-Output | Write-Verbose
        $meu = Get-MailUser -identity $mailboxUserName
    }
    catch{
        Add-Content $logFile -Value "$(get-date -f s) Could not find mailbox object $mailboxUserName" -PassThru | `
            Write-Output | Write-Host -ForegroundColor Red
        throw "Get-MailUser failed for $mailboxUserName with Error: $_"
        break
    }

    #Log current values
    Add-Content $logFile -Value "$(get-date -f s) Current ExternalEMailAddress: $($meu.ExternalEmailAddress)" -PassThru | `
            Write-Output | Write-Verbose
    Add-Content $logFile -Value "$(get-date -f s) Current PrimarySmtpAddress: $($meu.PrimarySmtpAddress)" -PassThru | `
            Write-Output | Write-Verbose
    
    #store the ExternlMailAddress in a value
    $newPrimarySMTP = ($meu.ExternalEmailAddress).substring(5)

    #create the source mail address
    $newExternalAddress =  "$contactName@source.domain.com"

    try{
        if($addressList -ne $null){
            Add-Content $logFile -Value "$(get-date -f s) Addresses found in address file for $mailboxUserName" -PassThru | `
                Write-Output | Write-Verbose
            #Get current list of proxy addresses
            $currentAddresses = (Get-MailUser -Identity $mailboxUserName).EmailAddresses
            #Foreach address in list, is it isn't in proxy addresses add it
            foreach($address in $addressList){
                if($currentAddresses -notcontains $address.EmailAddress){
                    try{
                        Add-Content $logFile -Value "$(get-date -f s) Adding address $($address.EmailAddress) to $mailboxUserName" -PassThru | `
                            Write-Output | Write-Verbose
                        Set-MailUser -Identity $mailboxUserName -EmailAddresses @{Add=$($address.EmailAddress)}
                    }
                    catch{
                        Add-Content $logFile -Value "$(get-date -f s) Failed when adding address $($address.EmailAddress) to $mailboxUserName" -PassThru | `
                            Write-Output | Write-Host -ForegroundColor Red
                        continue
                    }
                }
            }

        }

        if($meu.ExternalEmailAddress -like "*@source.domain.com"){
            Add-Content $logFile -Value "$(get-date -f s) External Address already set to source.domain.com, no changes will be made" -PassThru | `
                Write-Output | Write-Verbose
        }
        else{
            #Set the external mail address
            Add-Content $logFile -Value "$(get-date -f s) Attempting to set ExternalEMailAddress to $newExternalAddress" -PassThru | `
                Write-Output | Write-Verbose
            Set-MailUser -Identity $mailboxUserName -ExternalEMailAddress $newExternalAddress
        }
        if($meu.PrimarySmtpAddress -notlike "*@targetdomain.com"){
            Add-Content $logFile -Value "$(get-date -f s) Primary SMTP Address for $mailboxUserName already set to $($meu.PrimarySmtpAddress), no changes will be made" -PassThru | `
                Write-Output | Write-Verbose
        }
        else{
            #Set the primary mail address from the old ExternalEMailAddress value
            Add-Content $logFile -Value "$(get-date -f s) Attempting to set PrimarySmtpAddress to $newPrimarySMTP" -PassThru | `
                Write-Output | Write-Verbose
            Set-MailUser -Identity $mailboxUserName -PrimarySmtpAddress $newPrimarySMTP
            #Log new values
            Add-Content $logFile -Value "$(get-date -f s) New ExternalEMailAddress: $newExternalAddress" -PassThru | `
                Write-Output | Write-Verbose
            Add-Content $logFile -Value "$(get-date -f s) New PrimarySmtpAddress: $newPrimarySMTP" -PassThru | `
                Write-Output | Write-Verbose
        }

    }
    catch{
     Add-Content $logFile -Value "$(get-date -f s) Could not update the external address or primary address for $mailboxUserName" -PassThru | `
            Write-Output | Write-Host -ForegroundColor Red
        throw "Set-MailUser failed for $mailboxUserName with Error: $_"
        break
    }

}

#Import the CSV file
$csvFileImport = Import-Csv -Path $csvFile
$addressImport = Import-Csv -Path $addressFile -Delimiter ';'

#Run through the entries in the CSV file
foreach($entry in $csvFileImport){
    try{
        Add-Content $logFile -Value "$(get-date -f s) SetSmtpAddresses started for $($entry.TargetUserName)" -PassThru | `
            Write-Output | Write-Verbose
        $addresses = $addressImport | ?{$_.samAccountName -eq $entry.TargetUserName}
        SetSmtpAddresses -contactName $entry.TargetContactAlias -mailboxUserName $entry.TargetUserName -addressList $addresses
    }
    catch{
        Add-Content $logFile -Value "$(get-date -f s) SetSmtpAddresses failed for $($entry.TargetUserName)" -PassThru | `
            Write-Output | Write-Host -ForegroundColor Red
        Add-Content $logFile -Value "$(get-date -f s) Failed error was: $_" -PassThru | `
            Write-Output | Write-Host -ForegroundColor Red
    }
}

#Always cleanup your sessions!
Remove-PSSession $session

Add-Content $logFile -Value "$(get-date -f s) Script Complete!" -PassThru | `
        Write-Output | Write-Verbose