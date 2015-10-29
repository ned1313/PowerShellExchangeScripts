<#
    .NOTES
    Script requires a CSV file with the headers of the contact object mailNickName (alias) and the mailbox samAcountName.
    The headers should be TargetContactAlias and TargetUserName.  The script also requires the presence of the ActiveDirectory
    PowerShell module in order to update groups.

    Written by Ned Bellavance

    .PARAMETERS csvFile
    Required parameters of the path to the CSV file with the contact and mailbox user information with headers TargetContactAlias
    and TargetUserName.

    .PARAMETERS exchangeServer
    Optional parameter of the FQDN of the Exchange server establish a remote PSSession

    .PARAMETERS domainController
    The preferred domain controller for all Exchange and AD commands.

#>

[CmdletBinding()]

param(
    [string] $csvFile,
    [string] $exchangeServer="",
    [string] $domainController = ""
)

#Create Log file for the run
$logFile = ".\Merge-ContactsAndMailboxes_LogFile_$(Get-Date -f yyyy-MM-dd-hh-mm-ss).txt"

#Create the object output file
$objFile = ".\Merge-ContactsAndMailboxes_ObjectData_$(Get-Date -f yyyy-MM-dd-hh-mm-ss).txt"

#Create the address output file path
$addressFile = ".\$((Get-Item $pwd).Name)_AddressFile.csv"

#If the file doesn't exist give it some headers
if(-not (Test-Path $addressFile)){
    Add-Content -Path $addressFile -Value "samAccountName;EmailAddress"
}


#Check if ActiveDirectory module is loaded
if(-not (Get-Module -Name ActiveDirectory)){
    if(Get-Module -ListAvailable | ?{$_.Name -eq "ActiveDirectory"}){
        Add-Content $logFile -Value "$(get-date -f s) Loading Active Directory Module" -PassThru | `
            Write-Output | Write-Verbose
        Import-Module ActiveDirectory
    }
    else{
        Add-Content $logFile -Value "$(get-date -f s) ActiveDirectory module not available.  This is required for script" -PassThru | `
            Write-Output | Write-Host -ForegroundColor Red
        break
    }
}

#Try to create a remote PSSession with Exchange server
try{
    Add-Content $logFile -Value "$(get-date -f s) Establishing Exchange PSSession with server $exchangeServer" -PassThru | `
        Write-Output | Write-Verbose
    $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$exchangeServer/powershell"
    Import-PSSession $session
    Set-ADServerSettings -PreferredServer $domainController
}
catch{
    Add-Content $logFile -Value "$(get-date -f s) Could not establish remote PSSession with Exchange server $exchangeServer with error: $_" -PassThru | `
        Write-Output | Write-Host -ForegroundColor Red
    break
}

Function MergeAccounts {
    [CmdletBinding()]

    param(
        [string] $contactName,
        [string] $mailboxUserName
    )
    #Store Contact information
    try{
        Add-Content $logFile -Value "$(get-date -f s) Attempting to get contact $contactName" -PassThru | `
            Write-Output | Write-Verbose
        $contact = Get-MailContact $contactName
        WriteObjectInfo -sourceObj $contact -outFile $objFile
    }
    catch{
        Add-Content $logFile -Value "$(get-date -f s) Could not find contact object $contactName" -PassThru | `
            Write-Output | Write-Host -ForegroundColor Red
        throw "Get-MailContact failed for $contactName with Error: $_"
        break
    }

    #Store user mailbox information
    try{
        Add-Content $logFile -Value "$(get-date -f s) Attempting to get mailbox $mailboxUserName" -PassThru | `
            Write-Output | Write-Verbose
        $mailboxuser = Get-Mailbox $mailboxUserName
        WriteObjectInfo -sourceObj $mailboxuser -outFile $objFile
    }
    catch{
        Add-Content $logFile -Value "$(get-date -f s) Could not find mailbox object $mailboxUserName" -PassThru | `
            Write-Output | Write-Host -ForegroundColor Red
        throw "Get-Mailbox failed for $mailboxUserName with Error: $_"
        break
    }

    #Get-Contact groups if there are any
    try{
        Add-Content $logFile -Value "$(get-date -f s) Attempting to get Contact AD object for group membership" -PassThru | `
            Write-Output | Write-Verbose
        $contactAD = Get-ADObject -Server $domainController -Filter {(ObjectClass -eq "Contact") -and (mailNickName -eq $contactName)} -Properties *
    } 
    catch{
        Add-Content $logFile -Value "$(get-date -f s) Could not find AD contact object $contactName" -PassThru | `
            Write-Output | Write-Host -ForegroundColor Red
        throw "Get-ADObject failed for $contactName with Error: $_"
        break
    }
    try{
        #Remove the mailbox from the user
        Add-Content $logFile -Value "$(get-date -f s) Removing Mailbox from $mailboxUserName" -PassThru | `
            Write-Output | Write-Verbose
        Disable-Mailbox -Identity $mailboxUserName -Confirm:$false
        while(Get-Mailbox $mailboxUserName -WarningAction SilentlyContinue  -ErrorAction SilentlyContinue){
            Write-Verbose "$mailboxUserName not removed yet"
        }
    }
    catch{
        Add-Content $logFile -Value "$(get-date -f s) Removing mailbox failed for $mailboxUserName" -PassThru | `
        Write-Output | Write-Host -ForegroundColor Red
        throw "Disable-Mailbox failed for $mailboxUserName with Error: $_"
    }

    try{
        #Remove the contact object
        Add-Content $logFile -Value "$(get-date -f s) Removing contact object $contactName" -PassThru | `
            Write-Output | Write-Verbose
        Remove-ADObject $contactAD -Server $domainController -Confirm:$false
        while(Get-MailContact $contactName -WarningAction SilentlyContinue -ErrorAction SilentlyContinue){
            Write-Verbose "$contactName not removed yet"
        }
    }
    catch{
        Add-Content $logFile -Value "$(get-date -f s) Removing contact failed for $contactName" -PassThru | `
        Write-Output | Write-Host -ForegroundColor Red
        throw "Remove-ADObject failed for $contactName with Error: $_"
    }

    #Mail enable the user
    try{
        Add-Content $logFile -Value "$(get-date -f s) Mail enabling user $mailboxUserName" -PassThru | `
            Write-Output | Write-Verbose
        $mailUser = Enable-MailUser $mailboxUserName -Alias $mailboxuser.Alias -PrimarySmtpAddress $mailboxuser.PrimarySmtpAddress -ExternalEmailAddress $contact.ExternalEmailAddress
    }
    catch{
        Add-Content $logFile -Value "$(get-date -f s) Mail enabling user failed for $mailboxUserName" -PassThru | `
        Write-Output | Write-Host -ForegroundColor Red
        throw "Enable-MailUser failed for $mailboxUserName with Error: $_"
    }

    #Collect the proxyAddresses from the contact and mailbox
    $proxyAddresses = @()
    $proxyAddresses += $contact.EmailAddresses
    $proxyAddresses += $mailboxuser.EmailAddresses
    foreach($address in $proxyAddresses){
        if(-not ($mailuser.EmailAddresses -contains $address)){
            $mailuser.EmailAddresses += $address
        }
    }

    #Add the legacyExchangeDNs as x500 addresses
    $x5001 = "X500:$($contactAD.LegacyExchangeDN)"
    Add-Content $logFile -Value "$(get-date -f s) x500 Address from contact is $x5001" -PassThru | `
            Write-Output | Write-Verbose
    $x5002 = "X500:$($mailboxuser.LegacyExchangeDN)"
    Add-Content $logFile -Value "$(get-date -f s) x500 Address from mailbox is $x5002" -PassThru | `
            Write-Output | Write-Verbose

    #Add the targetdomain.com address
    $targetAddress = "smtp:$($mailboxuser.alias)@targetdomain.com"

    #Update the email addresses of the mail enabled user
    try{
        Add-Content $logFile -Value "$(get-date -f s) Updating the email addresses on $mailboxUserName" -PassThru | `
            Write-Output | Write-Verbose
        set-mailuser $mailboxUserName -EmailAddresses $mailUser.EmailAddresses
        Add-Content $logFile -Value "$(get-date -f s) Adding the x500 addresses" -PassThru | `
            Write-Output | Write-Verbose
        Set-MailUser $mailboxUserName -EmailAddresses @{Add=$x5001}
        Set-MailUser $mailboxUserName -EmailAddresses @{Add=$x5002}
        Set-MailUser $mailboxUserName -EmailAddresses @{Add=$targetAddress}
        Add-Content $logFile -Value "$(get-date -f s) Resetting the primary SMTP address on $mailboxUserName" -PassThru | `
            Write-Output | Write-Verbose
        set-mailuser $mailboxUserName -PrimarySmtpAddress $mailboxuser.PrimarySmtpAddress -EmailAddressPolicyEnabled:$false -HiddenFromAddressListsEnabled:$false
        Add-Content $logFile -Value "$(get-date -f s) Writing all Email Addresses to the Address File: $addressFile" -PassThru | `
            Write-Output | Write-Verbose
        $addresses =  (Get-MailUser $mailboxUserName).EmailAddresses
        $addresses | %{Add-Content -Path $addressFile -Value "$mailboxUserName;$_"}
    }
    catch{
        Add-Content $logFile -Value "$(get-date -f s) Update of email addresses failed for $mailboxUserName" -PassThru | `
            Write-Output | Write-Host -ForegroundColor Red
        throw "Set-MailUser failed on $mailboxUserName with Error: $_"
    }
    #Update the group membership of the mail enabled user
    try{
        if($contactAD.memberOf.count -gt 0){
            Add-Content $logFile -Value "$(get-date -f s) Groups found for contact $contactName, adding $mailboxUserName to groups" -PassThru | `
                Write-Output | Write-Verbose
            foreach($group in $contactAD.memberOf){
                Add-Content $logFile -Value "$(get-date -f s) Adding $mailboxUserName to $group" -PassThru | `
                    Write-Output | Write-Verbose
                Get-ADGroup $group -Server $domainController | Add-ADGroupMember -Members $mailboxUserName -Server $domainController
            }
        }
    }
    catch{
        Add-Content $logFile -Value "$(get-date -f s) Failed to add $mailboxUserName to one or more groups in $($contactAD.memberOf)" -PassThru | `
            Write-Output | Write-Host -ForegroundColor Red
        throw "Get-ADGroup or Add-ADGroupMember failer with on $mailboxUserName with Error: $_"
    }
}

Function WriteObjectInfo{
    [cmdletbinding()]

    param(
        $sourceObj,
        $outFile
    )
    Add-Content -Path $outFile -Value "Name:$($sourceObj.Name)"
    Add-Content -Path $outFile -Value "Alias:$($sourceObj.Alias)"
    Add-Content -Path $outFile -Value "DisplayName:$($sourceObj.DisplayName)"
    Add-Content -Path $outFile -Value "DistinguishedName:$($sourceObj.DistinguishedName)"
    Add-Content -Path $outFile -Value "Identity:$($sourceObj.Identity)"
    Add-Content -Path $outFile -Value "LegacyExchangeDN:$($sourceObj.LegacyExchangeDN)"
    if($sourceObj.RecipientType -eq "UserMailbox"){
      Add-Content -Path $outFile -Value "samAccountName:$($sourceObj.samAccountName)"  
    }
    Add-Content -Path $outFile -Value "PrimarySmtpAddress:$($sourceObj.PrimarySmtpAddress)"
    Add-Content -Path $outFile -Value "ExternalEmailAddress:$($sourceObj.ExternalEmailAddress)"
    foreach($address in $sourceObj.EmailAddresses){
        Add-Content -Path $outFile -Value "EmailAddress:$address"
    }
    Add-Content -Path $outFile -Value "`n"
}

#Import the CSV file
$csvFileImport = Import-Csv -Path $csvFile

#Run through the entries in the CSV file
foreach($entry in $csvFileImport){
    #Validate Input
    if($entry.TargetUserName.Length -eq 0 -or $entry.TargetContactAlias.Length -eq 0 -or $entry.SourceDomain.Length -eq 0){
        Add-Content $logFile -Value "$(get-date -f s) One of the entries is empty, skipping the line, please check the CSV" -PassThru | `
            Write-Output | Write-Host -ForegroundColor Yellow
        continue
    }
    try{
        Add-Content $logFile -Value "$(get-date -f s) Merge started for $($entry.TargetUserName)" -PassThru | `
            Write-Output | Write-Verbose
        MergeAccounts -contactName $entry.TargetContactAlias -mailboxUserName $entry.TargetUserName
    }
    catch{
        Add-Content $logFile -Value "$(get-date -f s) Merge failed for $($entry.TargetUserName)" -PassThru | `
        Write-Output | Write-Host -ForegroundColor Red
        Add-Content $logFile -Value "$(get-date -f s) Failed error was: $_" -PassThru | `
        Write-Output | Write-Host -ForegroundColor Red
    }
}

#Always cleanup your sessions!
Remove-PSSession $session
