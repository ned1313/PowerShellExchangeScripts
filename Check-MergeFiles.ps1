<#
.NOTES
This script is meant to validate the values in a SourceCSVFile and create the proper
merge files if there are no errors.  The script must be run with either the CheckUsers
or CheckWorkstations switch.  The SourceCSVFile should have as a minimum the fields 
    -Source User ID
    -Computer Name
    -Current Domain
The file is expected to have an extra header line, which the script will bypass.  The
script should be run from the migration directory. It uses the directory as the basis 
for naming the migration files that are produced.

Written By Ned Bellavance

.PARAMTER SourceCSVFile
The location of the CSV File conataining the entries to be processed

.PARAMETER targetDC
Optional parameter for specifying the target Domain Controller 

.PARAMETER targetExchangeServer
Option parameter for specifying the target Exchange Server

.PARAMETER CheckUsers
Switch parameter to run the user migration related checks

.PARAMETER CheckWorkstations
Switch parameter to run the workstation migration related checks



#>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string] $SourceCSVFile,
    [string] $targetDC = "",
    [string] $targetExchangeServer = "",
    [switch] $CheckUsers,
    [switch] $CheckWorkstations
)

if(-not ($CheckUsers -or $CheckWorkstations)){
    Write-Host "You must select at least one of CheckUsers or CheckWorkstations"
    break
}

#Create logfiles
$logFile = ".\Check-MergeFile_LogFile$(Get-Date -f yyyy-MM-dd-hh-mm-ss).txt"
$userErrorFile = ".\Check-MergeFile_UserFailures_$(Get-Date -f yyyy-MM-dd-hh-mm-ss).txt"
$workstationErrorFile = ".\Check-MergeFile_WorkstationFailures_$(Get-Date -f yyyy-MM-dd-hh-mm-ss).txt"
$parentFolder = (Get-Item $pwd).Name
$ContactMergeFile = "ContactMerge_$parentFolder.csv"
$userMigrationFile = "UserMigration_$parentFolder.txt"
$WorkstationMigrationFile = "WorkstationMigration_$parentFolder.txt"
$ProfileRemovalFile = "ProfileRemoval_$parentFolder.csv"

Import-Module ActiveDirectory

#Validate params
if(-not (Test-Path $SourceCSVFile)){
    Add-Content -Path $logFile -Value "$(get-date -f s) Path for SourceCSVFile $SourceCSVFile is invalid, script will exit" -PassThru | `
        Write-Host "Path for ContactMerge File $SourceCSVFile is invalid" -ForegroundColor Red
    exit
}

#Create hashTable for source DCs
$sourceDCList = @{}
$sourceDomainShort = @{}

#Create Exchange Server connection
#Try to create a remote PSSession with Exchange server
try{
    Add-Content -path $logFile -value "$(get-date -f s) Establishing Exchange PSSession with server $targetExchangeServer" -PassThru | Write-Output | Write-Verbose
    $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$targetExchangeServer/powershell"
    Import-PSSession $session
    Set-ADServerSettings -PreferredServer $targetDC
}
catch{
    Add-Content -path $logFile -value "$(get-date -f s) Could not establish remote PSSession with Exchange server $exchangeServer with error: $_" -PassThru | Write-Host -ForegroundColor Red
    break
}

<#
############################################################################################################################
    Helper Functions
############################################################################################################################
#>

#Function checks to see whether a username matches a legitimate user object
#Returns true or false
Function CheckADAccount{
    [CmdletBinding()]
    param(
        [string] $username,
        [string] $domainController
    )
    try{
        Get-ADUser -Identity $username -Server $domainController | Out-Null
        return $true
    }
    catch{
        return $false
    }
}

#Function retrieves the primary email address for the submitted user account
#Returns the email address or null
Function GetEmailAddress{
    [CmdletBinding()]
    param(
        [string] $username,
        [string] $domainController
    )
    try{
        $mail = (Get-ADUser -Identity $username -Server $domainController -Properties mail).mail
        return $mail
    }
    catch{
        return $null
    }
}

#Function retrieves a user account with the submitted email address
#Returns the samAccountName or the matching user or $null if not found
Function GetMatchingAccount{
    [CmdletBinding()]
    param(
        [string] $emailAddress,
        [string] $tDC
    )
    $matchingUser = (get-recipient $emailAddress -DomainController $tDC -ErrorAction SilentlyContinue).samAccountName
    if($matchingUser.count -gt 1){
        throw "Multiple recipients found with email address $emailAddress"
        break
    }
    return $matchingUser
}

#Function retrieves the recipient type for an emailAddress
#Returns the recipient type or $null if not found
Function GetAccountType{
    [CmdletBinding()]
    param(
        [string] $emailAddress,
        [string] $tDC
    )
    $type = (get-recipient $emailAddress -DomainController $tDC -ErrorAction SilentlyContinue).RecipientType
    return $type
}

#Function retrieves the sid for a source user and checks if that sid is in the sidhistory
#of the target user.
#Returns true or false
Function CheckSidHistory{
    [CmdletBinding()]
    param(
        [string] $sUser,
        [string] $sDC,
        [string] $tUser,
        [string] $tDC
    )
    try{
        $sid = (Get-ADUser -Identity $sUser -Server $sDC -Properties SID).sid
    }
    catch{
        throw "Cannot find $sUser on $sDC.  Error is: $_"
        break
    }
    try{
        $sidHistory = (Get-ADUser -Identity $tUser -Server $tDC -Properties sidHistory).sidHistory
    }
    catch{
        throw "Cannot find $tUser on $tDC.  Error is: $_"
        break
    }
    if($sidHistory -contains $sid){
        return $true
    }
    else{
        return $false
    }

}

#Function tries to find a matching contact for an emailaddress and compares the
#mail aliases to see if they match.
#Returns the alias if they match, returns null otherwise
Function CheckContact{
    [CmdletBinding()]
    param(
        [string] $emailAddress,
        [string] $sUser,
        [string] $sDC,
        [string] $tDC
    )
    
    $alias = (Get-ADUser -Identity $sUser -Server $sDC -Properties mailNickName -ErrorAction SilentlyContinue).mailNickName
    if($alias -eq $null){
        throw "Alias for $sUser not found on $sDC"
        break
    }
    $contact = Get-MailContact $emailAddress -DomainController $tDC -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
    if($contact -eq $null){
        throw "Contact for $emailAddress not found on $tDC"
    }
    if($contact.alias -eq $alias){
        return $alias
    }
    else{
        return $null
    }
}

#Function gets the proxy addresses for a source user and tries to find an address ending in @sourcedomain.com
#Returns the address or $null if not found
Function GetSourceDomainAddress{
    [CmdletBinding()]
    param(
        [string] $user,
        [string] $sDC
    )
    $ADUser = Get-ADUser -Identity $user -Server $sDC -Properties proxyAddresses,mailNickname -ErrorAction SilentlyContinue
    if($ADUser){
        $SourceDomainAddress = $ADUser.proxyAddresses | where{$_ -like "*@sourcedomain.com"}
        if($SourceDomainAddress.count -gt 1){
            throw "Multiple email addresses ending in @sourcedomain.com were found for user $user"
            break
        }
        $SourceDomainAddress = $SourceDomainAddress.split(":")[1]
        return $SourceDomainAddress

    }
    else{
        return $null
    }
}

#Function checks to see if there are more than one recipient with the same alias.  If
#so the mailbox object should be the target user for replication
Function CheckContactAlias{
    [CmdletBinding()]
    param(
        [string] $ContactAlias,
        [string] $targetUser
    )
    $rec = Get-Recipient $ContactAlias
    if($rec.count -gt 1){
        $rec = $rec | ?{$_.recipienttype -ne "MailContact"}
        $rec = $rec | ?{$_.samAccountName -ne $targetUser}
        if($rec.count -ne 0){
            return $false
        }
        else{
            return $true
        }
    }
    else{
        return $true
    }
}

#Check to see if there is another mailbox with the same alias as the target user
#Return true if there is a duplicate
Function CheckDuplicateAlias{
    [CmdletBinding()]
    param(
        [string] $targetUser,
        [string] $targetDC
    )
    $alias = (Get-ADUser -Identity $targetUser -Server $targetDC -Properties mailNickName).mailNickName
    $mbxs = Get-Mailbox $alias
    if($mbxs.count -gt 1){
        return $true
    }
    else{
        return $false
    }
}

#Check the workstations for DNS resolution
#Returns true or false
Function CheckWorkstationDNS{
    [CmdletBinding()]

    param(
        [Parameter(Mandatory=$true)]
        [string] $computerName,
        [Parameter(Mandatory=$true)]
        [string] $DomainName
    )

    #Convert computer name to FQDN
    $computerName = "$computerName.$DomainName"
    
    try{
        $ws = [Net.DNS]::GetHostEntry($computerName)
        return $true
    }
    catch{
        return $false
    }
}


#Parse CSV
Add-Content -path $logFile -value "$(get-date -f s) Loading content from $SourceCSVFile" -PassThru | Write-Output | Write-Verbose
$CSV =  Import-Csv $SourceCSVFile -Delimiter "`t"

if($CheckUsers){
    #Check the user info and create files
    $errorCount = 0

    #array to hold user info
    $userCSV = @()

    Add-Content -path $logFile -value "$(get-date -f s) Parsing user entries in $SourceCSVFile" -PassThru | Write-Output | Write-Verbose
    foreach($entry in $CSV){
        #Check if entry is null, log as an error if one field is blank, log as a warning if the whole line is blank
        if($entry.'Source User ID'.Length -eq 0 -or $entry.'Current Domain'.Length -eq 0){
            if($entry.'Source User ID'.Length -eq 0 -and $entry.'Current Domain'.Length -eq 0){
                Add-Content -path $logFile -value "$(get-date -f s) Blank line detected in CSV, skipping line" -PassThru | Write-Output | Write-Host -ForegroundColor Yellow
            }
            else{
                Add-Content -path $logFile -value "$(get-date -f s) Blank entry for Source User ID or Current Domain detected in CSV" -PassThru | `
                    Add-Content -path $userErrorFile -PassThru | `
                    Write-Host -ForegroundColor Red
                    $errorCount++
            }
            continue
        }
        #Create custom PS Object to hold output
        $obj = New-Object psobject

        #Clean up user data
        $user =  ($entry."Source User ID").Trim()

        #Clean up source domain
        $sourceDomain = ($entry."Current Domain").Trim()
        $sourceDC = $sourceDCList.$sourceDomain

        #Test if source DC was matched from Source domain value
        if($sourceDC -eq $null){
            Add-Content -path $logFile -value "$(get-date -f s) Source Domain of $sourceDomain for user $user does not match a valid value, please check source CSV" -PassThru | `
                Add-Content -path $userErrorFile -PassThru | `
                Write-Host -ForegroundColor Red
            $errorCount++
            continue
        }

        #Verify user exists in source domain
        if(CheckADAccount -username $user -domainController $sourceDC){
            Add-Content -path $logFile -value "$(get-date -f s) User $user was found on $sourceDomain" -PassThru | Write-Output | Write-Verbose
        }
        else{
            Add-Content -path $logFile -value "$(get-date -f s) User $user was not found on domain controller $sourceDC" -PassThru | `
                Add-Content -path $userErrorFile -PassThru | `
                Write-Host -ForegroundColor Red
            $errorCount++
            continue
        }

        #Get the mail attribute from the source user
        $sourcePrimaryEmail = GetEmailAddress -username $user -domainController $sourceDC
        if($sourcePrimaryEmail){
            Add-Content -path $logFile -value "$(get-date -f s) User $user has a primary email address of $sourcePrimaryEmail" -PassThru | Write-Output | Write-Verbose
        }
        else{
            Add-Content -path $logFile -value "$(get-date -f s) User $user does not have a primary email address" -PassThru | `
                Add-Content -path $userErrorFile -PassThru | `
                Write-Host -ForegroundColor Red
            $errorCount++
            continue
        }

        #Try and get the targetUser based on the mail address
        try{
            $targetUser = GetMatchingAccount -emailAddress $sourcePrimaryEmail -tDC $targetDC
        }
        catch{
            Add-Content -path $logFile -value "$(get-date -f s) Error encountered when searching for target user with email address $sourcePrimaryEmail.  Error is: $_" -PassThru | `
                Add-Content -path $userErrorFile -PassThru | `
                Write-Host -ForegroundColor Red
            $errorCount++
            continue
        }
        if($targetUser){
            Add-Content -path $logFile -value "$(get-date -f s) User $user has a matching primary email address of $sourcePrimaryEmail with $targetUser in the target domain" -PassThru | `
                Write-Output | Write-Verbose
        }
        else{
            Add-Content -path $logFile -value "$(get-date -f s) An account with a matching primary email address of $sourcePrimaryEmail could not be found for user $user in the target domain" -PassThru | `
                Add-Content -path $userErrorFile -PassThru | `
                Write-Host -ForegroundColor Red
            $errorCount++
            continue
        }

        #Make sure the matching user is a mailbox user and not a mail-enabled user
        $RecipientType = GetAccountType -emailAddress $sourcePrimaryEmail -tDC $targetDC
        if($RecipientType -eq "UserMailbox"){
            Add-Content -path $logFile -value "$(get-date -f s) Target Account $targetUser is Mailbox enabled as expected" -PassThru | `
                Write-Output | Write-Verbose
        }
        else{
            Add-Content -path $logFile -value "$(get-date -f s) Target Account $targetUser is not a UserMailbox as expected.  It is a $RecipientType.  Has this account already been migrated?" -PassThru | `
                Add-Content -path $userErrorFile -PassThru | `
                Write-Host -ForegroundColor Red
            $errorCount++
            continue
        }

        #Verify that the target user has the source user sid in its sidhistory
        if(CheckSidHistory -sUser $user -sDC $sourceDC -tUser $targetUser -tDC $targetDC){
            Add-Content -path $logFile -value "$(get-date -f s) User $user has a matching SID in the sidHistory of $targetUser" -PassThru | `
                Write-Output | Write-Verbose
        }
        else{
            Add-Content -path $logFile -value "$(get-date -f s) User $user does not have a matching SID in the sidHistory of $targetUser" -PassThru | `
                Add-Content -path $userErrorFile -PassThru | `
                Write-Host -ForegroundColor Red
            $errorCount++
            continue
        }

        #Make sure there is no duplicate alias for the target user
        if(CheckDuplicateAlias -targetUser $targetUser -targetDC $targetDC){
            Add-Content -path $logFile -value "$(get-date -f s) $targetUser has a duplicate mail alias with another mailbox.  Update alias to samAccountName." -PassThru | `
                Add-Content -path $userErrorFile -PassThru | `
                Write-Host -ForegroundColor Red
            $errorCount++
            continue
        }

        #Get the source address that will correspond to the contact
        try{
            $ContactAddress = GetSourceDomainAddress -user $user -sDC $sourceDC
        }
        catch{
            Add-Content -path $logFile -value "$(get-date -f s) Error encountered when retrieving the source address for $user.  Error is: $_" -PassThru | `
                Add-Content -path $userErrorFile -PassThru | `
                Write-Host -ForegroundColor Red
            $errorCount++
            continue
        }
        if($ContactAddress){
            Add-Content -path $logFile -value "$(get-date -f s) User $user has a source address of $ContactAddress" -PassThru | `
                Write-Output | Write-Verbose
        }
        else{
            Add-Content -path $logFile -value "$(get-date -f s) User $user does not have an email address ending in @sourcedomain.com" -PassThru | `
                Add-Content -path $userErrorFile -PassThru | `
                Write-Host -ForegroundColor Red
            $errorCount++
            continue
        }

        #Get the alias of the target contact
        try{
            $targetContact = CheckContact -emailAddress $ContactAddress -sUser $user -sDC $sourceDC -tDC $targetDC
        }
        catch{
            Add-Content -path $logFile -value "$(get-date -f s) Error encountered retrieving the alias for $user using email address $ContactAddress.  Error is: $_" -PassThru | `
                Add-Content -path $userErrorFile -PassThru | `
                Write-Host -ForegroundColor Red
            $errorCount++
            continue
        }
        if($targetContact){
            Add-Content -path $logFile -value "$(get-date -f s) Target contact in target domain has an alias of $targetContact" -PassThru | `
                Write-Output | Write-Verbose
        }
        else{
            Add-Content -path $logFile -value "$(get-date -f s) A target contact could not be found in target domain for source user $user with source address $ContactAddress" -PassThru | `
                Add-Content -path $userErrorFile -PassThru | `
                Write-Host -ForegroundColor Red
            $errorCount++
            continue
        }

        #If we've made it this far in the loop, then there were no errors
        Add-Content -path $logFile -value "$(get-date -f s) No errors were encountered for source user $user.  The following entries will be used Source User ID: $user, Target User ID: $targetUser, Target Contact Alias: $targetContact" -PassThru | `
                Write-Output | Write-Verbose

        #Add the values to the custom object and add it to the array
        Add-Member -InputObject $obj -MemberType NoteProperty -Name "SourceUserName" -Value $user
        Add-Member -InputObject $obj -MemberType NoteProperty -Name "TargetUserName" -Value $targetUser
        Add-Member -InputObject $obj -MemberType NoteProperty -Name "TargetContactAlias" -Value $targetContact
        Add-Member -InputObject $obj -MemberType NoteProperty -Name "SourceDomain" -Value $sourceDomain
        Add-Member -InputObject $obj -MemberType NoteProperty -Name "SourcePrimaryEmail" -Value $sourcePrimaryEmail
        $userCSV += $obj

    }

    #If there were errors, do not create merge files
    if($errorCount -gt 0){
        Add-Content -path $logFile "$(get-date -f s) Errors were encountered when checking the user objects.  The errors can be found in the file $userErrorFile.  The user merge files will not be created until no errors are encountered." -PassThru | `
        Add-Content -path $userErrorFile -PassThru | `
        Write-Host -ForegroundColor Yellow
    }
    else{
        Add-Content -path $logFile -value "$(get-date -f s) No errors were encountered when checking the user objects.  The merge files will now be create in the current directory" -PassThru | `
                Write-Output | Write-Verbose
        $userCSV | Export-Csv -Path $ContactMergeFile -NoTypeInformation -Encoding UTF8
        Set-Content $userMigrationFile -Value "samAccountName`tsamAccountName"
        foreach($user in $userCSV){
            Add-Content $userMigrationFile -Value "$($user.SourceUsername)`t$($user.TargetUserName)"
        }
    }
}


if($CheckWorkstations){
    Add-Content -Path $logFile -Value "$(get-date -f s) Checking Workstation DNS resolution" -PassThru | `
        Write-Output | Write-Verbose
    $errorCount = 0
    $workstationCSV = @()

    foreach($entry in $CSV){
        $obj = New-Object psobject

        #Check if entry is null, log as an error if one field is blank, log as a warning if the whole line is blank
        if($entry.'Source User ID'.Length -eq 0 -or $entry.'Current Domain'.Length -eq 0 -or $entry.'Computer Name'.Length -eq 0){
            if($entry.'Source User ID'.Length -eq 0 -and $entry.'Current Domain'.Length -eq 0 -and $entry.'Computer Name'.Length -eq 0){
                Add-Content -path $logFile -value "$(get-date -f s) Blank line detected in CSV, skipping line" -PassThru | Write-Output | Write-Host -ForegroundColor Yellow
            }
            else{
                Add-Content -path $logFile -value "$(get-date -f s) Blank entry for Source User ID, Current Domain, or Computer Name detected in CSV" -PassThru | `
                    Add-Content -path $workstationErrorFile -PassThru | `
                    Write-Host -ForegroundColor Red
                    $errorCount++
            }
            continue
        }

        #Cleanup input
        $workstation = ($entry."Computer Name").trim()
        $domain = ($entry."Current Domain").trim()
        $shortDomain = $sourceDomainShort.$domain
        $sourceDC = $sourceDCList.$domain
        $user = ($entry."Source User ID").trim()

        #Test if source DC was matched from Source domain value
        if($sourceDC -eq $null){
            Add-Content -path $logFile -value "$(get-date -f s) Source Domain of $sourceDomain for user $user does not match a valid value, please check source CSV" -PassThru | `
                Add-Content -path $workstationErrorFile -PassThru | `
                Write-Host -ForegroundColor Red
            $errorCount++
            continue
        }

        #Test to make sure the user exists in active directory
        if(CheckADAccount -username $user -domainController $sourceDC){
            Add-Content -path $logFile -value "$(get-date -f s) User $user was found on $domain" -PassThru | Write-Output | Write-Verbose
        }
        else{
            Add-Content -path $logFile -value "$(get-date -f s) User $user was not found on domain controller $sourceDC" -PassThru | `
                Add-Content -path $workstationErrorFile -PassThru | `
                Write-Host -ForegroundColor Red
            $errorCount++
            continue
        }
        
        #Get the mail attribute from the source user
        $sourcePrimaryEmail = GetEmailAddress -username $user -domainController $sourceDC
        if($sourcePrimaryEmail){
            Add-Content -path $logFile -value "$(get-date -f s) User $user has a primary email address of $sourcePrimaryEmail" -PassThru | Write-Output | Write-Verbose
        }
        else{
            Add-Content -path $logFile -value "$(get-date -f s) User $user does not have a primary email address" -PassThru | `
            Add-Content -path $workstationErrorFile -PassThru | `
            Write-Host -ForegroundColor Red
            $errorCount++
            continue
        }

        #Try and get the targetUser based on the mail address
        try{
            $targetUser = GetMatchingAccount -emailAddress $sourcePrimaryEmail -tDC $targetDC
        }
        catch{
            Add-Content -path $logFile -value "$(get-date -f s) Error encountered when searching for target user with email address $sourcePrimaryEmail.  Error is: $_" -PassThru | `
            Add-Content -path $workstationErrorFile -PassThru | `
            Write-Host -ForegroundColor Red
            $errorCount++
            continue
        }
        if($targetUser){
            Add-Content -path $logFile -value "$(get-date -f s) User $user has a matching primary email address of $sourcePrimaryEmail with $targetUser in the Target domain" -PassThru | `
                Write-Output | Write-Verbose
        }
        else{
            Add-Content -path $logFile -value "$(get-date -f s) An account with a matching primary email address of $sourcePrimaryEmail could not be found for user $user in the target domain" -PassThru | `
                Add-Content -path $workstationErrorFile -PassThru | `
                Write-Host -ForegroundColor Red
            $errorCount++
            continue
        }

        #Verify that the target user has the source user sid in its sidhistory
        if(CheckSidHistory -sUser $user -sDC $sourceDC -tUser $targetUser -tDC $targetDC){
            Add-Content -path $logFile -value "$(get-date -f s) User $user has a matching SID in the sidHistory of $targetUser" -PassThru | `
                Write-Output | Write-Verbose
        }
        else{
            Add-Content -path $logFile -value "$(get-date -f s) User $user does not have a matching SID in the sidHistory of $targetUser" -PassThru | `
                Add-Content -path $workstationErrorFile -PassThru | `
                Write-Host -ForegroundColor Red
            $errorCount++
            continue
        }

        #Check workstation DNS resolution and connectivity
        Add-Content -Path $logFile -Value "$(get-date -f s) Checking Workstation DNS resolution and ping for $workstation" -PassThru | `
            Write-Output | Write-Verbose 
        if(CheckWorkstationDNS -computerName $workstation -DomainName $domain){
            Add-Content -Path $logFile -Value "$(get-date -f s) Workstation DNS resolution and ping for $workstation was successful" -PassThru | `
            Write-Output | Write-Verbose
            Add-Member -InputObject $obj -MemberType NoteProperty -Name "TargetUserName" -Value $targetUser
            Add-Member -InputObject $obj -MemberType NoteProperty -Name "ComputerName" -Value $workstation
            Add-Member -InputObject $obj -MemberType NoteProperty -Name "SourceDomainShort" -Value $shortDomain
            Add-Member -InputObject $obj -MemberType NoteProperty -Name "SourceDomain" -Value $domain
            $workstationCSV += $obj
        }
        else{
            Add-Content -path $logFile -value "$(get-date -f s) Workstation DNS resolution and ping for $workstation FAILED" -PassThru | `
                Add-Content -path $workstationErrorFile -PassThru | `
                Write-Host -ForegroundColor Red
            $errorCount++
            continue
       }
    }

    #If there were errors do not create the files
    if($errorCount -gt 0){
        Add-Content -path $logFile "$(get-date -f s) Errors were encountered when checking the workstation objects.  The errors can be found in the file $workstationErrorFile.  The workstation merge files will not be created until no errors are encountered." | `
        Add-Content -path $workstationErrorFile -PassThru | `
        Write-Host -ForegroundColor Yellow
    }
    else{
        Add-Content -path $logFile -value "$(get-date -f s) No errors were encountered when checking the workstation objects.  The merge files will now be create in the current directory" -PassThru | `
            Write-Output | Write-Verbose
        $workstationCSV | Export-Csv $ProfileRemovalFile -NoTypeInformation -Encoding UTF8
        Set-Content -Path $WorkstationMigrationFile -Value ""
        foreach($entry in $workstationCSV){
            Add-Content -Path $WorkstationMigrationFile -Value "$($entry.ComputerName)`t$($entry.SourceDomainShort)"
        }
    }
}

#Always cleanup your sessions!
Remove-PSSession $session