### Modules ###
# Add check to ensure the following modules are installed
# AzureAD, MsolService, EXOPS

### Variables ###
$logDate = Get-Date -UFormat "%Y%m%d"
$logFh = ".\logs\$($logDate)_termlog.log"
$errChk = $null
$employeeOU = "<set value>" # Main OU for user objects
$shortTermOU = "<set value>" # Short term storage for terminated user objects
$longTermOU = "<set value>" # Long term storage for terminated user objects
$termedOU = "<set value>" # Staging OU for user objects prior to complete deletion
$termBackups = "<set value>" # Folder for backing up a PC or other files too if needed

$serviceAccount = "<service.account@foo.com>"
$servicePassword = Get-Content .\cred.txt | ConvertTo-SecureString
$creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $serviceAccount,$servicePassword
$exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $creds -Authentication Basic -AllowRedirection


### Functions ###
function RemoveGroups ($myUser) {
    # Remove Local Groups
    $adGroups = Get-ADPrincipalGroupMembership -Identity $myUser | Where-Object { $_.Name -ne "Domain Users" }
    if ( $null -ne $adGroups ) {
        Remove-ADPrincipalGroupMembership -Identity $myUser -MemberOf $adGroups -Confirm:$false
    }

    # Remove Online Groups
    $userObjId = ( Get-AzureADUser -Filter "userPrincipalName eq '$($myUser.UserPrincipalName)'" ).ObjectId
    $userGroups = Get-AzureADUserMembership -ObjectId $userObjId | Select-Object DisplayName
    $userOwnedGroups = Get-UnifiedGroup | Where-Object { (Get-UnifiedGroupLinks $_.Alias -LinkType Owners| foreach {$_.name}) -contains $myUser.Name}

    foreach ( $group in $userGroups ) {
        Remove-UnifiedGroupLinks -Identity $group.DisplayName -LinkType Members -Links $myUser.UserPrincipalName -Confirm:$false -ErrorAction "SilentlyContinue"
    }
    foreach ( $group in $userOwnedGroups ) {
        Remove-UnifiedGroupLinks -Identity $group.DisplayName -LinkType Owners -Links $myUser.UserPrincipalName -Confirm:$false -ErrorAction "SilentlyContinue"
    }
}

function MoveTermedUser ($myUser) {
    if ( $myUser.comment -eq "lithold" ) {
        Move-ADObject -Identity $user -TargetPath $longTermOU
    } else {
        Move-ADObject -Identity $user -TargetPath $shortTermOU
    }
}

function Logger ($msg) {
    $timeStamp = get-date -Format "HH:mm:ss"
    "$($timeStamp) - $($msg)" | Out-File -FilePath $logFh -Append
    return 1
}

### Main ###
# Get listing of users with term date of today or earlier
$currentTerminations = Get-ADUser -Filter "extensionAttribute2 -like '*'" -SearchBase $employeeOU -SearchScope 1 -Properties extensionAttribute2,manager,comment
if ( $null -eq $currentTerminations ) {
    exit
}

# Connect to online powershell sessions
Import-PSSession $exchangeSession -DisableNameChecking
Connect-AzureAD -Credential $creds
Connect-MsolService -Credential $creds

# Loop through list of users and determine if their last date of employement has passed, if so take steps to remove access
foreach ( $user in $currentTerminations ) {
    if ( (Get-Date) -gt ($user.extensionAttribute2) ) {
        ## Onprem AD Actions
        try {
            $managerObj = Get-ADUser $user.Manager
        } catch {
            $errChk = Logger "No manager value found for $($user.SamAccountname)"
        }
        # Disable User, Clear Company Field, Add description w term date, ensure the mailNickname is set, mark the mailbox as hidden
        Set-ADUser -Identity $user -Enabled:$false -Description "Termed $($user.extensionAttribute2)" -Replace @{mailNickname=$user.SamAccountName;msExchHideFromAddressLists=$true} -Clear Company
        
        # Remove user from groups
        RemoveGroups $user
        
        # Move the termed user to the correct Term OU
        MoveTermedUser $user

        ## Office 365 Actions
        # Initiate Sign-Out
        Revoke-AzureADUserAllRefreshToken -ObjectId $user.UserPrincipalName

        # Convert Mailbox to shared and grant manager access
        try {
            Set-Mailbox $user.UserPrincipalName -Type Shared
            Add-MailboxPermission -Identity $user.UserPrincipalName -User $managerObj.UserPrincipalName -AccessRights FullAccess -InheritanceType All -AutoMapping $true
        }
        catch {
            $errChk = Logger "Unable to set mailbox as Type: Shared for User: $($user.SamAccountName)"
        }

        # Remove Product Licenses
        try {
            $licList = ( Get-MsolUser -UserPrincipalName $user.UserPrincipalName ).Licenses
            foreach ( $lic in $licList ) {
                Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -RemoveLicenses $lic.AccountSkuId
            }
        } catch {
            $errChk = Logger "Failed to remove licenses for User: $($user.SamAccountName)"
        }

        # Create Terms Folder
        New-Item -Path $termBackups -Name $user.SamAccountName -ItemType Directory
    }
}

# Cleanup > 1 year terms from ShortTerm OU to Termed OU
$retentionDays = (Get-Date).AddYears(-1)
$termedUsers = Get-ADUser -Filter "extensionAttribute2 -like '*'" -SearchBase $shortTermOU -Properties extensionAttribute2
foreach ( $user in $termedUsers ) {
    if ( $retentionDays -gt ($user.extensionAttribute2) ) {
        Move-ADObject -Identity $user -TargetPath $termedOU
    }
}

# Cleanup > 1 year + 60 days terms from Termed OU
$retentionDays = (Get-Date).AddDays((-425))
$termedUsers = Get-ADUser -Filter "extensionAttribute2 -like '*'" -SearchBase $termedOU -Properties extensionAttribute2
foreach ( $user in $termedUsers ) {
    if ( $retentionDays -gt ($user.extensionAttribute2) ) {
        Remove-ADUser $user -Confirm:$false
        Remove-Item -Path "$($termBackups)$($user.SamAccountName)" -Recurse -Force
    }
}

Remove-PSSession $exchangeSession
