function Get-AddressType {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true,HelpMessage="Enter the email address of the mailbox whose type you want to look for",ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)]
        [string[]]$mailboxes
    )
    BEGIN{
        $warning = [System.Collections.ArrayList]::new()
    }
    PROCESS{
        foreach($mailbox in $mailboxes) {
            try {
                $mbx = Get-Mailbox -Identity $mailbox -ErrorAction Stop
            }
            catch {
                $mbx = $false
            }
            if($mbx) {
                $result = $mbx.PrimarySmtpAddress
                $type = $mbx.RecipientTypeDetails
            }
            else {
                try {
                    $dg = Get-DistributionGroup -Identity $mailbox -ErrorAction Stop
                }
                catch {
                    [void]$warning.Add($mailbox) 
                    continue                   
                }
                if($dg) {
                    $result = $dg.PrimarySmtpAddress
                    $type = $dg.RecipientTypeDetails
                }
            }
            [PSCustomObject] @{
                PrimarySmtpAddress = $result
                Type = $type
            }
        }
        foreach($warn in $warning) {
            Write-Warning "The following mailbox does not exist: $warn!"
        }
    }
    END{}
}

function Get-Owner {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true,HelpMessage="Enter the email address of the mailbox you want to find the owner of",ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)]
        [string[]]$mailboxes
    )
    BEGIN{
        $warning = [System.Collections.ArrayList]::new()
    }
    PROCESS{
        foreach($mailbox in $mailboxes) {
            try {
                $mbxOwner = Get-Mailbox $mailbox -ErrorAction Stop
            }
            catch {
                $mbxOwner = $false
            }
            if($mbxOwner) {
                if($mbxOwner.CustomAttribute1 -eq "") {
                    $address = $mbxOwner.PrimarySmtpAddress
                    $type = "Mailbox"
                    $results = "Unknown"
                }
                else {
                    $address = $mbxOwner.PrimarySmtpAddress
                    $type = "Mailbox"
                    $results = $mbxOwner.CustomAttribute1
                }
            }
            else {
                try {
                    $dgOwner = Get-DistributionGroup $mailbox -ErrorAction Stop
                }
                catch {
                    [void]$warning.Add($mailbox) 
                    continue 
                }
                if($dgOwner) {
                    $address = $dgOwner.PrimarySmtpAddress
                    $type = "Distributiongroup"
                    $results = [string]$dgOwner.ManagedBy
                }
            }
            [PSCustomObject] @{
                PrimarySmtpAddress = $address
                Type = $type
                Owner = $results
            }       
        }
        foreach($warn in $warning) {
            Write-Warning "The following mailbox does not exist: $warn!"   
        }
    }
    END{}
}   

function Get-UserMailboxPermssions {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true,HelpMessage="Enter the email address of the user whose rights you want to find",ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)]
        [string]$user
    )
    BEGIN{
        Write-Host "Getting all mailboxes. This could take a while ;)" -ForegroundColor Green
        $mailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails SharedMailbox,EquipmentMailbox,RoomMailbox | Select-Object PrimarySmtpAddress
        $total = $mailboxes.count
        $current = 1
        $percentage = 0
    }
    PROCESS{
        foreach($mailbox in $mailboxes) {
            Write-Progress -Activity "Checking for permissions" -Status "$Percentage% Complete:" -PercentComplete $percentage
            $current++
            $percentage = [int](($current / $total) * 100)
            
            $mbx = $mailbox.PrimarySmtpAddress

            $rights = ""
            $rights = Get-MailboxPermission -Identity $mbx | Where-Object {$_.user -eq "$user"} 

            if($rights) {
                $outcome = `
                [PSCustomObject]@{
                    Mailbox = $mbx
                    User = $user
                } | Export-Csv .\Search_Result.csv -Delimiter ";" -Append -NoTypeInformation     
                $outcome
            }
        }
        [string]$path = ($loc = Get-Location)
        $check = Get-Childitem -Directory $path"\Search_Result.csv" -ErrorAction SilentlyContinue

        if($check) {
            $message = Read-Host "The outcome has been exported to ""$path"". Do you want to return to the menu? y"             
        }
        else {
            $message = Read-Host "No results were found. Do you want to return to the menu? y"   
        }
        
        Clear-Host
        switch ($message) {
            y {
                $selection  
            }
            Default {
                Write-Host "That is not a valid selection. Try again." -ForegroundColor Red
                pause
            }
        }
    }
    END{}
}

function Add-Owner {
    [Cmdletbinding()]
    param(
        [Parameter(Mandatory=$true,HelpMessage="Enter the email address of the mailbox(es) where you want to add the new owner",ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)]
        [string[]]$mailboxes
    )
    BEGIN{
        $user = Read-Host "Enter the email address of the user you want to add as owner"
    }
    PROCESS{
        if ((Get-Mailbox -Identity $user)) {
            foreach($mailbox in $mailboxes) {
                if((Get-Mailbox -Identity $mailbox)) {
                    $CurrentOwners = (Get-Mailbox $mailbox)
                    $owner = $user.Split("@")[0]

                    $OwnerAttribute     = $CurrentOwners.CustomAttribute1
                    $NewOwnerAttribute  = $OwnerAttribute+";$owner"
                    $FinalAttribute     = $NewOwnerAttribute.Replace(";;",";")

                    Set-Mailbox $mailbox -CustomAttribute1 $FinalAttribute
                    Write-Host "The user $user has been added as owner of the mailbox $mailbox." -ForegroundColor Green
                }
                else {
                    Write-Host "The mailbox $mailbox does not exist!" -ForegroundColor Red
                }
            }
        }
        else {
            Write-Host "The email address of the owner you entered does not exist" -ForegroundColor Red
        }
        
    }
    END{}
}

function Replace-Owner {
    param()
    BEGIN{
        $old = Read-Host "Enter the email address of the owner you want to replace"
        $new = Read-Host "Enter the email address of the user you want to make the new owner"
        
        $mailboxes = @()
        do {
            $address = (Read-Host "Enter the email address of the mailbox(es) whose owner you want to replace")
            if ($address -ne "") {
                $mailboxes += $address
            }
        }
        until ($address -eq "")
    }
    PROCESS{
        try {
            $user = Get-Mailbox -Identity $new -ErrorAction Stop -WarningAction SilentlyContinue
        }
        catch {
            Write-Warning "The email address of the owner $new does not exist"
        }
        if($user) {
            foreach($mailbox in $mailboxes) {
                try {
                    $mbx = Get-Mailbox -Identity $mailbox -ErrorAction Stop
                }
                catch {
                    Write-Warning "The email address of the mailbox $mailbox does not exist"
                    continue
                }
                if($mbx) {
                    $OldOwner = $old.Split("@")[0]
                    $NewOwner = $new.Split("@")[0]

                    $CurrentOwners = $mbx.CustomAttribute1

                    #comparison werkt niet. Deze nog nakijken zodat gekeken wordt of de oude eigenaar er reeds in staat.
                    if($CurrentOwners -like $OldOwner) {
                        $FinalAttribute = $CurrentOwners.Replace("$OldOwner","$NewOwner")
    
                        Set-Mailbox $mailbox -CustomAttribute1 $FinalAttribute    
                        Write-Host "The user $new has replaced $old as new owner of the mailbox $mailbox" -ForegroundColor Green
                    }
                    else {
                        $FinalAttribute = $CurrentOwners+";$NewOwner"

                        Set-Mailbox $mailbox -CustomAttribute1 $FinalAttribute   
                        Write-Host "The old owner $old was not precent. The new owner $new has been added as owner of the mailbox $mailbox" -ForegroundColor Yellow
                    }
                }
            }
        }
    }  
    END{}
}

function Get-MailboxPermissions {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true,HelpMessage="Enter the email address you want to search for")]
        [string[]]$addresses
    )
    
    foreach($address in $addresses) {

        if((Get-mailbox $address)) {
            $FullAccess = Get-MailboxPermission $address | Where-object {$_.User -ne "NT AUTHORITY\SELF"}
            $SendAs = Get-RecipientPermission $address | Where-object {$_.Trustee -ne "NT AUTHORITY\SELF"}
            $FolderPermissions = Get-MailboxFolderPermission $address | Where-object {$_.User -notlike "Default" -and $_.User -notlike "Anonymous"} | Select-object User,AccessRights

            foreach($user in $FullAccess) {
                [PSCustomObject] @{
                    Address = $address
                    Type = "Mailbox"
                    Rights = "FullAccess"
                    Member = $user.User
                } 
            }
            foreach($user in $SendAs) {
                [PSCustomObject] @{
                    Address = $address
                    Type = "Mailbox"
                    Rights = "SendAs"
                    Member = $user.Trustee
                }
            }
            foreach($user in $FolderPermissions) {
                $gebruikers = $user.User.Displayname
                $UserMail = Get-Mailbox $gebruikers

                [PSCustomObject] @{
                    Address = $address
                    Type = "Mailbox"
                    Rights = $user.AccessRights
                    Member = $UserMail.PrimarySmtpAddress
                }
            }
        }
        elseif((Get-DistributionGroup $address)) {
            $DGusers = Get-DistributionGroupMember $address | Select-object Name

            foreach($user in $DGusers) {
                $UserMail = Get-mailbox $user.Name

                [PSCustomObject] @{
                    Address = $address
                    Type = "DistributionGroup"
                    Rights = "This is an member"
                    Member = $UserMail.PrimarySmtpAddress
                }
            }
        }
        else {
            Write-Host "$address does not exist!"

            [PSCustomObject] @{
                Address = $address
                Type = "Address does not exist"
                Rights = "Address does not exist"
                Member = "Address does not exist"
            }
        }
    }
}

$menu = @"
**********************************************************************************
Welcome into the management script!

**Know issue**
If only one address is entered in options 1 and 2, no result will be displayed. 
Workaround is to enter the address twice.


Make a selection...

1 = Check the address type
2 = Check who owns a particular address
3 = Find all mailbox rights for one specific user
4 = Check the rights on one or more mailboxes
5 = Add new mailbox owner
6 = Replace mailbox owner
7 = Remove a mailbox owner
8 = NEW!

x = Exit script

**********************************************************************************

Enter your choice:
"@

function Test-ExhangeConnection {
    param()
    try {
        Get-Mailbox -ResultSize 1 -ErrorAction Stop -WarningAction SilentlyContinue | Out-Null
    }
    catch {
        Write-Warning "You are not connected to Exchange Online. Please sign in..."
        Start-Sleep 2
        Connect-ExchangeOnline
    }
}

function Start-Tool {
    [CmdletBinding()]
    param()
    Test-ExhangeConnection

    Do{
        #cmd /c color 71
        $selection = Read-Host -Prompt $menu
        Clear-Host
        switch ($selection) {
            1 {
                Get-AddressType
                pause
            }
            2 {
                Get-Owner
                pause
            }
            3 {
                Get-UserMailboxPermssions
                pause
            }
            4 {
                Get-MailboxPermissions
                pause
            }
            5 {
                Add-Owner
                pause
            }
            6 {
                Replace-Owner
                pause
            }
            7 {
                Remove-Owner
                pause
            }
            8 {
                Write-Host "Still in development! Suggestions? Mail them to GitHub@visione.nl"
                pause
            }
            Default {
                Write-Host "That is not a valid selection. Try again." -ForegroundColor Red
                pause
            }
        }
    }
    Until ($selection -eq "x")
    <#Switch ($PSVersionTable.PSEdition) {
        'core' {cmd /c color 07}
        'desktop' {cmd /c color 56}
        Default {cmd /c color 56}
    }#>
}
Start-Tool