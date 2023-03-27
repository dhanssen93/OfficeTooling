function Get-AddressType {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true,HelpMessage="Enter the address that you want to search for",ValueFromPipeline=$true,
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
        [Parameter(Mandatory=$true,HelpMessage="Enter the address that you want to search the owner(s) for",ValueFromPipeline=$true,
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
        [Parameter(Mandatory=$true,HelpMessage="Enter the mailaddress from the user that you want the permissions to find for",ValueFromPipeline=$true,
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
            $message = Read-Host "There are now results. Do you want to return to the menu? y"   
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
        [Parameter(Mandatory=$true,HelpMessage="Enter the mailaddress of the mailbox(es) that you want to enter the owner on.",ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)]
        [string[]]$mailboxes
    )
    BEGIN{
        $user = Read-Host "Enter the mailaddress of the user that you want to make an owner."
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
            Write-Host "The mailaddress of the owner that you entered does not exist" -ForegroundColor Red
        }
        
    }
    END{}
}

function Replace-Owner {
    [Cmdletbinding()]
    param(
        [Parameter(Mandatory=$true,HelpMessage="Enter the mailaddress of the mailbox(es) that you want to replace the owner off.",ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)]
        [string[]]$mailboxes
    )
    BEGIN{
        $old = Read-Host "Enter the mailaddress of the owner that you want to replace."
        $new = Read-Host "Enter the mailaddress of the user that you want to make the owner."
    }
    PROCESS{
        if ((Get-Mailbox -Identity $new)) {
            foreach($mailbox in $mailboxes) {
                if((Get-Mailbox -Identity $mailbox)) {
                    $CurrentOwners = (Get-Mailbox $mailbox)
                    $OldOwner = $old.Split("@")[0]
                    $NewOwner = $new.Split("@")[0]

                    $OwnerAttribute = $CurrentOwners.CustomAttribute1
                    $FinalAttribute = $OwnerAttribute.Replace("$OldOwner","$NewOwner")

                    Set-Mailbox $mailbox -CustomAttribute1 $FinalAttribute    
                    Write-Host "The user $new has been added as owner of the mailbox $mailbox." -ForegroundColor Green
                }
                else {
                    Write-Host "The mailbox $mailbox does not exist!" -ForegroundColor Red
                }
            }
        }
        else {
            Write-Host "The mailaddress of the new owner that you entered does not exist" -ForegroundColor Red
        }
        
    }
    END{}
}

function Get-MailboxPermissions {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true,HelpMessage="Enter the mailaddress that you want to search the permissions for")]
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
******************************************************************************
Welcome into the management script! Make sure you are signed in into EXO!

**Know issues**
When using option 1 and 2 and entering just 1 address there will be now
result. Always enter 2 addresses. This could be the same as the first one.

Make a selection...

1 = Check address type
2 = Check for the owner(s) of an address
3 = Find all mailbox permissions for a specific user.
4 = Check who has permissions on specific mailbox(es).
5 = Add new mailbox owner
6 = Replace a mailbox owner
7 = Remove a mailbox owner
8 = NEW!

x = Exit this script

******************************************************************************

Enter your choice:
"@

function Start-Script {
    [CmdletBinding()] 
    param()

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
                Write-Host "Still in development!"
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