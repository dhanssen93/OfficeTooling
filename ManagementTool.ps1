#Start data
$location = "C:\Users\$env:Username\Desktop"
$date = (Get-Date -Format dd/MM/yyyy).Replace("-","") 

function Show-Menu {
    Clear-Host
    Write-Host "**************************************************************************************"
    Write-Host ""
    Write-Host "Welcome into the management script!"
    Write-Host ""
    Write-Host "Make a choice..."
    Write-Host ""
    Write-Host "1 = Check the address type"
    Write-Host "2 = Check who owns a particular address"
    Write-Host "3 = Find all mailbox rights for one specific user"
    Write-Host "4 = Check the rights on one or more mailboxes"
    Write-Host "5 = Add new mailbox owner"
    Write-Host "6 = Replace mailbox owner"
    Write-Host "7 = Get the size of one or more mailboxes and calculate the total"
    Write-Host "8 = NEW!"    
    Write-Host ""
    Write-Host "x = Exit script"
    Write-Host ""
    Write-Host "**************************************************************************************"
    Write-Host ""
}

function Start-InputMenu {
    Clear-Host
    Write-Host "*************************** Provide the needed information ***************************"
    Write-Host ""
    Write-Host "$description"
    Write-Host ""
}

function Test-ExchangeConnection {
    try {
        Get-Mailbox -ResultSize 1 -ErrorAction Stop -WarningAction SilentlyContinue | Out-Null
    }
    catch {
        Write-Warning "You are not connected to Exchange Online. Please sign in..."
        Start-Sleep 2
        Connect-ExchangeOnline
    }
}

function Get-AddressType {
    BEGIN{
        Test-ExchangeConnection
        $description = "Here you can enter one or multiple mailboxes to find out of what type the specific`naddress is. You can enter multiple addresses with one address on each line."
        $warning = [System.Collections.ArrayList]::new()
        Start-InputMenu
        $mailboxes = @()
        do {
            $address = (Read-Host "Enter an email address")
            if ($address -ne "") {
                $mailboxes += $address
            }
        }
        until ($address -eq "")
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
            $output = @{
                PrimarySmtpAddress = $result
                Type = $type
            }
            $outcome = New-Object -TypeName psobject -Property $output
            Write-Output $outcome
        }
        foreach($warn in $warning) {
            Write-Warning "The following mailbox does not exist: $warn!"
        }
    }
    END{}
}

function Get-Owner {
    BEGIN{
        Test-ExchangeConnection
        $description = "Here you can enter one or multiple mailboxes to check who are the owners of the`nspecific address. You can enter multiple addresses with one address on each line."
        $warning = [System.Collections.ArrayList]::new()
        Start-InputMenu
        $mailboxes = @()
        do {
            
            $address = (Read-Host "Enter an email address")
            if ($address -ne "") {
                $mailboxes += $address
            }
        }
        until ($address -eq "")
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
                    $results = "Unknown"
                    $type = "Mailbox"
                }
                else {
                    $address = $mbxOwner.PrimarySmtpAddress
                    $results = $mbxOwner.CustomAttribute1
                    $type = "Mailbox"
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
                    $results = [string]$dgOwner.ManagedBy
                    $type = "Distributiongroup"
                }
            }
            $output = @{
                PrimarySmtpAddress = $address
                Owner = $results
                Type = $type
            }
            $outcome = New-Object -TypeName psobject -Property $output
            Write-Output $outcome        
        }
        foreach($warn in $warning) {
            Write-Warning "The following mailbox does not exist: $warn!"   
        }
    }
    END{}
}   

function Get-UserMailboxPermssions {
    BEGIN{
        Test-ExchangeConnection
        $description = "Fill in the email address of the user that you want to find the permissions for.`nYou can only search for the permissions for one user."
        $name = "UserMailboxPermissions"+"_"+"$date"
        Start-InputMenu
        $user = (Read-Host "Enter an email address")
    }
    PROCESS{
        if($user -ne "") {
            try {
                $found = Get-Mailbox -Identity $user -ErrorAction Stop | Out-Null
            }
            catch {
                $found = $false
            }
            if($found) {
                Write-Host "Getting all mailboxes. This could take a while ;)" -ForegroundColor Green
                $mailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails SharedMailbox,EquipmentMailbox,RoomMailbox | Select-Object PrimarySmtpAddress
                $total = $mailboxes.count
                $current = 1
                $percentage = 0

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
                        } | Export-Excel "$location\$name.xslx" -BoldTopRow -Append -AutoSize
                        $outcome
                    }
                }
                $check = Get-Childitem -Directory "$location\$name.xslx" -ErrorAction SilentlyContinue
                if($check) {
                    $message = Read-Host "The results have been exported to"$location\$name.xslx". Do you want to return to the menu? y"             
                }
                else {
                    $message = Read-Host "No results were found. Do you want to return to the menu? y"   
                }
            }
            else {
                $message = Read-Host "This email address does not exist. Go back to the menu and try again white a valid`nemail address.`n`nPress y to go to the menu."
            }
        }
        else {
            $message = Read-Host "No email address has been entered. Go back to the menu and try again.`n`nPress y to go to the menu."
        }
        
        Clear-Host
        switch ($message) {
            y {
                Start-Tool  
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
    BEGIN{
        Test-ExchangeConnection
        $description = "Here you can add a new owner to one or more mailboxes. You can enter multiple`naddresses with one address on each line."
        Start-InputMenu
        $mailboxes = @()
        do {
            $address = (Read-Host "Enter an email address")
            if ($address -ne "") {
                $mailboxes += $address
            }
        }
        until ($address -eq "")
        $description = "Fill in the email address of the user that you want to add as owner. This can be`nonly one address."
        Start-InputMenu
        $user = (Read-Host "Enter an email address")
    }
    PROCESS{
        try {
            $mbx = Get-Mailbox -Identity $user -ErrorAction Stop
        }
        catch {
            $mbx = $false
        }
        if($mbx) {
            foreach($mailbox in $mailboxes) {
                try {
                    $mbx = Get-Mailbox -Identity $mailbox -ErrorAction Stop
                }
                catch {
                    $mbx = $false
                }
                if($mbx) {
                    $CurrentOwners = (Get-Mailbox $mailbox)
                    $owner = $user.Split("@")[0]

                    $OwnerAttribute     = $CurrentOwners.CustomAttribute1
                    $NewOwnerAttribute  = $OwnerAttribute+";$owner"
                    $FinalAttribute     = $NewOwnerAttribute.Replace(";;",";")

                    Set-Mailbox $mailbox -CustomAttribute1 $FinalAttribute
                    Write-Host "The user $user has been added as owner of the mailbox $mailbox" -ForegroundColor Green
                }
                else {
                    Write-Warning "The mailbox $mailbox does not exist!"
                }
            }
        }    
        else {
            Write-Warning "The email address $user of the owner you entered does not exist"
        }
    }
    END{}
}

function Switch-Owner {
    BEGIN{
        Test-ExchangeConnection
        $description = "Enter the email address of the owner that you want to replace on one or more`nmailboxes. You can only enter one address"
        Start-InputMenu       
        $old = (Read-Host "Enter an email address")

        $description = "Enter the email address of the owner that you want to make the new owner. You can`nonly enter one address"
        Start-InputMenu
        $new = (Read-Host "Enter an email address")
        
        $description = "Enter one or more email addresses of the mailboxes that you want to replace the`nowner of. You can enter multiple addresses with one address on each line."
        Start-InputMenu
        $mailboxes = @()
        do {
            $address = (Read-Host "Enter an email address")
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
                    $CurrentOwner = $CurrentOwners.Split(";")

                    if($CurrentOwner -contains $OldOwner) {
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
    Test-ExchangeConnection
    $description = "Enter one or more addresses that you want to find the users with permissions for.`nYou can enter multiple addresses with one address on each line."
    $name = "MailboxPermissions"+"_"+"$date"
    Start-InputMenu
    $addresses = @()
    do {
        $address = (Read-Host "Enter an email address")
        if ($address -ne "") {
            $addresses += $address
        }
    }
    until ($address -eq "")
    
    # Add try and catch
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
                } | Export-Excel "$location\$name.xslx" -WorksheetName "FullAccess" -BoldTopRow  -Append -AutoSize
            }
            foreach($user in $SendAs) {
                [PSCustomObject] @{
                    Address = $address
                    Type = "Mailbox"
                    Rights = "SendAs"
                    Member = $user.Trustee
                } | Export-Excel "$location\$name.xslx" -WorksheetName "SendAs" -BoldTopRow -Append -AutoSize
            }
            foreach($user in $FolderPermissions) {
                $gebruikers = $user.User.Displayname
                $UserMail = Get-Mailbox $gebruikers

                [PSCustomObject] @{
                    Address = $address
                    Type = "Mailbox"
                    Rights = $user.AccessRights
                    Member = $UserMail.PrimarySmtpAddress
                } | Export-Excel "$location\$name.xslx" -WorksheetName "FolderPermission" -BoldTopRow -Append -AutoSize
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
                } | Export-Excel "$location\$name.xslx" -WorksheetName "Distributionlist" -BoldTopRow -Append -AutoSize
            }
        }
        else {
            Write-Warning "$address does not exist!"
        }
    }
    $check = Get-Childitem -Directory "$location\$name.xslx" -ErrorAction SilentlyContinue
    if($check) {
        $message = Read-Host "The results have been exported to"$location\$name".`nDo you want to return to the menu? y"             
    }
    else {
        $message = Read-Host "No results have been found. Do you want to return to the menu? y"   
    }
    Clear-Host
    switch ($message) {
        y {
            Start-Tool
            #pause
        }
        Default {
            Write-Host "That is not a valid selection. Try again." -ForegroundColor Red
            pause
        }
    }
}

function Get-MailboxSizes { 
    Begin{
        Test-ExchangeConnection
        $description = "Enter one or more mailboxes where you would like to export the mailboxsize of.`nThe results will be exported to an Excel file."
        $name = "MailboxSizes"+"_"+"$date"
        Start-InputMenu
        $mailboxes = @()
        do{
            $address = (Read-Host "Enter the mailbox")
            if($address -ne "") {
                $mailboxes += $address
            }
        }
        Until($address -eq "")
    }
    Process{
        foreach($mailbox in $mailboxes) {
            try {
                $mbx = Get-Mailbox -Identity $mailbox -ErrorAction Stop
            }
            catch {
                $mbx = $false
            }
            if($mbx) {
                $ItemSize  = Get-MailboxStatistics -Identity $mailbox | Select-Object TotalItemSize
                $StringItemSize = $ItemSize.TotalItemSize.Value.ToString()
                $ReplaceItemSize = $StringItemSize -replace "^\d{1,3}.\d{1,3}\s\w{1,2}\s.",""
                $FinalItemSize = $ReplaceItemSize -replace "\s\w{1,5}.$",""
                $FinalSizeMB = [math]::round([int64]$FinalItemSize/1MB, 2)
                $FinalSizeGB = [math]::round([int64]$FinalItemSize/1GB, 2)

                #Counter for total
                $TotalSize += [int64]$FinalItemSize

                [PSCustomObject] @{
                    Mailbox = $mbx.PrimarySMTPAddress
                    Size_MB = $FinalSizeMB
                    Size_GB = $FinalSizeGB
                } | Export-Excel "$location\$name.xslx" -BoldTopRow -Append -AutoFilter -AutoSize
            }
            else {
                try {
                    $dg = Get-DistributionGroup -Identity $mailbox -ErrorAction Stop
                }
                catch {
                    Write-Warning "The email address $mailbox does not exist as a mailbox or distributiongroup."
                }
                if($dg) {
                    Write-Host "The email address $mailbox is a distributiongroup. This search does only work for mailboxes." -ForegroundColor Red 
                }
            }
        }
        $TotalMB = [math]::round($TotalSize/1MB, 2)
        $TotalGB = [math]::round($TotalSize/1GB, 2)
        
        [PSCustomObject] @{
            Mailbox = "Total"
            Size_MB = $TotalMB
            Size_GB = $TotalGB
        } | Export-Excel "$location\$name.xslx" -BoldTopRow -Append -AutoFilter -AutoSize 
        
        try {
            $check = Get-Childitem -Directory "$location\$name.xslx" -ErrorAction SilentlyContinue
        }
        catch {
            $message = Read-Host "No results have been found. Do you want to return to the menu? y"
            $check = $false 
        }
        if($check) {
            $message = Read-Host "The results have been exported to"$location\$name". `nDo you want to return to the menu? y"   
        }
        Clear-Host
        switch ($message) {
            y {
                Start-Tool
                pause
            }
            Default {
                Write-Host "That is not a valid selection. Try again." -ForegroundColor Red
                pause
            }
        }
    }
    End{}
}

function Start-Tool {
    Do{
        #cmd /c color 71
        Import-Excel
        Show-Menu
        $selection = Read-Host "Enter your choice"
        Clear-Host
        switch ($selection) {
            1 {
                Get-AddressType | Out-Host
                pause
            }
            2 {
                Get-Owner | Out-Host
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
                Switch-Owner
                pause
            }
            7 {
                Get-MailboxSizes
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