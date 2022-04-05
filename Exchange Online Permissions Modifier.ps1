# Import ExchangeOnlineManagement module so these commands work.
Import-Module ExchangeOnlineManagement

# Propmts user for tenant email address and assigns it to a variable.
$companyExchTenant = Read-Host("Exchange Online tenant email address")

# Connects to Company Exchange Online tenant.
Connect-ExchangeOnline -UserPrincipalName $companyExchTenant

# Choice menu
Function Show-Menu {
    Write-Host ""
    Write-Host "################ What do you want to do? #################"
    Write-Host "#                                                        #"
    Write-Host "#   1. Add read permissions to a mailbox.                #"
    Write-Host "#   2. Remove read permissions from a mailbox.           #"
    Write-Host "#   3. List all Exchange Online mailboxes.               #"
    Write-Host "#   4. List current permissions on a mailbox.            #"
    Write-Host "#                                                        #"
    Write-Host "#################### Enter q to quit #####################"
}
       
do{
    Show-Menu
    $selection = Read-Host("Your selection")
    
    # Add permissions
    if($selection -eq "1"){
        Write-Host("Enter full name or entire email address.") -ForegroundColor Yellow
        $mailboxToChange = Read-Host("Mailbox to add permissions to")
        $userToAdd = Read-Host("User to add to $mailboxToChange")
        Write-Host("Granting $userToAdd read permissions on $mailboxToChange's mailbox...")
        Add-MailboxPermission -Identity $mailboxToChange -User $userToAdd -AccessRights ReadPermission

    }
    
    # Remove permissions
    elseif($selection -eq "2"){
        Write-Host("Enter full name or entire email address.") -ForegroundColor Yellow
        $mailboxToChange = Read-Host("Mailbox to remove permissions from")
        $userToRemove = Read-Host("User to remove from $mailboxToChange")
        Write-Host("Removing $userToRemove from $mailboxToChange's mailbox...")
        Remove-MailboxPermission -Identity $mailboxToChange -User $userToRemove -AccessRights ReadPermission
    }
    
    # List all Exchange Online mailboxes
    elseif($selection -eq "3"){
        Write-Host("Listing all mailboxes...")
        Get-Mailbox -ResultSize unlimited | Select-Object -Property Name,Alias,PrimarySmtpAddress | Format-Table -Property @{e='Name'; width=32},@{e='Alias';width=32},@{e='PrimarySmtpAddress';width=32}
        }

    # List permissions on a specific mailbox
    elseif($selection -eq "4"){
        WRite-Host("Enter alias or full email address.") -ForegroundColor Yellow
        $mailboxPermissions = Read-Host("Enter the mailbox do you want to see permissions for")
        Get-MailboxPermission $mailboxPermissions
    }
}until($selection -eq "q")