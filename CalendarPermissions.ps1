#Set variables
$AdminUsername = "admin@tenant.onmicrosoft.com"
$AdminPassword = "xxxxxx"
$SecurePassword = ConvertTo-SecureString $AdminPassword -AsPlainText -Force
$credential = New-Object -TypeName System.Management.Automation.PSCredential -argumentlist $AdminUsername,$SecurePassword

#Connect session
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Credential -Authentication Basic -AllowRedirection 
Import-PSSession $Session

#Generate a list of all mailboxes
$users = Get-Mailbox -Resultsize Unlimited | where-object {
    $_.name -match '\D\d{6}$'
}

#Set Default access to Reviewer for all mailboxes
foreach ($user in $users) {
$language = Get-MailboxRegionalConfiguration -Identity $user
Write-Host -ForegroundColor green "Setting permission for $($user.alias)..."
IF($language.Language -eq "da-DK")
    { 
    Set-MailboxFolderPermission -Identity "$($user.alias):\Kalender" -User Default -AccessRights Reviewer
    }
IF($language.Language -eq "en-US" -or $language.Language -eq "en-UK")
    { 
    Set-MailboxFolderPermission -Identity "$($user.alias):\Calendar" -User Default -AccessRights Reviewer
    }
}
