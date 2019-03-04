#---------------------------
# 
#Script to create Remote Mailbox object from local Active Directory
#
# This can be used if an Online mailbox is provisioned directly from a local AD account, and not using either Enable-RemoteMailbox or the Hybrid migration wizard
# It will modify the AD object to become a Remote Mailbox object seen from the perspective of the local Exchange server.
#
#
# The script must not be run on an Exchange server, due to some cmdlets being the same betweens Exchane On-premise and Exchange Online.
#
#---------------------------
clear
#Input parameters for running script

$cloudadmin = Get-Credential -UserName admin@tenant.onmicrosoft.com -Message "Office 365 admin credentials"
$localadmin = Get-Credential -UserName Domain\admin -Message "On-premise domain admin credentials"
$exchangeserver = "exchange.domain.local"
$domaincontroller = "dc.domain.local"
$samaccountname = "name"
$targetaddress = "$samaccountname@tenant.mail.onmicrosoft.com"
$emailaddress = "$samaccountname@brock.dk"

#fixed variables to change
$displaytype = "-2147483642"
$typedetails = "2147483648"
$recipienttype = "4"

#Check prerequisites

Write-Host Kontrollerer forudsætninger for $samaccountname -ForegroundColor Yellow
Get-ADUser -Identity $samaccountname -server $domaincontroller -ErrorAction SilentlyContinue -ErrorVariable NoUser | Out-Null
If ($NoUser)
    {
    Write-Host User not found in local directory. Exiting. -ForegroundColor Red
    BREAK
    }
$userdisplay = Get-ADUser -Identity $samaccountname -Server $domaincontroller -Properties * | Select-Object displayName
$SearchString = $null
$usermsexch = Get-ADUser -Identity $samaccountname -Server $domaincontroller -Properties * | Select-Object *msexch* | %{$_.psobject.properties} | ?{$_.Value -ne $SearchString}

IF($userdisplay.DisplayName -eq $null)
    {
    Write-Host "No displayName defined on user object. Define value and retry" -ForegroundColor Red
    BREAK
    }
IF($usermsexch -ne $null)
    {
    Write-Host "MsExch AD attribute values found. Remove values and retry" -ForegroundColor Red
    BREAK
    }

Write-Host "User object appears to meet prerequisites. Continuing." -ForegroundColor Yellow

#Connect to local Exchange server

Write-Host Connecting to $exchangeserver using $localadmin.UserName -ForegroundColor Yellow
$LocalSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$exchangeserver/PowerShell/ -Authentication Kerberos -Credential $localadmin
Import-PSSession $LocalSession -DisableNameChecking

#Create mail user from user account

Write-Host "Creating mail user account for $samaccountname" -ForegroundColor Green
Enable-MailUser -Identity $samaccountname -ExternalEmailAddress $emailaddress -alias $samaccountname -DomainController $domaincontroller | Out-Null

#Logout from local Exchange server

Write-Host "Disconnecting from $exchangeserver" -ForegroundColor Yellow
Remove-PSSession $LocalSession

#Login to Exchange Online

Write-Host Connecting to Exchange Online using $cloudadmin.UserName -ForegroundColor Yellow
$CloudSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $cloudadmin -Authentication Basic -AllowRedirection
Import-PSSession $CloudSession

#Modify parameters for account

Write-Host "Fetching properties from Exchange Online and writing them to the local Active Directory account of $samaccountname" -ForegroundColor Green
$365MboxGUID = get-mailbox -identity $samaccountname | select -ExpandProperty ExchangeGuid
Set-ADUser $samaccountname -replace @{msExchMailboxGuid=$365MboxGUID;targetAddress=$targetaddress;msExchRecipientDisplayType=$displaytype;msExchRecipientTypeDetails=$typedetails;msExchRemoteRecipientType=$recipienttype} -Server $domaincontroller
Set-ADUser $samaccountname -Add @{proxyAddresses="smtp:$targetaddress"} -Server $domaincontroller

#Disconnect from Exchange Online

Write-Host "Disconnecting from Exchange Online" -ForegroundColor Yellow
Remove-PSSession $CloudSession

#Sync to Azure AD
Start-ADSyncSyncCycle -PolicyType Delta
Write-Host "Active Directory Connect Sync Cycle has been started" -ForegroundColor Green
