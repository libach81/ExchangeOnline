#---------------------------
# 
#Script to create Remote Mailbox object from local Active Directory
#
# This can be used if an Online mailbox is provisioned directly from a local AD account, and not using either Enable-RemoteMailbox or the Hybrid migration wizard
# It will modify the AD object to become a Remote Mailbox object seen from the perspective of the local Exchange server.
# 
# It is recommended to run on the AD Connect server
#
# The script must not be run on an Exchange server, due to some cmdlets being the same betweens Exchane On-premise and Exchange Online.
#
#---------------------------

#Input parameters for running script

$cloudadmin = "admin@tenant.onmicrosoft.com"
$localadmin = "DOMAIN\admin"
$exchangeserver = "fqdn"
$domaincontroller = "fqdn"
$samaccountname = "SAM Account name"
$targetaddress = ($samaccountname + "@tenant.mail.onmicrosoft.com")
$emailaddress = "thqthemail.com"

#fixed variables to change
$displaytype = "-2147483642"
$typedetails = "2147483648"
$recipienttype = "4"


#Create mail user from user account

Write-Host "Creating mail user account for $samaccountname" -ForegroundColor Green
Invoke-Command -ComputerName $exchangeserver -Credential $localadmin -ScriptBlock {New-MailUser -Identity $samaccountname -ExternalEmailAddress $emailaddress -alias $samaccountname -DomainController $domaincontroller}
Write-Host "Initiating Active Directory replication" -ForegroundColor Green
Invoke-Command -ComputerName $domaincontroller -Credential $localadmin -ScriptBlock {Get-ADDomainController -Filter *).Name | Foreach-Object {repadmin /syncall $_ (Get-ADDomain).DistinguishedName /e /A}
Start-AdSyncCycle -Policytype Delta
Write-Host "Active Directory Connect Sync Cycle has been started. Waiting for it to complete" -ForegroundColor Green
Start-Sleep 60

#Login to Exchange Online
Write-Host "Login on to Exchange online using $cloudadmin" -ForegroundColor Green
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $cloudadmin -Authentication Basic -AllowRedirection
Import-PSSession $Session

#Modify parameters for account

Write-Host "Fetching properties from Exchange Online and writing them to the local Active Directory account of $samaccountname" -ForegroundColor Green
$365MboxGUID = get-mailbox -identity $samaccountname | select -ExpandProperty ExchangeGuid
Set-ADUser $samaccountname -replace @{msExchMailboxGuid=$365MboxGUID;targetAddress=$targetaddress;msExchRecipientDisplayType=$displaytype;msExchRecipientTypeDetails=$typedetails;msExchRemoteRecipientType=$recipienttype} -Server $domaincontroller
Set-ADUser $samaccountname -Add @{proxyAddresses=smtp: + $targetaddress} -Server $domaincontroller
