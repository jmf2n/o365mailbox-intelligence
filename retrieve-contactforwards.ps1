# add the Exchange powershell snapin for on-premises Exchange 2010 environment; may be different for Exchange 2013/2016

Add-PSSnapin Microsoft.Exchange.Management.Powershell.SnapIn

# In this script we will gather all forwarded mailboxes in the organization that are referencing contact objects.  Then we will expand the target of the contacts into the actual email addresses they are forwarding to.

$Mailboxes = Get-Mailbox -resultsize unlimited -filter {ForwardingAddress -ne $null} | select alias,*forward*
$outputcsv = forwardedmailboxes.csv

# Grab all of the mailboxes that have a forward set and if it's to a contact, the target of the contact

$Results = ForEach($Mailbox in $Mailboxes) {
$ForwardingAddress = $null
$ContactDN = $null
	if($Mailbox.ForwardingAddress -ne $null){
$Contact = Get-Recipient $Mailbox.ForwardingAddress
$ForwardingAddress = $Contact.ExternalEmailAddress
$ContactDN = $Contact.Identity
}
elseif ($Mailbox.ForwardingSMTPAddress -ne $null){
$ForwardingAddress = $Mailbox.ForwardingSMTPAddress}

	New-Object -TypeName PSObject -Property @{
		SamAccountName = $Mailbox.alias
        ContactDN = $ContactDN
        TargetOfForwardingAddress = $ForwardingAddress
        DeliverToMailboxAndForward =$mailbox.DeliverToMailboxAndForward
	}
}
# Here we can just export the results to the screen or pipe them out to CSV with the usual parameters.

$Results | select samaccountname,contactdn,targetofforwardingaddress,delivertomailboxandforward | export-csv $outputcsv -NoTypeInformation
