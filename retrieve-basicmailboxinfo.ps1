# add the Exchange powershell snapin for on-premises Exchange 2010 environment; may be different for Exchange 2013/2016

Add-PSSnapin Microsoft.Exchange.Management.Powershell.SnapIn

# First set the value of the mailboxes variable to scope the information we want to gather.  Any args passed to get-mailbox are valid, here it's by OU

$Mailboxes = Get-Mailbox -OrganizationalUnit "OU=Exchange Resources,DC=fabrikam,DC=com"
$outputcsv = basic_mailboxinfo.csv

# Loop through the selected mailboxes grabbed by Get-Mailbox, then define variables that contain the information from the commands you want to populate in a new object

$Results = ForEach($Mailbox in $Mailboxes) {
	$Stats = $Mailbox | 
	Get-MailboxStatistics

	$Perms = $Mailbox | 
	Get-MailboxPermission | 
	Where-Object {
	    ($_.AccessRights -like “*FullAccess*”) -and 
		(-not $_.IsInherited) -and
        (-not $_.Deny) -and 
		($_.User -notlike "S-1-5-21*") -and 
		($_.User -notlike “*NT AUTHORITY\SELF*”)
	} |

# Expand the selected property so that all users are populated properly in the output column

	Select-Object -ExpandProperty User

# Build the new object by compiling the various properties from each of the different commands

	New-Object -TypeName PSObject -Property @{
		DisplayName = $Stats.DisplayName
		ItemCount = $Stats.ItemCount
		TotalItemSize = $Stats.TotalItemSize
        LastLoggedOnUserAccount = $Stats.LastLoggedOnUserAccount
		LastLogonTime = $Stats.LastLogonTime
		UserPermissions = $Perms -Join "; "
        ForwardingAddress = $mailbox.ForwardingAddress
        ForwardingSMTPAddress = $mailbox.ForwardingSMTPAddress
        DeliverToMailboxAndForward =$mailbox.DeliverToMailboxAndForward
	}
}

# Here we can just export the results to the screen or pipe them out to CSV with the usual parameters. Grab whatever properties you want and the order here sets the output order in the columns sent to the CSV

$Results | select displayname, itemcount, totalitemsize, lastloggedonuseraccount, lastlogontime, userpermissions, forwardingaddress, forwardingSMTPaddress, delivertomailboxandforward | export-csv $outputcsv -NoTypeInformation
