# add the Exchange powershell snapin for on-premises Exchange 2010 environment; may be different for Exchange 2013/2016

Add-PSSnapin Microsoft.Exchange.Management.Powershell.SnapIn

# First we set the value of the mailboxes variable to scope the information we want to gather.  Any switches usually passed to get-mailbox are valid to use.  Example variable uses OU filtering.

$Mailboxes = Get-Mailbox -OrganizationalUnit "OU=Migration,OU=Exchange Resources,DC=fabrikam,DC=com"
$outputcsv = bookingpolicy_calendarinfo.csv

# This action grabs the full access permissions and the important booking policy properties on room/equipment mailboxes that don't get copied over during a mailbox migration to office 365

$Results = ForEach($Mailbox in $Mailboxes) {
	$Stats = $Mailbox | 
	Get-MailboxStatistics
    
    $Calendar = $Mailbox |
    Get-CalendarProcessing

	$Perms = $Mailbox | 
	Get-MailboxPermission | 
	Where-Object {
	    ($_.AccessRights -like “*FullAccess*”) -and 
		(-not $_.IsInherited) -and
        (-not $_.Deny) -and 
		($_.User -notlike "S-1-5-21*") -and 
		($_.User -notlike “*NT AUTHORITY\SELF*”)
	} |
	Select-Object -ExpandProperty User

	New-Object -TypeName PSObject -Property @{
		DisplayName = $Stats.DisplayName
		ItemCount = $Stats.ItemCount
		TotalItemSize = $Stats.TotalItemSize
        LastLoggedOnUserAccount = $Stats.LastLoggedOnUserAccount
		LastLogonTime = $Stats.LastLogonTime
		UserPermissions = $Perms -Join "; "
        ResourceDelegates =  [string]::join(“;”, ($Calendar.ResourceDelegates)) 
        BookInPolicy = [string]::join(“;”, ($Calendar.BookInPolicy))
        AllBookInPolicy = [string]::join(“;”, ($Calendar.AllBookInPolicy))
        RequestInPolicy = [string]::join(“;”, ($Calendar.RequestInPolicy))
        AllRequestInPolicy = [string]::join(“;”, ($Calendar.AllRequestInPolicy))
        RequestOutOfPolicy = [string]::join(“;”, ($Calendar.RequestOutOfPolicy))
        AllRequestOutOfPolicy = [string]::join(“;”, ($Calendar.AllRequestOutOfPolicy))      
	}
} 

# Here we can just export the results to the screen or pipe them out to CSV with the usual parameters

$Results | select displayname,itemcount, totalitemsize, lastloggedonuseraccount, lastlogontime, userpermissions, resourcedelegates, bookinpolicy, allbookinpolicy, requestinpolicy, allrequestinpolicy, requestoutofpolicy, allrequestoutofpolicy | export-csv $outputcsv -NoTypeInformation
