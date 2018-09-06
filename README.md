# PowerShell-Scripts-for-O365-Admin

##POWERSHELL O365 COMMANDS

##------------------------------------------------------------------MAILBOXES-----------------------------------------------------------------------------------


## Disable Clutter for an individual Mailbox 
    Set-Clutter -Identity UPN -Enable $False

## Add SMTP Forwarding Address 
    Set-Mailbox EMAILADDRESS -ForwardingSmtpAddress SMTPADDRESS -DeliverToMailboxAndForward $false

##Remove SMTP Forwarding Address 
    Set-Mailbox EMAILADDRESS -ForwardingSmtpAddress $null

##Remove SMTP Forwarding Address in bulk 
    Import-Csv "C:\o365\removeforwarding.csv" |foreach {Set-Mailbox $_.upn -ForwardingSmtpAddress $null -DeliverToMailboxAndForward $False}

##Email Individual Mailbox Archive 
    Enable-Mailbox -Identity UPN –archive

##Set Individual Mailbox Retention Policy 
    Set-Mailbox EMAILADDRESS -retentionpolicy "NAME OF POLICY"
  
##Exporting Distribution Group Member List to CSV file (Can be run in Cogmotive as well) 
    Get-DistributionGroupMember {Name of Group} | Sort -Property DisplayName | Select DisplayName, Alias, Department | Export-CSV C:DGmemberlist.csv
    Get-DistributionGroupMember NAME | Sort -Property DisplayName | Select emailaddresses >C:\Users\mburke\Desktop\Temp\DistNames.txt
    Get-DistributionGroupMember -identity <LISTNAME> c:\temp\<LISTNAME>.txt

##Add SMB Full Access with Automapping Disabled 
    Add-MailboxPermission SMBEMAILADDRESS -user UPN -AccessRights FullAccess -Automapping $False

##Add Send As permissions to a Mailbox 
    Add-RecipientPermission SMBEMAILADDRESS -trustee UPN -accessrights SendAs

##Get Distribution Group Information 
    Get-DistributionGroup EMAILADDRESS | fl 

##Add Mailbox Size (Bulk)
    Get-Content "C:\o365\MailboxSize.csv" | Set-Mailbox -ProhibitSendReceiveQuota 107374184400 -ProhibitSendQuota 106300440576 -IssueWarningQuota 105226698752 

##Confirm Mailbox Size (Bulk)
    Get-Content "C:\o365\MailboxSize.csv" | Get-Mailbox | Select ProhibitSendReceiveQuota,ProhibitSendQuota,IssueWarningQuota | Export-Csv C:\o365\MailboxSizeResults.csv 

##Getting all users from a domain
    Get-Mailbox -ResultSize unlimited | where {$_.primarysmtpaddress -like "*@smpi74.fr"} | select name, primarysmtpaddress

##Checking PrimarySMTP for multiple users
    Get-Content "C:\o365\PrimarySMTP.csv" | Get-Mailbox | Select PrimarySmtpAddress  

##Mass Adding External Contacts
    PS C:\Windows\system32> Import-Csv "C:\o365\contactstest.csv" | ForEach{New-MailContact -DisplayName $_.DisplayName -FirstName $_.FirstName -LastName $_.LastName -ExternalEmailAddress $_.EmailAddress -Alias $_.Alias -Name $_.Name 

##Message Trace Location
    Search-Mailbox EMAILADDRESS -SearchDumpster -SearchQuery 'subject:"website*"' -TargetMailbox Jacob.thoreson-cloud@itwo365.com -targetfolder inbox

##Purge Mailbox Data
    Search-Mailbox -Identity "<MailboxOrMailUserIdParameter>" -DeleteContent -force

##Set OOO
    Set-MailboxAutoReplyConfiguration pschmitt@youremail.com –AutoReplyState Enabled –ExternalMessage “” –InternalMessage “”
    ##Scheduled –StartTime “1/8/2013” –EndTime “1/15/2013” for scheduling the OOO
##Turn OFF OOO
    Set-MailboxAutoReplyConfiguration username –AutoReplyState Disabled –ExternalMessage $null –InternalMessage $null
##Check Inbox Rules for mailbox
    Get-InboxRule -Mailbox Joe@Contoso.com | fl
 
##Adding Delegate to managers calendar
    Add-MailboxFolderPermission -Identity ayla@contoso.com:\Calendar -User laura@contoso.com -AccessRights Editor -SharingPermissionFlags Delegate,CanViewPrivateItems

##Bulk Alias Add to DG
    Import-Csv C:\o365\Book1.csv | foreach {Set-DistributionGroup cylance@itw.com -EmailAddress @{add=$_.alias}}  

##---------------------------------------------------------SPAM/EOP----------------------------------------------------------------------



##Get Junk Email Configuration 
    Get-MailboxJunkEmailConfiguration EMAILADDRESS
##Check the blocked senders and safe senders at the client level, use the following commands 
    (Get-MailboxJunkEmailConfiguration EMAIL ADDRESS).blockedsendersanddomains
    (Get-MailboxJunkEmailConfiguration EMAIL ADDRESS).trustedsendersanddomains
 
##Add Address or Domain to Mailbox Blocked Senders/Domain list 
    Set-MailboxJunkEmailConfiguration EMAILADDRESS -TrustedSendersAndDomains @{Add=""}
    Set-MailboxJunkEmailConfiguration EMAILADDRESS -BlockedSendersAndDomains @{Add=""}
##Remove address or domain from Blocked (or trusted) list 
    Set-MailboxJunkEmailConfiguration EMAILADDRESS -TrustedSendersAndDomains @{remove=""}
    Set-MailboxJunkEmailConfiguration EMAILADDRESS -BlockedSendersAndDomains @{remove=""} 

##Perimeter Message Trace (Does not show up in Message Trace, sender will receive NDR)
    Get-PerimeterMessageTrace -Recipient EMAILADDRESS

#EXPORT MESSAGE TRACE TO CSV
    Get-MessageTrace -SenderAddress “Email Address” -StartDate 2/20/2018 -EndDate 2/22/2018 | Export-Csv C:\Users\jthoreson\Documents\klopezMT.csv 


#Content Search
    New-ComplianceSearchAction -SearchName "Content Search Name" -Purge -PurgeType SoftDelete

#Confirm Completion of Content Search
    Get-ComplianceSearch "Content Search Name" | Select Status

#Export of Message Trace
    $dateEnd = get-date; Get-MessageTrace -SenderAddress arnie.buchanan@buehler.com -StartDate 08/09/2018 -EndDate $dateEnd -PageSize 5000 | Export-Csv C:\o365\buchanon.csv


##-----------------------------------------------------------Shared Mailbox---------------------------------------------------------------------------------------



#Send a copy of Sent Item into SMB Sent folder and user sent folder
##Enable the feature
##For emails Sent As the shared mailbox:
    set-mailbox <mailbox name> -MessageCopyForSentAsEnabled $True
##For emails Sent On Behalf of the shared mailbox:
    set-mailbox <mailbox name> -MessageCopyForSendOnBehalfEnabled $True
##Disable the feature
##For messages Sent As the shared mailbox:
    set-mailbox <mailbox name> -MessageCopyForSentAsEnabled $False
##For emails Sent On Behalf of the shared mailbox:
    set-mailbox <mailbox name> -MessageCopyForSendOnBehalfEnabled $False



##---------------------------------------------CALENDARS---------------------------------------------------------------- 

##Finding who has what permissions to a Calendar
    Get-MailboxFolderPermission EMAILADDRESS:\calendar | fl
 
##Add Full Access permissions to a calendar(Change the Command to Set-MailboxFolderPermission if the permission for the user already exist.)
    Add-MailboxFolderPermission EMAILADDRESS:\calendar –User UPN –AccessRights PERMISSIONS

##Add Bulk Permissions to a Calendar after creating an Array of the UPN's in a Variable
    ForEach ($Var in $UPNVar) {Add-MailboxFolderPermission CALENDAREMAILADDRESS:\calendar -User $Var -AccessRights PERMISSION}

#Owner:	Allows full rights to the mailbox (Calendar or Folder) , including assigning permissions; it is recommended not to assign this role to anyone.
#Publishing Editor:	Create, read, edit, and delete all items; create subfolders.
#Editor:	Create, read, edit, and delete all items.
#Publishing Author:	Create and read items; create subfolders; edit and delete items created by the user.
#Author:	Create and read items; edit and delete items they’ve created.
#Nonediting Author:	Create and read items; delete items created by the user.
#Reviewer:	Read items.
#Limited Details:	Read Title and Time of items.
#Contributor:	Create items.
#Free/Busy time, subject, location:	View the time, subject, and location of the appointment or meeting on your calendar.
#Free/Busy time:	Shows only as Free or Busy on your calendar. No details are provided.
#None:	No permissions are set for the selected user on the specified calendar or folder.

##MassCreateRooms
    Import-Csv "C:\o365\MassRoomCreate.csv" | ForEach-Object {New-Mailbox -Name $_.DisplayName -Alias $_.Alias -PrimarySmtpAddress $_.PrimarySMTPAddress -Room}

##Mass CreateEquipmentCalendars
    Import-Csv "C:\o365\MassEquipmentCreate.csv" | ForEach-Object {New-Mailbox -Name $_.DisplayName -Alias $_.Alias -PrimarySmtpAddress $_.PrimarySMTPAddress -Equipment}

##Add Conference room to “Room List”
    Add-DistributionGroupMember –Identity "Name" -Member cr-TestRoom@itw.com 

##Adding Rooms to Rooms (keep subject the same)
    Set-CalendarProcessing -Identity <RESOURCEMAILBOX> -DeleteSubject $False -AddOrganizerToSubject $False

##ROOM LIST
##Making New Room List
New-DistributionGroup -Name "ex" -DisplayName "ex"  –RoomList 

## Room to Room List
Add-DistributionGroupMember –Identity "Name of Room List" –Member "Name of Room Mailbox"


##------------------------------------------------Recover Mailboxes---------------------------------------------------------------

##1.	You need to get the ExchangeGUID from the soft-deleted mailbox. The easiest way is via the Exchange Online PowerShell cmdlet:
Get-Mailbox name@example.com -SoftDeletedMailbox | Format-List Name,ExchangeGuid
 
##2.	Once you have the ExchangeGUID, you can use the cmdlet:
New-MailboxRestoreRequest -SourceMailbox <insert-ExchangeGUID-here> -TargetMailbox <new-mailbox-name> -TargetRootFolder “SharedMailboxData” -AllowLegacyDNMismatch
 ##The “-TargetRootFolder” parameter is optional; it allows you to specify a folder to restore the contents into.

##Check status of restore
Get-MailboxRestoreRequest


##-------------------------------------------------Migration----------------------------------------------------------------------



##Script to use when validating the proxy addresses as the UPN:
    Import-Csv "C:\o365\validateproxy.csv" |foreach {get-mailbox $_.upn |select primarysmtpaddress}
 ##Creating Multiple DG’s
    Import-Csv "C:\o365\massDGcreate.csv" | ForEach{New-DistributionGroup -DisplayName $_.DisplayName -Name $_.Name -PrimarySmtpAddress $_.EmailAddress -Alias $_.Alias -ManagedBy $_.Owner -MemberDepartRestriction Closed -MemberJoinRestriction Closed} 



##-----------------------------------------------Skype--------------------------------------------------------------------------------


##Sign in Skype Module

Import-Module SkypeOnlineConnector
$sfboSession = New-CsOnlineSession -Credential $credential
Import-PSSession $sfboSession

##Change user to static Conference ID (number must be 7 digits)
Set-CsOnlineDialInConferencingUser -Identity "Amos Marble"  -ResetLeaderPIN 8271964



####--------------------------------Service Accounts---------------------------------------------------------------------------------------------

##Setting up Service account as Resource Room
$password = ConvertTo-SecureString -String <createpassword> -AsPlainText -Force
Set-Mailbox -Identity SmtpTest@itw.com -EnableRoomMailboxAccount $true -RoomMailboxPassword $password 


##Removing auto deleting of emails
Set-CalendarProcessing “room mailbox” -AutomateProcessing AutoUpdate
