
#poshlab.ga
$ClientID = '94ba46ef-478c-4e8e-98e5-99a99a833db7'
$Secret = 'ajC~M30veNTSm932.BzpXBb5ASb.G3V~t-'
$TenantID = 'b1f4ac95-b2d2-41db-ba4c-5627a94ad435'

#downergroup.com
# $ClientID = 'b70d3ea8-1691-4d57-9e8b-740f05bd4855'
# $Secret = 'jIUGFCdtga4iu/kBn7F06JbyhC8g7/0U05m/RC5IkA8='
# $TenantID = 'downergroup.com'

$msalParams = @{
    ClientId = $ClientID
    TenantId = $TenantID
    Scopes   = "https://outlook.office.com/.default"
	ClientSecret = (ConvertTo-SecureString $Secret -AsPlainText -Force)
}
$token = Get-MsalToken @msalParams

Remove-Module Ews.Ps.Move.Email -ErrorAction SilentlyContinue
Import-Module C:\GitHub\EWS.PS.Move.Email\Ews.Ps.Move.Email.psd1


$mailbox = 'june@poshlab.ga'
$pFolders = Get-EwsFolder -Token $token -MailboxAddress $mailbox -MailboxType Primary -Folder
$aFolders = Get-EwsFolder -Token $token -MailboxAddress $mailbox -MailboxType Archive

# $pFolders = Get-EwsPsMailboxFolder -Token $token -MailboxAddress $mailbox -MailboxType Primary
# $aFolders = Get-EwsPsMailboxFolder -Token $token -MailboxAddress $mailbox -MailboxType Archive

$SourceFolderID = Get-EwsFolder -Token $token -MailboxAddress $mailbox -MailboxType Primary -FolderName Inbox
$TargetFolderID = Get-EwsFolder -Token $token -MailboxAddress $mailbox -MailboxType Archive -FolderName Archive

Move-EwsItem -Token $token -MailboxAddress $mailbox -SourceFolderID $SourceFolderID -TargetFolderID $TargetFolderID -StartDate (Get-Date).AddDays(-3)
#Move-EwsItem -Token $token -MailboxAddress $mailbox -SourceFolderID $TargetFolderID -TargetFolderID $SourceFolderID -StartDate (Get-Date).AddDays(-3) -EndDate (Get-Date)


# $folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$mailbox)
# $Inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
# $mailitems = $inbox.FindItems(1000)
# $mailitems | ForEach {$_.Load()}
# $mailitems | Select Sender, InternetMessageID, LastModifiedTime

# $ItemView = new-object -TypeName Microsoft.Exchange.WebServices.Data.ItemView -ArgumentList (1000)