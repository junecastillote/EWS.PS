$ClientID = '94ba46ef-478c-4e8e-98e5-99a99a833db7'
$Secret = 'ajC~M30veNTSm932.BzpXBb5ASb.G3V~t-'
#$TenantID = 'b1f4ac95-b2d2-41db-ba4c-5627a94ad435'
$TenantID = 'poshlab.ga'
$msalParams = @{
    ClientId = $ClientID
    TenantId = $TenantID
    Scopes   = "https://outlook.office.com/.default"
	ClientSecret = (ConvertTo-SecureString $Secret -AsPlainText -Force)
	#Silent = $true
}
$token = Get-MsalToken @msalParams


$mailbox = 'june@poshlab.ga'
. .\src\functions.ps1
$pFolders = Get-EwsPsMailboxFolder -Token $token -MailboxAddress $mailbox -MailboxType Primary
$aFolders = Get-EwsPsMailboxFolder -Token $token -MailboxAddress $mailbox -MailboxType Archive

$SourceFolderID = 'AQMkADRmZTI3MWRlLWY2NTEtNDdlYS04MGE0LTNmODZhNzFhNTMzAGIALgAAAw+T9fhAm8pLhZSATLiuFvEBAEqUOJ1EAgdClXDW7Gfy3cMAAAIBDAAAAA=='

Move-EwsPsMessageToFolder -Token $token -MailboxAddress $mailbox -SourceFolderID $SourceFolderID