$clientId = '007f848b-eb76-48df-b512-351d7cdc5cfd'
$tenantID = '900ac657-a9b3-4335-9ea8-e9bdaeaa1950'
$certificateThumbprint = 'EFB097C025AA503FCF810D3FFB7C2684749BF496'
$clientSecret = 'FU.8Q~9J.4msG7reZjrr.FzOneEHSJCTGWY.ncsi'

Import-Module .\Ews.Ps.psd1 -Force

connect-Ews -TenantId $tenantID -ClientId $clientId -CertificateThumbprint $certificateThumbprint -Verbose
# connect-Ews -TenantId $tenantID -ClientId $clientId -ClientSecret $clientSecret -Verbose

$mailbox = 'june@poshlab.xyz'

$sourceFolder = Get-EwsMailboxFolder -MailboxAddress $mailbox -MailboxType Archive -Verbose -FolderClass Email | Where-Object { $_.DisplayName -eq 'Archive' }
$targetFolder = Get-EwsMailboxFolder -MailboxAddress $mailbox -MailboxType Primary -Verbose -FolderClass Email | Where-Object { $_.DisplayName -eq 'Archive' }

$items = $sourceFolder | ForEach-Object { Get-EwsMailboxMessage -Folder $_ }

$copy_result = Copy-EwsMailboxMessage -MessageItem $items -TargetFolder $targetFolder -Verbose -TestMode:$false -Move:$true