function Copy-EwsMailboxMessage {
    [CmdletBinding()]
    param (
        [parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        $MessageItem,

        [parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        $TargetFolder,

        [parameter()]
        [bool]$TestMode = $true,

        [parameter()]
        [bool]$Deduplicate = $true
    )

    if (@($TargetFolder).Count -gt 1) {
        Write-Error "Specify only a single folder object to the -TargetFolder parameter. It does not accept multiple values."
        return $null
    }

    if ($PSVersionTable.PSVersion -gt 7.2) {
        $PSStyle.Progress.View = 'Classic'
    }

    if (!($Token = Get-EwsAccessToken)) {
        Write-Error "EWS is not connected. Run the Connect-Ews command first."
        return $null
    }

    ## Mailbox Address
    $MailboxAddress = @($MessageItem)[0].MailboxAddress

    ## Create the EWS Object
    $Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService -ArgumentList 'Exchange2013_SP1'

    ## Exchange Online EWS URL
    $Service.Url = 'https://outlook.office365.com/EWS/Exchange.asmx'

    ## EWS Authentication
    $Service.UseDefaultCredentials = $false
    $Service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials -ArgumentList $Token

    ## Who are we impersonating?
    $service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxAddress)

    ## We're impersonating, so we need to anchor to the target mailbox
    ## https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/impersonation-and-ews-in-exchange#performance-considerations-for-ews-impersonation
    $service.HttpHeaders.Add('X-AnchorMailbox', $MailboxAddress)

    $totalItems = $MessageItem.Count

    ## === Get target folder items InternetMessageId and cache it. ===
    $targetCache = Get-EwsMessageIdCache -Service $Service -TargetFolder $TargetFolder

    $i = 1

    $result = @()
    foreach ($Item in $MessageItem) {

        if ($TestMode) {
            $prefix = "[TEST MODE] "
        }
        else {
            $prefix = ''
        }

        Write-Progress -Activity "$($prefix)Copy messages from $($item.folder) to $($TargetFolder.DisplayName)" -Status "$i of $($totalItems)" -PercentComplete (($i / $totalItems) * 100)
        $i++

        $outputObject = [PSCustomObject]@{
            MailboxAddress    = $Item.MailboxAddress
            MessageId         = $Item.InternetMessageId
            SourceMailboxType = "$($Item.MailboxType) Mailbox"
            SourcePath        = ($Item.Path.Replace('Top of Information Store\', 'Root\'))
            TargetMailboxType = "$($TargetFolder.MailboxType) Mailbox"
            TargetPath        = ($TargetFolder.Path.Replace('Top of Information Store\', 'Root\'))
            Result            = ''
            Note              = ''
        }

        if ($Deduplicate) {
            if ($targetCache.Contains($Item.InternetMessageId)) {
                Write-Warning "[$($item.InternetMessageId)] already exists in target mailbox folder. Skipping."
                $outputObject.Result = 'Skipped'
                $outputObject.Note = "Message already exists in the target."
                $result += $outputObject
                continue
            }
        }

        if ($TestMode -eq $true) {
            "$($prefix)Copying message: [$($Item.MailboxType)|$($Item.MailboxAddress):$($Item.Path.Replace('Top of Information Store',''))\$($Item.InternetMessageId)] to folder [$($TargetFolder.MailboxType)|$($TargetFolder.MailboxAddress):$($TargetFolder.Path.Replace('Top of Information Store',''))]" | Out-Default
        }
        elseif ($TestMode -eq $false) {
            try {
                $Message = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($service, $Item.Id)
                $Message.Copy($TargetFolder.Id) > $null
                $outputObject.Result = 'Copied'
                $outputObject.Note = ''
            }
            catch {
                $outputObject.Result = 'Failed'
                $outputObject.Note = $_.Exception.Message
            }
        }

        $result += $outputObject
    }
    Write-Progress -Completed
    $result
}