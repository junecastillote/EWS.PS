function Get-EwsMessageIdCache {
    [CmdletBinding()]
    param (
        [parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        $Folder
    )

    $MailboxAddress = @($Folder)[0].MailboxAddress

    if (!($Token = Get-EwsAccessToken)) {
        Write-Error "EWS is not connected. Run the Connect-Ews command first."
        return $null
    }

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

    $ItemView = New-Object -TypeName Microsoft.Exchange.WebServices.Data.ItemView -ArgumentList (1000)
    $itemView.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet(
        [Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly,
        [Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::InternetMessageId
    )

    $cache = New-Object 'System.Collections.Generic.HashSet[string]'
    do {
        $results = $Service.FindItems($TargetFolder.Id, $itemView)
        foreach ($item in $results.Items) {
            if ($item.InternetMessageId) {
                $null = $cache.Add($item.InternetMessageId)
            }
        }
        $itemView.Offset += $results.Items.Count
    } while ($results.MoreAvailable)

    return $cache
}