function Get-EwsMessageIdCache {
    param (
        [Microsoft.Exchange.WebServices.Data.ExchangeService]$Service,
        [Microsoft.Exchange.WebServices.Data.Folder]$TargetFolder
    )

    $itemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)
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