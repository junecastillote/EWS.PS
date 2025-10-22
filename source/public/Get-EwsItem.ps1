function Get-EwsItem {
    [CmdletBinding(DefaultParameterSetName = 'All')]
    param (
        [parameter(Mandatory, ParameterSetName = 'All')]
        [parameter(Mandatory, ParameterSetName = 'DateFilter')]
        [ValidateNotNullOrEmpty()]
        [string]$MailboxAddress,

        [parameter(Mandatory, ParameterSetName = 'All')]
        [parameter(Mandatory, ParameterSetName = 'DateFilter')]
        [ValidateNotNullOrEmpty()]
        [Microsoft.Exchange.WebServices.Data.Folder]$Folder,

        [parameter(Mandatory, ParameterSetName = 'DateFilter')]
        [datetime]$StartDate,

        [parameter(Mandatory, ParameterSetName = 'DateFilter')]
        [datetime]$EndDate
    )

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

    $result = [System.Collections.Generic.List[System.Object]]@()

    # If StartDate and EndDate are used, create the Search Filter collection
    if ($PSCmdlet.ParameterSetName -eq 'DateFilter') {
        $SearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And)
        $startDateFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThan([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived, $StartDate)
        $endDateFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThan([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived, $EndDate)
        $SearchFilter.Add($startDateFilter)
        $SearchFilter.Add($endDateFilter)
        Write-Verbose "Date filter: $StartDate to $EndDate"
    }

    do {
        if ($PSCmdlet.ParameterSetName -eq 'DateFilter') {
            $FindItemResults = $service.FindItems($Folder.Id, $SearchFilter, $ItemView)
        }
        else {
            $FindItemResults = $service.FindItems($Folder.Id, $ItemView)
        }

        if ($FindItemResults) {
            $result.AddRange($FindItemResults)
        }

        $ItemView.offset += $FindItemResults.Items.Count
    } while ($FindItemResults.MoreAvailable -eq $true)

    $result
}