function Move-EwsItem {
    [CmdletBinding(DefaultParameterSetName = 'All')]
    param (
        [parameter(Mandatory, ParameterSetName = 'All')]
        [parameter(Mandatory, ParameterSetName = 'DateFilter')]
        [ValidateNotNullOrEmpty()]
        [string]$MailboxAddress,

        [parameter(Mandatory, ParameterSetName = 'All')]
        [parameter(Mandatory, ParameterSetName = 'DateFilter')]
        [ValidateNotNullOrEmpty()]
        $SourceFolder,

        [parameter(Mandatory, ParameterSetName = 'All')]
        [parameter(Mandatory, ParameterSetName = 'DateFilter')]
        [ValidateNotNullOrEmpty()]
        $TargetFolder,

        [parameter(Mandatory, ParameterSetName = 'DateFilter')]
        [datetime]$StartDate,

        [parameter(Mandatory, ParameterSetName = 'DateFilter')]
        [datetime]$EndDate,

        [parameter(ParameterSetName = 'All')]
        [parameter(ParameterSetName = 'DateFilter')]
        [boolean]$TestMode = $true

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

    # If StartDate and EndDate are used, create the Search Filter collection
    if ($PSCmdlet.ParameterSetName -eq 'DateFilter') {
        $SearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And)
        $startDateFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThan([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived, $StartDate)
        $endDateFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThan([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived, $EndDate)
        $SearchFilter.Add($startDateFilter)
        $SearchFilter.Add($endDateFilter)
    }

    do {
        if ($PSCmdlet.ParameterSetName -eq 'DateFilter') {
            $FindItemResults = $service.FindItems($SourceFolder.Id, $SearchFilter, $ItemView)
        }
        else {
            $FindItemResults = $service.FindItems($SourceFolder.Id, $ItemView)
        }

        $i = 1
        foreach ($Item in $FindItemResults.Items) {
            if ($TestMode -eq $true) {
                Write-Progress -Activity "[LIST ONLY]] $($SourceFolder.DisplayName) to $($TargetFolder.DisplayName)" -Status "$i of $($FindItemResults.TotalCount)" -PercentComplete (($i / $FindItemResults.TotalCount) * 100)
                $Item | Select-Object DateTimeReceived, Sender, Subject
            }
            elseif ($TestMode -eq $false) {
                $Message = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($service, $Item.Id)
                $Message.Move($TargetFolder.Id) > $null
                Write-Progress -Activity "Moving messages from $($SourceFolder.DisplayName) to $($TargetFolder.DisplayName)" -Status "$i of $($FindItemResults.TotalCount)" -PercentComplete (($i / $FindItemResults.TotalCount) * 100)
            }
            $i++
        }
        $ItemView.offset += $FindItemResults.Items.Count
    } while ($FindItemResults.MoreAvailable -eq $true)
}