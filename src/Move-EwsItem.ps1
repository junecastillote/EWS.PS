Function Move-EwsItem {
    [CmdletBinding(DefaultParameterSetName = 'All')]
    param (
        [parameter(Mandatory, ParameterSetName = 'All')]
        [parameter(Mandatory, ParameterSetName = 'DateFilter')]
        [ValidateNotNullOrEmpty()]
        $Token,

        [parameter(Mandatory, ParameterSetName = 'All')]
        [parameter(Mandatory, ParameterSetName = 'DateFilter')]
        [ValidateNotNullOrEmpty()]
        [string]$MailboxAddress,

        [parameter(Mandatory, ParameterSetName = 'All')]
        [parameter(Mandatory, ParameterSetName = 'DateFilter')]
        [ValidateNotNullOrEmpty()]
        $SourceFolderID,

        [parameter(Mandatory, ParameterSetName = 'All')]
        [parameter(Mandatory, ParameterSetName = 'DateFilter')]
        [ValidateNotNullOrEmpty()]
        $TargetFolderID,

        [parameter(Mandatory, ParameterSetName = 'DateFilter')]
        [datetime]$StartDate,

        [parameter(Mandatory, ParameterSetName = 'DateFilter')]
        [datetime]$EndDate,

        [parameter()]
        [string]$EwsDLL = 'C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll'
    )
    Import-Module -Name $EwsDLL -ErrorAction Stop
    $Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService -ArgumentList 'Exchange2013_SP1'
    $Service.Url = 'https://outlook.office365.com/EWS/Exchange.asmx'
    $Service.UseDefaultCredentials = $false
    $Service.Credentials = [Microsoft.Exchange.WebServices.Data.OAuthCredentials]::new($Token.AccessToken)
    $service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxAddress);
    $service.HttpHeaders.Add('X-AnchorMailbox', $MailboxAddress)

    $ItemView = new-object -TypeName Microsoft.Exchange.WebServices.Data.ItemView -ArgumentList (1000)

    if ($PSCmdlet.ParameterSetName -eq 'DateFilter') {
        $SearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And)
        $startDateFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThan([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived, $StartDate)
        $endDateFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThan([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived, $EndDate)
        $SearchFilter.Add($startDateFilter)
        $SearchFilter.Add($endDateFilter)
    }

    do {
        if ($PSCmdlet.ParameterSetName -eq 'DateFilter') {
            $FindItemResults = $service.FindItems($SourceFolderID.Id, $SearchFilter, $ItemView)
        }
        else {
            $FindItemResults = $service.FindItems($SourceFolderID.Id, $ItemView)
        }

        $i = 1
        foreach ($Item in $FindItemResults.Items) {
            $Message = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($service, $Item.Id)
            $Message.Move($TargetFolderID.Id) > $null

            Write-Progress -Activity "Moving messages from $($SourceFolderID.DisplayName) to $($TargetFolderID.DisplayName)" -Status "$i of $($FindItemResults.TotalCount)" -PercentComplete (($i / $FindItemResults.TotalCount) * 100)
            $i++
        }
        $ItemView.offset += $FindItemResults.Items.Count
    } while ($FindItemResults.MoreAvailable -eq $true)
}