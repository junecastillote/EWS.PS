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

    ## Check registry if EWS Managed API is installed
    $EwsDLL = (($(Get-ItemProperty -ErrorAction SilentlyContinue -Path Registry::$(Get-ChildItem -ErrorAction SilentlyContinue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Exchange\Web Services' | Sort-Object Name -Descending | Select-Object -First 1 -ExpandProperty Name)).'Install Directory') + "Microsoft.Exchange.WebServices.dll")
    if (!($EwsDLL) -or !(Test-Path $EwsDLL)) {
        Write-Error "The EWS Managed API is not found. Go to https://www.microsoft.com/en-us/download/details.aspx?id=42951 to download and install."
        Return $null
    }

    ## Import the EWS Managed API Module
    Import-Module -Name $EwsDLL -ErrorAction Stop

    ## Create the EWS Object
    $Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService -ArgumentList 'Exchange2013_SP1'

    ## Exchange Online EWS URL
    $Service.Url = 'https://outlook.office365.com/EWS/Exchange.asmx'

    ## EWS Authentication
    $Service.UseDefaultCredentials = $false
    $Service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials -ArgumentList ($Token.AccessToken)

    ## Who are we impersonating?
    $service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxAddress);

    ## We're impersonating, so we need to anchor to the target mailbox
    ## https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/impersonation-and-ews-in-exchange#performance-considerations-for-ews-impersonation
    $service.HttpHeaders.Add('X-AnchorMailbox', $MailboxAddress)

    $ItemView = new-object -TypeName Microsoft.Exchange.WebServices.Data.ItemView -ArgumentList (1000)

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
                $Item | Select-Object DateTimeReceived,Sender,Subject
            }
            elseif ($TestMode -eq $false) {
                $Message = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($service, $Item.Id)
                $Message.Move($TargetFolder.Id) > $null
                Write-Progress -Activity "Moving messages from $($SourceFolder.DisplayName) to $($TargetFolder.DisplayName)" -Status "$i of $($FindItemResults.TotalCount)" -PercentComplete (($i / $FindItemResults.TotalCount) * 100)
            }

            Write-Progress -Activity "Moving messages from $($SourceFolder.DisplayName) to $($TargetFolder.DisplayName)" -Status "$i of $($FindItemResults.TotalCount)" -PercentComplete (($i / $FindItemResults.TotalCount) * 100)
            $i++
        }
        $ItemView.offset += $FindItemResults.Items.Count
    } while ($FindItemResults.MoreAvailable -eq $true)
}