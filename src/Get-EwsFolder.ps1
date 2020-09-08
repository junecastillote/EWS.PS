Function Get-EwsFolder {
    [CmdletBinding(DefaultParameterSetName='Default')]
    param (
        [parameter(Mandatory,ParameterSetName='Default')]
        [parameter(Mandatory,ParameterSetName='byFolderName')]
        [parameter(Mandatory,ParameterSetName='byFolderID')]
        [ValidateNotNullOrEmpty()]
        $Token,

        [parameter(Mandatory,ParameterSetName='Default')]
        [parameter(Mandatory,ParameterSetName='byFolderName')]
        [parameter(Mandatory,ParameterSetName='byFolderID')]
        [ValidateNotNullOrEmpty()]
        [string]$MailboxAddress,

        [parameter(Mandatory,ParameterSetName='Default')]
        [parameter(Mandatory,ParameterSetName='byFolderName')]
        [parameter(Mandatory,ParameterSetName='byFolderID')]
        [ValidateSet('Primary', 'Archive')]
        [string]$MailboxType,

        [parameter(Mandatory,ParameterSetName='byFolderName')]
        [ValidateNotNullOrEmpty()]
        [string]$FolderName,

        [parameter(Mandatory,ParameterSetName='byFolderID')]
        [ValidateNotNullOrEmpty()]
        [string]$FolderID
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

    ## If the target mailbox is Primary, get the Primary Mailbox Folders
    if ($MailboxType -eq 'Primary') {
        $ConnectToMailboxRootFolders = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.Webservices.Data.WellKnownFolderName]::MsgFolderRoot, $MailboxAddress)
    }
    ## If the target mailbox is Archive, get the Archive Mailbox Folders
    elseif ($MailboxType -eq 'Archive') {
        $ConnectToMailboxRootFolders = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.Webservices.Data.WellKnownFolderName]::ArchiveMsgFolderRoot, $MailboxAddress)
    }

    ## Bind the mailbox folders
    $EWSParentFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $ConnectToMailboxRootFolders)

    ## Create the FolderView
    $FolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(100)
    $FolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep

    ## If -FolderName is specified, look for the said folder using its DisplayName
    if ($FolderName) {
        $SearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $FolderName)
        $MailboxFolderList = $EWSParentFolder.FindFolders($SearchFilter, $FolderView)
    }
    ## If -FolderID is specified, get all folders and filter to look for the said folder using its ID
    elseif ($FolderID) {
        $MailboxFolderList = ($EWSParentFolder.FindFolders($FolderView) | Where-Object {$_.ID -eq $FolderID})
    }
    ## If -FolderName and -FolderID are NOT specified, get ALL folders
    else {
        $MailboxFolderList = $EWSParentFolder.FindFolders($FolderView)
    }

    return $MailboxFolderList
}

