Function Get-EwsPsMailboxFolder {
    [CmdletBinding()]
    param (
        [parameter(Mandatory)]
        $Token,

        [parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$MailboxAddress,

        [parameter(Mandatory)]
        [ValidateSet('Primary', 'Archive')]
        [string]$MailboxType,

        [parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$FolderName,

        [parameter()]
        [string]$EwsDLL = 'C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll'
    )

    Import-Module -Name $EwsDLL -ErrorAction Stop
    $Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService -ArgumentList 'Exchange2013_SP1'
    $Service.Url = 'https://outlook.office365.com/EWS/Exchange.asmx'
    $Service.UseDefaultCredentials = $false
    #$Service.Credentials = [Microsoft.Exchange.WebServices.Data.OAuthCredentials]$AccessToken
    $Service.Credentials = [Microsoft.Exchange.WebServices.Data.OAuthCredentials]::new($Token.AccessToken)
    $service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxAddress);

    if ($MailboxType -eq 'Primary') {
        $ConnectToMailboxRootFolders = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.Webservices.Data.WellKnownFolderName]::MsgFolderRoot, $MailboxAddress)
    }
    elseif ($MailboxType -eq 'Archive') {
        $ConnectToMailboxRootFolders = New-Object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.Webservices.Data.WellKnownFolderName]::ArchiveMsgFolderRoot, $MailboxAddress)
    }

    $EWSParentFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service, $ConnectToMailboxRootFolders)
    $FolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(100)
    $FolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep

    if ($FolderName) {
        $SearchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName, $FolderName)
        $MailboxFolderList = $EWSParentFolder.FindFolders($SearchFilter, $FolderView)
        return $MailboxFolderList
    }
    else {
        $MailboxFolderList = $EWSParentFolder.FindFolders($FolderView)
        return $MailboxFolderList
    }
}

Function Move-EwsPsMessageToFolder {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$SourceFolderID,

        [Parameter(Mandatory)]
        [ValidateSet('Primary','Archive')]
        [string]$SourceMailboxType,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$TargetFolderID,

        [Parameter(Mandatory)]
        [ValidateSet('Primary','Archive')]
        [string]$TargetMailboxType
    )
}