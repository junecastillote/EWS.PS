function Get-EwsFolder {
    [CmdletBinding(DefaultParameterSetName = 'Default')]
    param (
        [parameter(Mandatory, ParameterSetName = 'Default')]
        [parameter(Mandatory, ParameterSetName = 'byFolderName')]
        [parameter(Mandatory, ParameterSetName = 'byFolderID')]
        [ValidateNotNullOrEmpty()]
        [string]$MailboxAddress,

        [parameter(ParameterSetName = 'Default')]
        [parameter(ParameterSetName = 'byFolderName')]
        [parameter(ParameterSetName = 'byFolderID')]
        [ValidateSet('Primary', 'Archive')]
        [string]$MailboxType = 'Primary',

        [parameter(Mandatory, ParameterSetName = 'byFolderName')]
        [ValidateNotNullOrEmpty()]
        [string]$FolderName,

        [parameter(Mandatory, ParameterSetName = 'byFolderID')]
        [ValidateNotNullOrEmpty()]
        [string]$FolderID
    )

    ## Helper: for -FolderName/-FolderID mode (cached recursive bind)
    function Get-FolderPath_Recursive {
        param (
            [Microsoft.Exchange.WebServices.Data.Folder]$Folder
        )

        if ($FolderCache.ContainsKey($Folder.Id.UniqueId) -and $FolderCache[$Folder.Id.UniqueId].Path) {
            return $FolderCache[$Folder.Id.UniqueId].Path
        }

        $parts = @($Folder.DisplayName)
        $current = $Folder

        while ($current.ParentFolderId -and $current.ParentFolderId.UniqueId -ne $EWSParentFolder.Id.UniqueId) {
            $parentId = $current.ParentFolderId.UniqueId

            if ($FolderCache.ContainsKey($parentId) -and $FolderCache[$parentId].Path) {
                $parts += $FolderCache[$parentId].Path
                break
            }

            try {
                $parent = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service, $current.ParentFolderId)
                if ($null -eq $parent) { break }

                $FolderCache[$parent.Id.UniqueId] = @{
                    DisplayName = $parent.DisplayName
                    ParentId    = if ($parent.ParentFolderId) { $parent.ParentFolderId.UniqueId } else { $null }
                    Path        = $null
                }

                $parts += $parent.DisplayName
                $current = $parent
            }
            catch { break }
        }

        [array]::Reverse($parts)
        $full = ($parts -join '\')

        $FolderCache[$Folder.Id.UniqueId] = @{
            DisplayName = $Folder.DisplayName
            ParentId    = if ($Folder.ParentFolderId) { $Folder.ParentFolderId.UniqueId } else { $null }
            Path        = $full
        }

        return $full
    }

    ## Helper: approved verb for path resolution
    function Resolve-AllFolderPaths {
        foreach ($id in $FolderCache.Keys) {
            if ($FolderCache[$id].Path) { continue }

            $stack = @()
            $curId = $id
            while ($curId -and $FolderCache.ContainsKey($curId) -and -not $FolderCache[$curId].Path) {
                $stack += $curId
                $curId = $FolderCache[$curId].ParentId
            }

            $prefixParts = @()
            if ($curId -and $FolderCache.ContainsKey($curId) -and $FolderCache[$curId].Path) {
                $prefixParts = $FolderCache[$curId].Path -split '\\'
            }
            elseif ($curId -and $FolderCache.ContainsKey($curId)) {
                $prefixParts = @($FolderCache[$curId].DisplayName)
            }

            for ($i = $stack.Count - 1; $i -ge 0; $i--) {
                $nodeId = $stack[$i]
                $display = $FolderCache[$nodeId].DisplayName
                $prefixParts += $display
                $FolderCache[$nodeId].Path = ($prefixParts -join '\')
            }
        }
    }

    if (!($Token = Get-EwsAccessToken)) {
        Write-Error "EWS is not connected. Run the Connect-Ews command first."
        return $null
    }

    $InformationPreference = 'Continue'

    ## Create EWS object
    $Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService -ArgumentList 'Exchange2013_SP1'
    $Service.Url = 'https://outlook.office365.com/EWS/Exchange.asmx'
    $Service.UseDefaultCredentials = $false
    $Service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials -ArgumentList $Token
    $Service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId(
        [Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxAddress)
    $Service.HttpHeaders.Add('X-AnchorMailbox', $MailboxAddress)

    ## Determine mailbox root
    if ($MailboxType -eq 'Primary') {
        $ConnectToMailboxRootFolders = New-Object Microsoft.Exchange.WebServices.Data.FolderId(
            [Microsoft.Exchange.Webservices.Data.WellKnownFolderName]::MsgFolderRoot, $MailboxAddress)
    }
    elseif ($MailboxType -eq 'Archive') {
        $ConnectToMailboxRootFolders = New-Object Microsoft.Exchange.WebServices.Data.FolderId(
            [Microsoft.Exchange.Webservices.Data.WellKnownFolderName]::ArchiveMsgFolderRoot, $MailboxAddress)
    }

    ## Bind the mailbox root
    $EWSParentFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service, $ConnectToMailboxRootFolders)

    ## Create FolderView
    $FolderView = New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)
    $FolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep

    $MailboxFolderList = @()
    $FolderCache = @{}

    ## Start timing for FindFolders
    $swFind = [System.Diagnostics.Stopwatch]::StartNew()

    $MailboxFolderList = @($EWSParentFolder.FindFolders($FolderView))

    $swFind.Stop()
    Write-Verbose ("FindFolders completed in {0:N2} seconds" -f $swFind.Elapsed.TotalSeconds)

    if ($MailboxFolderList.Count -lt 1) {
        return $null
    }

    ## Path computation timing
    $swPath = [System.Diagnostics.Stopwatch]::StartNew()


    $FolderCache[$EWSParentFolder.Id.UniqueId] = @{
        DisplayName = $EWSParentFolder.DisplayName
        ParentId    = $null
        Path        = $EWSParentFolder.DisplayName
    }

    foreach ($f in $MailboxFolderList) {
        $FolderCache[$f.Id.UniqueId] = @{
            DisplayName = $f.DisplayName
            ParentId    = if ($f.ParentFolderId) { $f.ParentFolderId.UniqueId } else { $null }
            Path        = $null
        }
    }

    Resolve-AllFolderPaths

    foreach ($folder in $MailboxFolderList) {
        if ($FolderCache.ContainsKey($folder.Id.UniqueId)) {
            $folder | Add-Member -NotePropertyName 'Path' -NotePropertyValue $FolderCache[$folder.Id.UniqueId].Path -Force
        }
        else {
            $p = Get-FolderPath_Recursive -Folder $folder
            $folder | Add-Member -NotePropertyName 'Path' -NotePropertyValue $p -Force
        }
    }

    $swPath.Stop()
    Write-Verbose ("Path computation completed in {0:N2} seconds" -f $swPath.Elapsed.TotalSeconds)
    Write-Verbose "Retrieved $($MailboxFolderList.Count) folders."

    $MailboxFolderList | Add-Member -MemberType NoteProperty -Name MailboxAddress -Value $MailboxAddress
    $MailboxFolderList | Add-Member -MemberType NoteProperty -Name MailboxType -Value $MailboxType

    if ($PSCmdlet.ParameterSetName -eq 'byFolderName') {
        $MailboxFolderList | Where-Object { $_.DisplayName -eq $FolderName }
    }
    elseif ($PSCmdlet.ParameterSetName -eq 'byFolderId') {
        $MailboxFolderList | Where-Object { $_.Id.UniqueId -eq $FolderID }
    }
    else {
        $MailboxFolderList
    }
}