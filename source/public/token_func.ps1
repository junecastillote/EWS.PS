function Connect-Ews {
    [CmdletBinding(DefaultParameterSetName = 'CertificateThumbprint')]
    param(
        [Parameter(Mandatory)]
        [string]$TenantId,

        [Parameter(Mandatory)]
        [string]$ClientId,

        [Parameter(Mandatory, ParameterSetName = 'ClientSecret')]
        [string]$ClientSecret,

        [Parameter(Mandatory, ParameterSetName = 'CertificateThumbprint')]
        [string]$CertificateThumbprint,

        [Parameter()]
        [string]$CacheFolder = "$env:LOCALAPPDATA\Ews.PS.TokenCache"
    )

    $Scope = "https://outlook.office.com/.default"

    # Create cache folder if missing
    if (-not (Test-Path $CacheFolder)) {
        New-Item -Path $CacheFolder -ItemType Directory | Out-Null
    }

    # Generate cache filename based on hash of key parameters
    $hashInput = "$TenantId|$ClientId|$Scope"
    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    $hashBytes = $sha256.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($hashInput))
    $hashString = [BitConverter]::ToString($hashBytes) -replace "-", ""
    $CachePath = Join-Path $CacheFolder "$hashString.json"

    # region === Helper functions ===
    function Read-EncryptedCache($path) {
        try {
            $encrypted = Get-Content $path
            # Decrypt
            $secure = ConvertTo-SecureString $encrypted
            $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
            $json = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
            return ($json | ConvertFrom-Json)
        }
        catch {
            Write-Verbose "Failed to decrypt cache file: $_"
            return $null
        }
    }

    function Write-EncryptedCache($path, $data) {
        try {
            $json = $data | ConvertTo-Json -Compress
            # Encrypt
            $secure = ConvertTo-SecureString $json -AsPlainText -Force
            $encrypted = ConvertFrom-SecureString $secure
            Set-Content -Path $path -Value $encrypted -Encoding UTF8
        }
        catch {
            Write-Verbose "Failed to encrypt cache file: $_"
        }
    }

    function Get-NewToken {
        param(
            [string]$TenantId,
            [string]$ClientId,
            [string]$ClientSecret,
            [string]$CertificateThumbprint,
            [string]$Scope
        )

        $msalParams = @{
            ClientId = $ClientId
            TenantId = $TenantId
            Scopes   = $Scope
        }

        if ($CertificateThumbprint) {
            $cert = Get-ChildItem Cert:\CurrentUser\My\$CertificateThumbprint -ErrorAction SilentlyContinue
            if (-not $cert) {
                $cert = Get-ChildItem Cert:\LocalMachine\My\$CertificateThumbprint -ErrorAction SilentlyContinue
            }
            if (-not $cert) {
                throw "Certificate with thumbprint $CertificateThumbprint not found in CurrentUser or LocalMachine store."
            }
            Write-Verbose "Using certificate-based authentication"
            $msalParams.ClientCertificate = $cert
        }
        elseif ($ClientSecret) {
            Write-Verbose "Using client secret authentication"
            $msalParams.ClientSecret = (ConvertTo-SecureString $ClientSecret -AsPlainText -Force)
        }
        else {
            throw "Either -ClientSecret or -CertificateThumbprint must be specified."
        }

        try {
            $tokenResponse = Get-MsalToken @msalParams
            return $tokenResponse
        }
        catch {
            throw "Token acquisition failed: $_"
        }
    }
    # endregion

    # region === Load or acquire token ===
    $tokenResponse = $null

    if (Test-Path $CachePath) {
        Write-Verbose "Token cache file exists..."
        $cache = Read-EncryptedCache $CachePath
        if ($cache) {
            $expiry = [datetime]$cache.expires_on
            if ($expiry -gt (Get-Date).AddMinutes(5)) {
                Write-Verbose "Using cached token (expires $expiry)"
                $tokenResponse = [PSCustomObject]@{
                    AccessToken = $cache.access_token
                    ExpiresOn   = $expiry
                }
            }
            else {
                Write-Verbose "Cached token expired or near expiration — refreshing..."
                $tokenResponse = Get-NewToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret -CertificateThumbprint $CertificateThumbprint -Scope $Scope
                $tokenData = @{
                    access_token = $tokenResponse.AccessToken
                    expires_on   = $tokenResponse.ExpiresOn
                }
                Write-EncryptedCache -path $CachePath -data $tokenData
                Write-Verbose "New token acquired and cached (expires $($tokenData.expires_on))"
            }
        }
    }

    if (-not $tokenResponse) {
        Write-Verbose "No valid cache found — requesting new token..."
        $tokenResponse = Get-NewToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret -CertificateThumbprint $CertificateThumbprint -Scope $Scope
        $tokenData = @{
            access_token = $tokenResponse.AccessToken
            expires_on   = $tokenResponse.ExpiresOn
        }
        Write-EncryptedCache -path $CachePath -data $tokenData
        Write-Verbose "New token acquired and cached (expires $($tokenData.expires_on))"
    }

    # Store token in global session context
    $Global:EwsAuthContext = @{
        AccessToken           = $tokenResponse.AccessToken
        # ExpiresOn             = $tokenResponse.ExpiresOn.UtcDateTime
        ExpiresOn             = $tokenResponse.ExpiresOn.ToLocalTime()
        TenantId              = $TenantId
        ClientId              = $ClientId
        Scope                 = $Scope
        CachePath             = $CachePath
        AuthType              = if ($CertificateThumbprint) { "Certificate" } else { "ClientSecret" }
        CertificateThumbprint = $CertificateThumbprint
        ClientSecret          = $ClientSecret
    }

    Write-Verbose "Connected to EWS — token expires $($tokenResponse.ExpiresOn)"
    # endregion
}

# Helper: Returns valid token and auto-refreshes if expired
function Get-EwsAccessToken {
    [CmdletBinding()]
    param (

    )

    if (-not $Global:EwsAuthContext) {
        Write-Verbose "No EWS session found — run Connect-Ews first."
        return $null
    }

    Write-Verbose "EWS session found."
    # $expiresOn = [datetime]$Global:EwsAuthContext.ExpiresOn
    $expiresOn = $Global:EwsAuthContext.ExpiresOn
    if ($expiresOn -gt (Get-Date).AddMinutes(5)) {
        return $Global:EwsAuthContext.AccessToken
    }

    $cert = Get-ChildItem Cert:\CurrentUser\My\$($Global:EwsAuthContext.CertificateThumbprint) -ErrorAction SilentlyContinue
    if (-not $cert) {
        $cert = Get-ChildItem Cert:\LocalMachine\My\$($Global:EwsAuthContext.CertificateThumbprint) -ErrorAction SilentlyContinue
    }
    if (-not $cert) {
        throw "Certificate with thumbprint $CertificateThumbprint not found in CurrentUser or LocalMachine store."
    }

    Write-Verbose "EWS token expired or near expiration — auto-refreshing..."
    $tokenResponse = Get-MsalToken -ClientId $Global:EwsAuthContext.ClientId `
        -TenantId $Global:EwsAuthContext.TenantId `
        -Scopes $Global:EwsAuthContext.Scope `
    @(
        if ($Global:EwsAuthContext.AuthType -eq "Certificate") {
            # @{ ClientCertificate = (Get-ChildItem Cert:\CurrentUser\My\$($Global:EwsAuthContext.CertificateThumbprint) -ErrorAction SilentlyContinue) }
            @{ ClientCertificate = $cert }
        }
        else {
            @{ ClientSecret = (ConvertTo-SecureString $Global:EwsAuthContext.ClientSecret -AsPlainText -Force) }
        }
    )



    if ($tokenResponse -and $tokenResponse.AccessToken) {
        # Update cache + global context
        $tokenData = @{
            access_token = $tokenResponse.AccessToken
            # expires_on   = $tokenResponse.ExpiresOn.UtcDateTime
            expires_on   = $tokenResponse.ExpiresOn
        }
        Write-EncryptedCache -path $Global:EwsAuthContext.CachePath -data $tokenData
        $Global:EwsAuthContext.AccessToken = $tokenResponse.AccessToken
        # $Global:EwsAuthContext.ExpiresOn = $tokenResponse.ExpiresOn.UtcDateTime
        $Global:EwsAuthContext.ExpiresOn = $tokenResponse.ExpiresOn
        Write-Verbose "Token auto-refreshed successfully (expires $($tokenResponse.ExpiresOn))"
        return $tokenResponse.AccessToken
    }
    else {
        Write-Warning "Token refresh failed — please re-run Connect-Ews."
        return $null
    }
}
