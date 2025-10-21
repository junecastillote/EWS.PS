function Get-AppOnlyTokenWithCertCache {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)] [string]$TenantId,
        [Parameter(Mandatory)] [string]$ClientId,
        [Parameter(Mandatory)] [string]$CertThumbprint,
        # [Parameter()] [string]$Scope = "https://outlook.office.com/.default",
        [string]$CacheFolder = "$env:TEMP\AppTokenCache",
        [int]$RefreshThresholdMinutes = 5,          # Refresh if token expires within 5 minutes
        [switch]$ForceRefresh                        # Skip cache and force refresh
    )

    if (-not $Scope) {
        $Scope = "https://outlook.office.com/.default"
    }

    # region --- Helper: Base64Url encode ---
    function Base64UrlEncode([byte[]]$bytes) {
        [Convert]::ToBase64String($bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')
    }
    # endregion

    # region --- Helper: Generate cache file name per tenant/client/scope ---
    if (-not (Test-Path $CacheFolder)) {
        New-Item -Path $CacheFolder -ItemType Directory -Force | Out-Null
    }

    # Unique hash for this Tenant + Client + Scope combination
    $hashInput = "$TenantId|$ClientId|$Scope"
    $hash = [BitConverter]::ToString((New-Object Security.Cryptography.SHA256Managed).ComputeHash([Text.Encoding]::UTF8.GetBytes($hashInput))).Replace("-", "")
    $cachePath = Join-Path $CacheFolder "$hash.json"
    # endregion

    # region --- Try to load cached token ---
    $cachedToken = $null
    if (-not $ForceRefresh -and (Test-Path $cachePath)) {
        try {
            $cachedToken = Get-Content $cachePath -Raw | ConvertFrom-Json
        }
        catch {}
    }

    if ($cachedToken -and $cachedToken.access_token -and $cachedToken.expires_on) {
        $expiresOn = [datetime]::Parse($cachedToken.expires_on)
        $timeLeft = ($expiresOn - (Get-Date)).TotalMinutes
        if ($timeLeft -gt $RefreshThresholdMinutes) {
            Write-Verbose "Using cached token (expires in $([math]::Round($timeLeft)) minutes)"
            return $cachedToken
        }
        else {
            Write-Verbose "Token nearing expiration — refreshing..."
        }
    }
    elseif ($ForceRefresh) {
        Write-Verbose "Force refresh requested — ignoring cache..."
    }
    # endregion

    # region --- Acquire certificate ---
    $cert = Get-ChildItem Cert:\CurrentUser\My\$CertThumbprint -ErrorAction SilentlyContinue
    if (-not $cert) {
        $cert = Get-ChildItem Cert:\LocalMachine\My\$CertThumbprint -ErrorAction SilentlyContinue
    }
    if (-not $cert) {
        throw "Certificate with thumbprint $CertThumbprint not found in CurrentUser or LocalMachine store."
    }
    # endregion

    # region --- Create signed JWT ---
    $header = @{
        alg = "RS256"
        typ = "JWT"
        x5t = [System.Convert]::ToBase64String(($cert.GetCertHash()))
    }

    $now = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
    $payload = @{
        aud = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
        iss = $ClientId
        sub = $ClientId
        jti = [guid]::NewGuid().ToString()
        nbf = $now
        exp = $now + 600  # valid 10 minutes
    }

    $headerEncoded = Base64UrlEncode ([System.Text.Encoding]::UTF8.GetBytes((ConvertTo-Json $header -Compress)))
    $payloadEncoded = Base64UrlEncode ([System.Text.Encoding]::UTF8.GetBytes((ConvertTo-Json $payload -Compress)))
    $jwtToSign = "$headerEncoded.$payloadEncoded"

    # Sign JWT with private key
    $rsa = $cert.GetRSAPrivateKey()
    $signature = $rsa.SignData([System.Text.Encoding]::UTF8.GetBytes($jwtToSign),
        [Security.Cryptography.HashAlgorithmName]::SHA256,
        [Security.Cryptography.RSASignaturePadding]::Pkcs1)
    $signedJwt = "$jwtToSign.$(Base64UrlEncode $signature)"
    # endregion

    # region --- Request token from Entra ID ---
    $body = @{
        client_id             = $ClientId
        scope                 = $Scope
        client_assertion      = $signedJwt
        client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
        grant_type            = 'client_credentials'
    }

    $tokenEndpoint = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"

    try {
        $response = Invoke-RestMethod -Method Post -Uri $tokenEndpoint -Body $body
        $expiresOn = (Get-Date).AddSeconds([int]$response.expires_in)
        $cachedToken = [PSCustomObject]@{
            access_token = $response.access_token
            expires_on   = $expiresOn.ToString("u")
            scope        = $Scope
            token_type   = $response.token_type
        }

        # Save token cache
        $cachedToken | ConvertTo-Json | Set-Content -Path $cachePath -Encoding UTF8
        Write-Verbose "Token acquired and cached until $($expiresOn.ToLocalTime())"
        return $cachedToken
    }
    catch {
        throw "Failed to acquire token: $($_.Exception.Message)"
    }
    # endregion
}
