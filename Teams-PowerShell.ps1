$clientId = Get-AutomationVariable -Name "clientId"
$tenantId = Get-AutomationVariable -Name "tenantId"
$graphScopes = "https://graph.microsoft.com/User.Read.All https://graph.microsoft.com/Group.ReadWrite.All https://graph.microsoft.com/AppCatalog.ReadWrite.All offline_access"
$teamsAdminScopes = "48ac35b8-9aa8-4d74-927d-1f4a14a0b239/user_impersonation"

# Get a token by signing in with a device code
function Get-TokenWithDeviceCode {
    param (
        [Parameter(Mandatory = $true)][string]$scopes,
        [Parameter(Mandatory = $true)][string]$resourceName
    )

    try {
        Write-Host "-----`n`rPlease sign in for $($resourceName) `n`r-----"

        # Get Device Code for all scopes with consent link
        $codeBody = @{ 
            client_id = $clientId
            scope     = $scopes
        }
        $codeRequest = Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/devicecode" -Body $codeBody
    
        # Print Device code and link to console
        Write-Host "`n$($codeRequest.message)"
        $tokenBody = @{
            grant_type  = "urn:ietf:params:oauth:grant-type:device_code"
            device_code = $codeRequest.device_code
            client_id   = $clientId
        }

        # Get token for the scopes
        while ([string]::IsNullOrEmpty($tokenRequest.access_token)) {
            $tokenRequest = try {
                Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Body $tokenBody
            }
            catch {
                $errorMessage = $_.ErrorDetails.Message | ConvertFrom-Json
                # If not waiting for auth, throw error
                if ($errorMessage.error -ne "authorization_pending") {
                    throw
                }
            }
            Start-Sleep -Seconds 1
        }
        # Set automation variable with new refresh token
        if ($tokenRequest.refresh_token) {
            Set-AutomationVariable -Name "refreshToken" -Value $tokenRequest.refresh_token
        }
    }
    catch {
        Write-Host "Error getting token: $($_.Exception.Message)"
        exit 1
    }
}

# Get a token by using a refresh token
function Get-TokenWithRefreshToken {
    param (
        [Parameter(Mandatory = $true)][string]$scopes,
        [Parameter(Mandatory = $true)][string]$refreshToken
    )
    try {
        $tokenBody = @{
            grant_type    = "refresh_token"
            refresh_token = $refreshToken
            client_id     = $clientId
            scope         = $scopes
        }
        $tokenRequest = Invoke-RestMethod -Method POST -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Body $tokenBody
        return $tokenRequest
    }
    catch {
        # Perhaps refresh token expired, get new token and try again
        Write-Host "Error getting token: $($_.Exception.Message)"
        Write-Host "Getting new tokens with device codes due to error"
        Get-NewRefreshTokens
        exit 1
    }
}

function Get-NewRefreshTokens {
    Get-TokenWithDeviceCode $teamsAdminScopes "Microsoft Teams Admin"
    Get-TokenWithDeviceCode $graphScopes "Microsoft Graph"
    Write-Host "Refresh token stored, please re-run the script"
}

# Get existing refresh token
$refreshToken = Get-AutomationVariable -Name "refreshToken"

if ($refreshToken) {
    Write-Host "Refresh token found, obtaining tokens"   
    $graphToken = Get-TokenWithRefreshToken $graphScopes $refreshToken
    $teamsAdminToken = Get-TokenWithRefreshToken $teamsAdminScopes $refreshToken

    # Set automation variable with new refresh token for next time
    if ($graphToken.refresh_token) {
        Set-AutomationVariable -Name "refreshToken" -Value $graphToken.refresh_token
    }
    
    Write-Host "Connecting to Microsoft Teams"
    Connect-MicrosoftTeams -AccessTokens @($graphToken.access_token, $teamsAdminToken.access_token)
    
    # Sample Teams PowerShell command
    Get-CsOnlineUser | Format-Table DisplayName, UserPrincipalName
}
else {
    Write-Host "No refresh token found, obtaining tokens with device codes"
    Get-NewRefreshTokens
    exit 1
}