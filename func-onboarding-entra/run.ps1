using namespace System.Net

# NewHire (Entra only) - HTTP-triggered Azure Function (PowerShell)
# Same as func-onboarding-demo, without ITFlow: creates the user in Entra ID via
# Microsoft Graph and returns a one-time credential for the flow to hand over.
#
# App settings:
#   TENANT_DOMAIN            e.g. contoso.onmicrosoft.com
#   Graph auth, option B (app registration in the target tenant):
#     GRAPH_TENANT_ID, GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET
#   Graph auth, option A: leave GRAPH_CLIENT_SECRET empty -> Managed Identity
#   CREDENTIAL_TYPE          'password' (default) or 'tap' (Temporary Access Pass)
#                            'tap' needs the TAP policy enabled in Entra and the
#                            UserAuthenticationMethod.ReadWrite.All permission
#   TAP_LIFETIME_MINUTES     optional, default 480 (8 hours)
#   USAGE_LOCATION           optional, default US (needed later to assign licenses)
#
# The response contains a credential: turn on "Secure outputs" on the HTTP action
# in Power Automate so it is not stored in the run history.

param($Request, $TriggerMetadata)

$step = 'read input'

try {
    # 0. Read the body: an object when sent as JSON, a string otherwise
    $in = $Request.Body
    if ($in -is [string]) { $in = if ($in.Trim()) { $in | ConvertFrom-Json } else { $null } }
    $first = "$($in.firstName)".Trim()
    $last  = "$($in.lastName)".Trim()
    if (-not $first -or -not $last) {
        Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
            StatusCode = [HttpStatusCode]::BadRequest
            Body       = @{ status = 'failed'; step = $step; error = 'firstName and lastName are required (send JSON with Content-Type: application/json)' }
        })
        return
    }
    $credentialType = if ($env:CREDENTIAL_TYPE) { $env:CREDENTIAL_TYPE.ToLower() } else { 'password' }

    # 1. Get a Graph token
    $step = 'Graph token'
    if ($env:GRAPH_CLIENT_SECRET) {
        $token = (Invoke-RestMethod -Method Post `
            -Uri "https://login.microsoftonline.com/$($env:GRAPH_TENANT_ID)/oauth2/v2.0/token" `
            -Body @{
                client_id     = $env:GRAPH_CLIENT_ID
                client_secret = $env:GRAPH_CLIENT_SECRET
                scope         = 'https://graph.microsoft.com/.default'
                grant_type    = 'client_credentials'
            }).access_token
    }
    else {
        $tokenUri = "$($env:IDENTITY_ENDPOINT)?resource=https://graph.microsoft.com&api-version=2019-08-01"
        $token = (Invoke-RestMethod -Uri $tokenUri -Headers @{ 'X-IDENTITY-HEADER' = $env:IDENTITY_HEADER }).access_token
    }
    $graphHeaders = @{ Authorization = "Bearer $token" }

    # 2. Build the UPN and a temporary password
    # José García-Núñez -> jose.garcianunez (drop accents first, then anything not a-z 0-9 .)
    $nick = ("$first.$last").Normalize([Text.NormalizationForm]::FormD) -replace '\p{Mn}', ''
    $nick = $nick.ToLower() -replace '[^a-z0-9.]', ''
    $upn  = "$nick@$($env:TENANT_DOMAIN)"

    $upper = 'ABCDEFGHJKLMNPQRSTUVWXYZ'; $lower = 'abcdefghijkmnopqrstuvwxyz'
    $digit = '23456789';                  $symbol = '!@#$%*?'
    $all   = $upper + $lower + $digit + $symbol
    $chars = @(
        $upper[(Get-Random -Maximum $upper.Length)]
        $lower[(Get-Random -Maximum $lower.Length)]
        $digit[(Get-Random -Maximum $digit.Length)]
        $symbol[(Get-Random -Maximum $symbol.Length)]
    ) + (1..12 | ForEach-Object { $all[(Get-Random -Maximum $all.Length)] })
    $password = -join ($chars | Sort-Object { Get-Random })

    # 3. Create the user in Entra ID
    $step = 'Graph create user'
    $newUser = @{
        accountEnabled    = $true
        displayName       = "$first $last"
        givenName         = $first
        surname           = $last
        mailNickname      = $nick
        userPrincipalName = $upn
        jobTitle          = $in.jobTitle
        department        = $in.department
        usageLocation     = if ($env:USAGE_LOCATION) { $env:USAGE_LOCATION } else { 'US' }
        passwordProfile   = @{
            forceChangePasswordNextSignIn = $true
            password                      = $password
        }
    } | ConvertTo-Json -Depth 3

    $user = Invoke-RestMethod -Method Post -Uri 'https://graph.microsoft.com/v1.0/users' `
        -Headers $graphHeaders -Body $newUser -ContentType 'application/json'

    # 4. One-time credential for the new hire
    if ($credentialType -eq 'tap') {
        # Temporary Access Pass: lets the user sign in once and register MFA / passwordless
        $step = 'Graph Temporary Access Pass'
        $lifetime = if ($env:TAP_LIFETIME_MINUTES) { [int]$env:TAP_LIFETIME_MINUTES } else { 480 }
        $tapBody  = @{ lifetimeInMinutes = $lifetime; isUsableOnce = $false } | ConvertTo-Json
        $tap = $null
        # A just-created user can take a few seconds to replicate: retry briefly
        for ($i = 1; -not $tap; $i++) {
            try {
                $tap = Invoke-RestMethod -Method Post `
                    -Uri "https://graph.microsoft.com/v1.0/users/$($user.id)/authentication/temporaryAccessPassMethods" `
                    -Headers $graphHeaders -Body $tapBody -ContentType 'application/json'
            }
            catch {
                if ($i -ge 5) { throw }
                Start-Sleep -Seconds 3
            }
        }
        $credential = @{
            type      = 'temporaryAccessPass'
            value     = $tap.temporaryAccessPass
            expiresAt = ([datetime]$tap.startDateTime).AddMinutes($tap.lifetimeInMinutes).ToString('o')
        }
    }
    else {
        $credential = @{ type = 'password'; value = $password; mustChangeAtNextSignIn = $true }
    }

    # Secure outputs on the Power Automate HTTP action keeps the credential out of run history
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::OK
        Body       = @{ status = 'created'; userPrincipalName = $upn; id = $user.id; credential = $credential }
    })
}
catch {
    # ErrorDetails holds the response body (Graph explains the reason there)
    $detail = if ($_.ErrorDetails.Message) { $_.ErrorDetails.Message } else { $_.Exception.Message }
    Write-Error "Onboarding failed for $first $last at step '$step': $detail"
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::InternalServerError
        Body       = @{ status = 'failed'; step = $step; error = $detail }
    })
}
