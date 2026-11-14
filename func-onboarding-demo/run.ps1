using namespace System.Net

# NewHire - HTTP-triggered Azure Function (PowerShell)
# Creates the user in Entra ID via Microsoft Graph and documents it in ITFlow
# (free, open-source IT documentation/PSA for MSPs): contact + credential + ticket.
#
# App settings:
#   TENANT_DOMAIN            e.g. contoso.onmicrosoft.com
#   Graph auth, option B (app registration in the target tenant):
#     GRAPH_TENANT_ID, GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET
#   Graph auth, option A: leave GRAPH_CLIENT_SECRET empty -> Managed Identity
#   ITFlow (step skipped while ITFLOW_BASE_URL is empty):
#     ITFLOW_BASE_URL          e.g. https://itflow.example.com
#     ITFLOW_API_KEY           API key scoped to the demo client
#     ITFLOW_DECRYPT_PASSWORD  the key's credential (decrypt) password
#     ITFLOW_CLIENT_ID         numeric id of the demo client

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
    $nick = ("$first.$last").ToLower() -replace '[^a-z0-9.]', ''
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
        usageLocation     = 'US'   # required later for license assignment
        passwordProfile   = @{
            forceChangePasswordNextSignIn = $true
            password                      = $password
        }
    } | ConvertTo-Json -Depth 3

    $user = Invoke-RestMethod -Method Post -Uri 'https://graph.microsoft.com/v1.0/users' `
        -Headers $graphHeaders -Body $newUser -ContentType 'application/json'

    # 4. Document in ITFlow (third-party): contact -> credential -> ticket
    $itflow = 'skipped (ITFLOW_BASE_URL not set)'
    if ($env:ITFLOW_BASE_URL) {
        # ITFlow returns HTTP 200 with success "False" on errors, so check it explicitly
        function Invoke-ITFlow([string]$Endpoint, [hashtable]$Body) {
            $Body.api_key   = $env:ITFLOW_API_KEY
            $Body.client_id = [int]$env:ITFLOW_CLIENT_ID
            $script:step    = "ITFlow $Endpoint"
            $r = Invoke-RestMethod -Method Post -Uri "$($env:ITFLOW_BASE_URL)/api/v1/$Endpoint" `
                -ContentType 'application/json' -Body ($Body | ConvertTo-Json)
            if ($r.success -ne 'True') { throw "ITFlow $Endpoint failed: $($r.message)" }
            return $r
        }

        # 4a. Contact
        $contactId = (Invoke-ITFlow 'contacts/create.php' @{
            contact_name       = $user.displayName
            contact_email      = $upn
            contact_title      = $in.jobTitle
            contact_department = $in.department
        }).data[0].insert_id

        # 4b. Credential, linked to the contact (ITFlow encrypts it)
        Invoke-ITFlow 'credentials/create.php' @{
            api_key_decrypt_password = $env:ITFLOW_DECRYPT_PASSWORD
            credential_name          = "Microsoft 365 - $($user.displayName)"
            credential_username      = $upn
            credential_password      = $password
            credential_uri           = 'https://portal.office.com'
            credential_contact_id    = $contactId
            credential_description   = "Created by Azure Functions onboarding on $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
        } | Out-Null

        # 4c. Ticket for the audit trail, then resolve it
        $ticketId = (Invoke-ITFlow 'tickets/create.php' @{
            ticket_subject    = "New hire onboarding: $($user.displayName)"
            ticket_details    = "Account $upn created automatically ($($in.jobTitle), $($in.department)). Credentials stored in ITFlow."
            ticket_contact_id = $contactId
        }).data[0].insert_id

        Invoke-ITFlow 'tickets/resolve.php' @{ ticket_id = $ticketId } | Out-Null
        $itflow = @{ contactId = $contactId; ticketId = $ticketId }
    }

    # The password never leaves the function except into ITFlow
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::OK
        Body       = @{ status = 'created'; userPrincipalName = $upn; id = $user.id; itflow = $itflow }
    })
}
catch {
    # ErrorDetails holds the response body (Graph and ITFlow explain the reason there)
    $detail = if ($_.ErrorDetails.Message) { $_.ErrorDetails.Message } else { $_.Exception.Message }
    Write-Error "Onboarding failed for $first $last at step '$step': $detail"
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::InternalServerError
        Body       = @{ status = 'failed'; step = $step; error = $detail }
    })
}