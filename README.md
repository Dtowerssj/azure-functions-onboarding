# Serverless employee onboarding with Azure Functions

Demo code from the talk **Serverless IT Operations**: a new-hire request goes from
Microsoft Forms to Power Automate, which calls an Azure Function written in PowerShell.
The function creates the user in Entra ID through Microsoft Graph.

```
Microsoft Forms ──► Power Automate ──HTTP POST──► Azure Function (PowerShell) ──► Microsoft Graph (Entra ID)
                                                                              └─► ITFlow (optional)
```

The Function App has two functions. Deploy both and call the one you need:

| Function | What it does | Use it if |
|---|---|---|
| `func-onboarding-demo` | Creates the user in Entra ID, then documents it in [ITFlow](https://itflow.org) (contact, encrypted credential, resolved ticket). The password goes only to ITFlow. | You use ITFlow as your IT documentation / PSA. |
| `func-onboarding-entra` | Creates the user in Entra ID and returns a one-time credential (temporary password or Temporary Access Pass) for the flow to hand over. | You don't use ITFlow. |

## Repository layout

```
func-onboarding-demo/     Entra ID + ITFlow  (run.ps1, function.json, sample.dat)
func-onboarding-entra/    Entra ID only      (run.ps1, function.json, sample.dat)
host.json, profile.ps1, requirements.psd1   Function App settings (PowerShell 7.4, Functions v4)
local.settings.example.json                 Template for local runs: copy to local.settings.json
.vscode/                                    VS Code launch/debug settings
Serverless IT Operations.pptx               Slides from the talk
```

## Prerequisites

- An Azure subscription to host the Function App (Consumption or Flex Consumption plan is fine).
- A Microsoft 365 / Entra ID tenant where the users will be created. For testing, use a
  [Microsoft 365 developer sandbox](https://developer.microsoft.com/microsoft-365/dev-program).
- [Azure Functions Core Tools v4](https://learn.microsoft.com/azure/azure-functions/functions-run-local),
  PowerShell 7.4 and VS Code with the Azure Functions extension.
- Power Automate and Microsoft Forms (for the end-to-end flow).

## 1. Give the function access to Microsoft Graph

The function gets a Graph token in one of two ways.

### Option A: App registration (works across tenants)

Use this when the Function App lives in a different tenant from the users (for example,
your Azure subscription is personal and the users live in a developer sandbox).

1. In the **target** tenant: Entra admin center → App registrations → New registration.
2. API permissions → Microsoft Graph → **Application** permissions:
   - `User.ReadWrite.All` (both functions)
   - `UserAuthenticationMethod.ReadWrite.All` (only if you use `CREDENTIAL_TYPE=tap`)
3. **Grant admin consent**.
4. Certificates & secrets → New client secret. Copy the value: it is shown only once.
5. Set `GRAPH_TENANT_ID`, `GRAPH_CLIENT_ID` and `GRAPH_CLIENT_SECRET` in the Function App.

### Option B: Managed identity (same tenant, no secret)

When the Function App and the users are in the **same** tenant, leave `GRAPH_CLIENT_SECRET`
empty and the function uses the Function App's managed identity.

1. Function App → Identity → System assigned → **On**.
2. Grant it Graph application permissions (no portal UI for this; use PowerShell):

```powershell
Connect-MgGraph -Scopes 'AppRoleAssignment.ReadWrite.All', 'Application.Read.All'
$graph = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"
$mi    = Get-MgServicePrincipal -Filter "displayName eq '<your-function-app-name>'"
foreach ($perm in 'User.ReadWrite.All') {   # add 'UserAuthenticationMethod.ReadWrite.All' for TAP
    $role = $graph.AppRoles | Where-Object Value -eq $perm
    New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $mi.Id `
        -PrincipalId $mi.Id -ResourceId $graph.Id -AppRoleId $role.Id
}
```

## 2. App settings

Set these in the Function App (Settings → Environment variables) or in `local.settings.json`
for local runs. **Never commit real values**: `local.settings.json` is in `.gitignore`.

| Setting | Function | Description |
|---|---|---|
| `TENANT_DOMAIN` | both | Domain for new UPNs, e.g. `contoso.onmicrosoft.com` |
| `GRAPH_TENANT_ID`, `GRAPH_CLIENT_ID`, `GRAPH_CLIENT_SECRET` | both | App registration (option A). Leave the secret empty to use the managed identity. |
| `CREDENTIAL_TYPE` | entra | `password` (default) or `tap` for a [Temporary Access Pass](https://learn.microsoft.com/entra/identity/authentication/howto-authentication-temporary-access-pass). TAP needs the TAP policy enabled in Entra. |
| `TAP_LIFETIME_MINUTES` | entra | Optional, default `480` |
| `USAGE_LOCATION` | entra | Optional, default `US` (needed to assign licenses later) |
| `ITFLOW_BASE_URL` | demo | Your ITFlow URL. If empty, the ITFlow step is skipped. |
| `ITFLOW_API_KEY` | demo | ITFlow API key, scoped to one client |
| `ITFLOW_DECRYPT_PASSWORD` | demo | The API key's decrypt password (needed to store credentials) |
| `ITFLOW_CLIENT_ID` | demo | Numeric id of the ITFlow client |

## 3. Run locally

```powershell
Copy-Item local.settings.example.json local.settings.json   # then fill in your values
func start
```

```powershell
Invoke-RestMethod -Method Post -Uri http://localhost:7071/api/func-onboarding-entra `
    -ContentType 'application/json' `
    -Body '{"firstName":"Test","lastName":"User","jobTitle":"Analyst","department":"IT"}'
```

## 4. Deploy

From VS Code: Azure Functions → **Deploy to Function App**, or:

```powershell
func azure functionapp publish <your-function-app-name>
```

Then set the app settings from step 2 in the portal and copy the function URL
(Function → **Get function URL**, it includes the `code=` key).

## 5. Power Automate flow

1. Trigger: **When a new response is submitted** (Microsoft Forms) → **Get response details**.
2. **HTTP** action: `POST` to the function URL, header `Content-Type: application/json`, body:
   ```json
   {
     "firstName": "@{outputs('Get_response_details')?['body/<first-name-question-id>']}",
     "lastName":  "@{outputs('Get_response_details')?['body/<last-name-question-id>']}",
     "jobTitle":  "@{outputs('Get_response_details')?['body/<job-title-question-id>']}",
     "department":"@{outputs('Get_response_details')?['body/<department-question-id>']}"
   }
   ```
3. **Parse JSON** on the response and branch on `status` (`created` / `failed`).
4. With `func-onboarding-entra`: send `credential.value` to the hiring manager (for example
   with **Send an email (V2)**). Turn on **Settings → Secure outputs** on the HTTP action (and
   **Secure inputs** on the email action) so the credential is not stored in the flow's run history.

Treat the function URL as a secret: the `code=` key in it is what authorizes the call.

## Responses

```jsonc
// 200, func-onboarding-entra
{ "status": "created", "userPrincipalName": "test.user@contoso.onmicrosoft.com", "id": "<object-id>",
  "credential": { "type": "password", "value": "<one-time password>", "mustChangeAtNextSignIn": true } }

// 200, func-onboarding-demo
{ "status": "created", "userPrincipalName": "test.user@contoso.onmicrosoft.com", "id": "<object-id>",
  "itflow": { "contactId": 12, "ticketId": 34 } }

// 400 / 500, both
{ "status": "failed", "step": "Graph create user", "error": "<message from Graph or ITFlow>" }
```

## Notes

- The UPN is `firstname.lastname@TENANT_DOMAIN`. If that UPN already exists, Graph returns an
  error and the function reports it with `"step": "Graph create user"`.
- New users must change their password at first sign-in.
- Use a test tenant while you try this out: the function creates real accounts.
