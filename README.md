# Mac Battery Health

<img width="255" height="255" alt="image" src="https://github.com/user-attachments/assets/bda5fc93-df20-491a-8d9e-54005a1942c4" />

---

_All-in-one Mac battery health reporting solution that includes a SwiftUI macOS app, an Azure Automation runbook, and an Intune custom attribute script._

---

## 🎯 Purpose

You need actionable visibility into Mac battery health for lifecycle planning, proactive replacement, and better user experience.  
This repo offers modular tools you can use independently or in conjunction with each other:

- **Intune Custom Attribute (foundation)** – Deploy once to collect battery telemetry on each Mac into Intune/Graph.
- **Mac Battery Analyzer (SwiftUI app)** – Standalone local analysis and visualization. *(Optionally reads fleet data from Graph if the custom attribute is deployed.)*
- **Azure Automation Runbook** – Scheduled fleet report emailed as HTML + CSV. *(Requires the custom attribute; does **not** require the app.)*

### What you get
- A **macOS app** for real-time analysis, visualization, and export
- An **Azure Automation runbook** for scheduled fleet reporting
- An **Intune custom attribute script** for continuous inventory enrichment

By combining these tools, IT admins can **identify failing batteries early, plan replacements, and ensure device reliability** across a managed fleet.

---

## ✨ Features

### Mac Battery Analyzer (SwiftUI App)
- 📊 Fleet-wide KPIs: total devices, average health, high cycle counts, devices on external power
- 📈 Histogram view of battery health distribution
- 🔍 Interactive table with sorting, filtering, and search
- 💻 Local device panel with live battery/AC stats (refreshes every 30s)
- 📤 Export to **CSV** or **HTML**
- 🔗 Microsoft Graph/Intune integration with pagination and retry/backoff logic
<img width="1519" height="927" alt="image" src="https://github.com/user-attachments/assets/0acad953-1489-4623-9484-694780647365" />

### Azure Automation Runbook
- Automates a scheduled report of macOS battery health data
- Sends a rich HTML formatted email with a detailed csv report attached.
<img width="1725" height="1893" alt="image" src="https://github.com/user-attachments/assets/1ba7b040-7cf6-4c83-b718-2a9fc84e5740" />



### Intune Custom Attribute Script
- Collects local macOS battery statistics
- Publishes as Intune custom attributes for reporting/compliance

---

## 📋 Requirements

- macOS 14+  
- Xcode 15+ (SwiftUI, Combine, Charts)  
- Microsoft Intune license
- Azure Automation Resources (subscriptions, automation accounts)  
- Graph API permissions:
  - `DeviceManagementManagedDevices.Read.All`

---

## 🚀 Getting Started

### Deploy the Custom Attribute Script (Intune)
`battery-custom-attribute.sh`

#### Intune deployment
- In Intune admin center → Devices → macOS → Custom attributes for macOS → Add  
- Upload `battery-custom-attribute.sh`
- Choose **String** as the data type of the attribute  
- Assign to your target device group(s)  

---

#### Azure app registration (for the SwiftUI Mac app)

To allow the Mac Battery Analyzer app to authenticate users and call Microsoft Graph, you must create an **App Registration** in Entra ID (Azure AD):

1. **Sign in** to the [Azure portal](https://portal.azure.com) and open **Microsoft Entra ID** (formerly Azure AD).

2. In the left menu, go to **App registrations** → **+ New registration**.

3. Fill in the registration details:
   - **Name:** `Mac Battery Analyzer` (or something recognizable for your org)
   - **Supported account types:**  
     - Choose **Accounts in this organizational directory only (Single tenant)** if this is for your company only.  
     - Choose **Multitenant** only if you intend to distribute outside your org.
   - **Redirect URI:**  
     - Platform: **Mobile and desktop applications (iOS/macOS)**  
     - URI: `msauth.com.battery.analyzer://auth`  
       (replace `com.battery.analyzer` with the **Bundle Identifier** of your Xcode project, or keep it the same!

   Click **Register**.

4. Once created, copy down:
   - **Application (client) ID**  
   - **Directory (tenant) ID**  
   You’ll paste these into your Swift app’s configuration.

5. In the app registration blade, go to **Authentication**:
   - Under **Redirect URIs**, confirm the URI is present.
   - Check **Public client/native (mobile & desktop)** → Enable.
   - Save changes.

6. Go to **API permissions** → **+ Add a permission**:
   - Choose **Microsoft Graph** → **Delegated permissions**.
   - Search for and add:
     - `DeviceManagementManagedDevices.Read.All`
   - Click **Add permissions**.

7. Still in **API permissions**, click **Grant admin consent**

---

### Configure the SwiftUI App (Mac Battery Analyzer)
The Xcode project is sanitized. You must supply your own Azure AD (Entra ID) details.

#### Xcode Project setup

### Clone the repository
```bash
git clone https://github.com/<your-org>/Mac-Battery-Health.git
cd Mac-Battery-Health
```

### Configure the SwiftUI App (Mac Battery Analyzer)
The Xcode project is sanitized. You must supply your own Azure AD (Entra ID) details.


Open the project in Xcode:
```bash
open "Mac Battery Analyzer.xcodeproj"
```

In your AuthenticationManager.swift, update:
- Client ID
- Tenant ID
- Redirect URI
  <img width="1759" height="962" alt="image" src="https://github.com/user-attachments/assets/68238ff1-9378-40bc-8360-14a4ea51380c" />


#### Run
- Build & Run → Sign in with Microsoft  
- Devices load automatically

✅ Now when you run the app:
- The Microsoft sign-in window will appear.
<img width="1519" height="927" alt="image" src="https://github.com/user-attachments/assets/64bf843c-8eae-48fb-a35b-52a49173620c" />

- After login, the app will receive a token with delegated Graph permissions.
- You will see the app populate with devices and battery health information.
<img width="1519" height="927" alt="image" src="https://github.com/user-attachments/assets/67544ba4-39d7-404e-a32e-54e880878901" />


---


### Configure the Azure Automation Runbook
`Mac-Battery-Runbook.ps1` is sanitized. Edit placeholders before importing (Tenant/Client IDs, email sender/recipients).

### 🧩 What the Runbook Does

The **Mac-Battery-Runbook.ps1** automates a scheduled fleet report for macOS battery health:
- Queries Intune/Graph for managed macOS devices and their battery-related attributes (populated by your Intune custom attribute script).
- Aggregates results and generates:
  - A **CSV** attachment with per-device metrics (e.g., design capacity, cycle count, health %, external power state, last seen).
  - A **rich HTML** email body with summary KPIs and tables.
- Sends the email to the recipients you configure in the script.
- (Optional) Can be scheduled to run monthly (or at any cadence you choose).

> Tip: this runbook **reads** data that your `battery-custom-attribute.sh` posts to Intune. Make sure the custom attribute script is deployed and devices have reported at least once.

---

### 📦 Required PowerShell Modules

You need modules in **two places**:

1) **On your workstation** (only if you import/publish via PowerShell)
- `Az.Accounts`
- `Az.Automation`
- `Az.Resources` *(recommended when automating resource lookups)*

2) **In the Automation Account** (for the runbook to execute)
- `Az.Accounts` *(if your runbook calls any Az cmdlets or uses Managed Identity context helpers)*
- **One of the following for Microsoft Graph** (pick the one your script uses):
  - **If using Graph PowerShell SDK**:  
    - `Microsoft.Graph.Authentication`  
    - `Microsoft.Graph.DeviceManagement` *(and/or other sub-modules you use, e.g., `Microsoft.Graph.Users`, `Microsoft.Graph.Mail` if the runbook uses `Send-MgUserMail`)*
  - **If using raw REST calls** with `Invoke-RestMethod`:  
    - No Graph module required, but you must obtain a token (Managed Identity or App Registration) and add the correct **Graph App Roles**.

> If your runbook sends email through Graph, include `Microsoft.Graph.Mail` (or a parent bundle that contains it).  
> If it sends via SMTP/Exchange Online, include the relevant module/secret configuration instead.

---

### 🔐 Authentication & Permissions

Choose **one** auth model:

**A) Managed Identity (recommended)**
The runbook script is currently configured to use Managed Identity. If you wish to use other authentication methods, that will need to be adjusted in the script. 
- Enable **System-assigned Managed Identity** on the Automation Account.
- Grant the MI the following **Microsoft Graph Application permissions** (App Roles):
  - `DeviceManagementManagedDevices.Read.All` *(to read Intune devices & custom attributes)*
  - `Mail.Send` *(only if the runbook uses Graph to send email)*
- **Admin consent** is required after adding App Roles.

> If you wish to manually grant these via PowerShell, run the following:
### Grant Graph API permissions to the Automation Account (Managed Identity)

```powershell
# Connect to Microsoft Graph with sufficient privileges (e.g. Global Admin or Cloud App Admin)
Connect-MgGraph -Scopes "Application.ReadWrite.All"

# Get the Service Principal for the Automation Account Managed Identity
$managedIdentity = Get-MgServicePrincipal -Filter "displayName eq 'your-automation-account'"

# Required permissions for Mac-Battery-Runbook.ps1
$permissions = @(
    "DeviceManagementManagedDevices.Read.All",  # Read Intune managed devices + custom attributes
    "Mail.Send"                                 # Send reports via Graph Mail API
)

# Get the Microsoft Graph service principal (well-known appId)
$graphServicePrincipal = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'"

# Loop through and assign each role
foreach ($permission in $permissions) {
    $appRole = $graphServicePrincipal.AppRoles | Where-Object { $_.Value -eq $permission }

    if ($appRole) {
        New-MgServicePrincipalAppRoleAssignment `
            -ServicePrincipalId $managedIdentity.Id `
            -PrincipalId $managedIdentity.Id `
            -ResourceId $graphServicePrincipal.Id `
            -AppRoleId $appRole.Id
        Write-Output "Assigned $permission to $($managedIdentity.DisplayName)"
    }
    else {
        Write-Warning "Permission $permission not found on Graph SP"
    }
}
```


**B) App Registration (client secret/cert)**
- Create an Entra ID App Registration and assign the same Graph **Application** permissions above.
- Store the secret/cert in **Automation Account → Variables/Credentials/Key Vault**.
- The runbook will acquire a token using this app to call Graph.

> If you only use **REST** + token acquisition, you still need App Roles + consent.  
> If you use **Graph PowerShell SDK**, the same App Roles apply.

---

### 🧰 Installing Modules

**Local (PowerShell)**
```powershell
# Install for the current user
Install-Module Az.Accounts,Az.Automation,Az.Resources -Scope CurrentUser -Repository PSGallery -Force
# Optional Graph SDK if you import/use it locally
Install-Module Microsoft.Graph -Scope CurrentUser -Repository PSGallery -Force
```

**Automation Account (Portal)**
1. Go to **Automation Account → Shared Resources → Modules**.
2. Click **+ Browse Gallery**, search and import:
   - `Az.Accounts` (and any other Az modules you require)
   - `Microsoft.Graph.Authentication`
   - `Microsoft.Graph.DeviceManagement`
   - `Microsoft.Graph.Mail` *(if using Graph to send email)*
3. Wait for modules to finish provisioning (status becomes **Available**).

> If your script imports specific Graph sub-modules directly (e.g., `Import-Module Microsoft.Graph.DeviceManagement`), ensure those exact sub-modules exist in the Automation Account.

---

### ⚙️ Configure `Mac-Battery-Runbook.ps1`

Edit these **three** placeholders in the script:

1) **Recipient email** (default param)
- Find the `param` block at the top and set your distro/list:
```powershell
[string]$EmailRecipient = 'team-email@contoso.com',
```

2) **Custom Attribute ID (GUID)**
- Replace the placeholder with your **Intune Custom Attribute ID** (the GUID from Intune → macOS → Custom attributes):
```powershell
[string]$CustomAttributeId = '11111111-2222-3333-4444-555555555555',
```

3) **Sender mailbox UPN used for Graph sendMail**
- Update the **UPN** in the `sendMail` call (must be a real, licensed mailbox, e.g., ITAutomation@contoso.com):
```powershell
$emailUri = "https://graph.microsoft.com/v1.0/users/ITAutomation@contoso.com/sendMail"

```
<img width="839" height="208" alt="image" src="https://github.com/user-attachments/assets/ed3fdcc5-3892-470f-a052-518aa17c2040" />
<img width="1104" height="159" alt="image" src="https://github.com/user-attachments/assets/5a74f2b7-0e51-4438-8424-eb3b549d50f2" />

---

### 🔐 Required Graph App Roles for the Managed Identity

Grant these **Application permissions** (App Roles) to your Automation Account’s MI and **admin consent** them:
- `DeviceManagementManagedDevices.Read.All`
- `Mail.Send`

---

### ✅ Quick Verification

After you **Publish** the runbook:
1. Click **Start** (Run) once manually.
2. Confirm:
   - The runbook can acquire a token (MI or App Reg).
   - Devices are returned from Intune.
   - Email is delivered with the HTML body and CSV attachment.
3. Then link your **Monthly** schedule.

---

### 🧪 Troubleshooting

- **Runbook fails at token acquisition**  
  - Verify API permissions and admin consent.  
  - If using App Reg, confirm secret/cert exists and is referenced correctly.
- **No devices returned**  
  - Ensure your Intune custom attribute script is deployed and devices have checked in.  
  - Confirm the Graph query/filters target **macOS** devices.
- **Email not delivered**  
  - If using Graph: ensure `Mail.Send` App Role and the sender is allowed.  
  - If using SMTP: validate SMTP host, port, auth method, and credentials/Key Vault references.

---

---

## ✍️ Author
**Eddie Jimenez** ([@eddie-jimenez](https://github.com/eddie-jimenez))

## 💬 Contact
- Mac Admins Slack: [**@Eddie J**](https://macadmins.slack.com/team/U06SP52GZSM) *(requires Mac Admins Slack membership)*
- For bugs/requests: open a GitHub issue in this repo

## 📄 License
Released under the **MIT License** — free to use, modify, and distribute.  
See [`LICENSE`](./LICENSE) for the full text.

## 🙌 Final Notes
Built to be useful. If it helps you, awesome — try to credit me and enjoy!




