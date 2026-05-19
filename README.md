# M365 Power Platform Inventory

This repository contains a PowerShell script that inventories Power Platform artifacts across environments and exports the results to CSV.

## Script

- Script file: Get-M365PowerPlatformInventory.ps1
- Output: A CSV inventory of Flows, Apps, and Agents

## What the script collects

For each environment, the script attempts to collect:

- Flows (Cloud Flows)
- Apps (Power Apps)
- Agents (Copilot Studio / Power Virtual Agents)

It also enriches owner details and creation timestamps when available, and can optionally include connection references and agent tool/action summaries.

## Requirements

The script can install missing modules automatically unless -SkipModuleInstall is used.

Commonly used modules/cmdlets:

- Microsoft.PowerApps.Administration.PowerShell
- Microsoft.PowerApps.PowerShell
- Az.Accounts
- Microsoft.Graph.Authentication
- Microsoft.Graph.Users
- pac (Power Platform CLI, optional fallback path)

## Required Permissions

To run this inventory script successfully, your user account must have sufficient permissions in the M365 tenant:

### Power Platform Admin Permissions
- **Power Platform Administrator** role (tenant-wide or environment-specific)
- **Dynamics 365 Administrator** role (for Dataverse access)
- **System Administrator** security role in each Dataverse environment being inventoried

### Microsoft Entra / Microsoft 365 Permissions
- **Global Administrator** or **Cloud Application Administrator** role to read user/service principal details via Microsoft Graph
- Read access to user profiles and organizational data

### What These Permissions Enable
These permissions allow the script to:
- Access all Power Platform environments across the tenant
- Retrieve detailed information about Flows, Apps, and Agents
- Resolve owner information using Dataverse queries and Microsoft Graph lookups
- Access environment metadata and artifact properties

### Running as a Service Principal
If running the script as an application/service principal (non-interactive), ensure it has:
- Power Platform API permissions with Administrator scope
- Microsoft Graph permissions: `User.Read.All`, `Application.Read.All`
- Dataverse application user with System Administrator role

## Usage

Run from this repository folder:

```powershell
.\Get-M365PowerPlatformInventory.ps1
```

Include connection references and agent tool/action summaries:

```powershell
.\Get-M365PowerPlatformInventory.ps1 -IncludeConnectionReferences
```

Write to a specific output file:

```powershell
.\Get-M365PowerPlatformInventory.ps1 -OutputCsvPath .\inventory.csv
```

Disable automatic module installation:

```powershell
.\Get-M365PowerPlatformInventory.ps1 -SkipModuleInstall
```

## Output column dictionary

The CSV contains the following columns.

### ArtifactType
Type of artifact in the row.

Expected values:

- Flow
- App
- Agent

### Environment
Display name of the Power Platform environment where the artifact exists.

### EnvironmentId
Environment identifier used by admin APIs/cmdlets.

### GUID
Artifact identifier.

Examples:

- Flow: Flow ID / workflow ID
- App: App ID
- Agent: Bot/Copilot ID

### ArtifactName
Human-readable artifact name.

Examples:

- Flow display name
- App display name
- Agent display name

### Owner
Primary owner value shown for reporting.

Behavior:

- Flow/App: Prefer display name + email when available; otherwise email; otherwise owner object ID.
- Agent: If OwnerObjectId is available, Owner is written as that GUID for consistency.

### OwnerEmail
Resolved owner email/UPN when available.

Resolution path can include:

- Direct owner fields on the artifact
- Microsoft Graph lookup by object ID
- Dataverse system user lookup (fallback path)

### OwnerObjectId
Raw owner identifier extracted from source systems.

Important:

- This may be a Microsoft Entra object ID.
- In some Agent/Dataverse paths, this can be a Dataverse systemuserid instead of an Entra object ID.

### DateCreated
Artifact creation timestamp in ISO 8601 format when available.

### ConnectionReferenceCount
Count of resolved connection references (or agent tool/action items when applicable).

### ConnectionReferences
Semicolon-delimited list of connection references (Flows/Apps) or tool/action labels (Agents) when -IncludeConnectionReferences is used.

## Troubleshooting

### Error: ThrowCrmSecurityException: The user with id

This error appears in the CSV output when the script encounters a security/permissions issue while attempting to retrieve artifact details or owner information from Dataverse.

**Symptoms:**
- CSV rows show "Error: ThrowCrmSecurityException: The user with id" in the ArtifactName column
- Owner, OwnerEmail, and OwnerObjectId columns are empty
- Typically appears for artifacts in specific environments or when retrieving owner details

**Root Causes:**

1. **Insufficient Permissions** — The executing account lacks the required Dataverse/Power Platform roles:
   - Missing System Administrator role in the environment
   - Not assigned Power Platform Administrator role
   - Missing Dynamics 365 Administrator role

2. **User Not Found** — The owner user ID stored in the artifact is invalid or has been deleted:
   - User was deleted from the directory after creating the artifact
   - User ID is corrupted in Dataverse
   - Cross-tenant scenario or identity mismatch

3. **User Disabled** — The owner account is disabled in Azure AD or Dataverse:
   - Account has been deprovisioned
   - Account is soft-deleted and not yet hard-deleted

4. **Environment Access Restrictions** — The executing account cannot access a specific environment:
   - Not added as a member of the environment
   - Environment has restricted access policies

**Resolutions:**

1. **Verify Your Permissions**
   - Confirm you are a **Power Platform Administrator** or have System Administrator role in all environments
   - Check that your account has **Global Administrator** or **Cloud Application Administrator** role in Azure AD
   - If using a service principal, verify it has the correct application permissions in Azure AD and is assigned the system user role in Dataverse

2. **Ensure Environment Access**
   - Verify your user account is added to each Power Platform environment being inventoried
   - Check that there are no environment security policies blocking your access

3. **Verify Owner Users Exist**
   - Investigate if the artifact owner user still exists in the directory
   - If the user was deleted, contact the environment admin to update or reassign the artifact ownership
   - Check Dataverse for orphaned user records

4. **Check Dataverse User Status**
   - Verify the owner user is active (not disabled) in the Dataverse environment
   - Re-enable or recreate the user if necessary

5. **Elevated Permissions May Be Required**
   - If you have appropriate admin roles and still see this error, the artifact may have been created with a user that no longer exists
   - Contact your Power Platform administrator to investigate the specific artifact

6. **Retry with Diagnostic Logging**
   - Consider running the script with verbose error logging to capture more details about which user ID is failing
   - Share the detailed error with your Power Platform or Dataverse admin team

## Notes

- Some fields can be blank due to API limitations, missing permissions, or unavailable enrichment paths.
- Agent collection may use PAC CLI fallback when no supported Agent PowerShell cmdlet is available.
- CSV outputs are intentionally ignored by git in this repository.
