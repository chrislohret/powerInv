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

## Notes

- Some fields can be blank due to API limitations, missing permissions, or unavailable enrichment paths.
- Agent collection may use PAC CLI fallback when no supported Agent PowerShell cmdlet is available.
- CSV outputs are intentionally ignored by git in this repository.
