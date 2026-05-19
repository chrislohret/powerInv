[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputCsvPath = (Join-Path -Path $PWD -ChildPath ("M365_PowerPlatform_Inventory_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss"))),

    [Parameter(Mandatory = $false)]
    [switch]$IncludeConnectionReferences,

    [Parameter(Mandatory = $false)]
    [switch]$SkipModuleInstall
)

<#
Script: Get-M365PowerPlatformInventory.ps1

Purpose:
- Collects Power Platform inventory across environments (Flows, Apps, and Agents)
- Exports a consolidated CSV with owner, creation date, and optional connection/tool metadata

How to run:
1. Default output in current directory:
    .\Get-M365PowerPlatformInventory.ps1

2. Include connection references/tools:
    .\Get-M365PowerPlatformInventory.ps1 -IncludeConnectionReferences

3. Write to a specific CSV path:
    .\Get-M365PowerPlatformInventory.ps1 -OutputCsvPath .\inventory.csv

4. Run without auto-installing missing modules:
    .\Get-M365PowerPlatformInventory.ps1 -SkipModuleInstall

Parameters:
- OutputCsvPath (string)
  Path to the CSV output file. Defaults to M365_PowerPlatform_Inventory_yyyyMMdd_HHmmss.csv in the current folder.

- IncludeConnectionReferences (switch)
  When set, attempts to include Flow/App connection references and Agent tool/action summaries.

- SkipModuleInstall (switch)
  When set, missing modules are not installed automatically and the script errors if required modules are absent.

PowerShell cmdlets and commands used:
- Module/bootstrap: Get-Module, Install-Module, Import-Module, Get-Command
- Power Platform auth/admin: Add-PowerAppsAccount, Get-AdminPowerAppEnvironment, Get-AdminEnvironment,
  Get-AdminFlow, Get-Flow, Get-AdminPowerApp, Get-AdminPowerAppConnectionReferences,
  Get-AdminPowerVirtualAgent, Get-AdminPowerVirtualAgents, Get-AdminCopilotStudioAgent, Get-AdminPowerAppCopilot
- Microsoft Graph (owner resolution): Get-MgContext, Connect-MgGraph, Get-MgUser
- Azure token acquisition for Dataverse: Get-AzContext, Connect-AzAccount, Get-AzAccessToken
- Dataverse API + export: Invoke-RestMethod, Export-Csv
- Background execution/timeouts for PAC calls: Start-Job, Wait-Job, Receive-Job, Stop-Job, Remove-Job
- External CLI fallback: pac (Power Platform CLI)
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$script:entraLookupInitialized = $false
$script:entraLookupEnabled = $false
$script:entraLookupFailureWarningShown = $false
$script:entraUserEmailCache = @{}

$script:dataverseTokenCache = @{}
$script:dataverseSystemUserEmailCache = @{}
$script:azAccountsAvailable = $false

# Ensures a required module is present and imported, with optional on-demand install.
function Ensure-Module {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $false)]
        [switch]$SkipInstall
    )

    if (-not (Get-Module -ListAvailable -Name $Name)) {
        if ($SkipInstall) {
            throw "Required module '$Name' is not installed. Install it first or run without -SkipModuleInstall."
        }

        Write-Host "Installing module '$Name' for current user..." -ForegroundColor Yellow
        Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber
    }

    Import-Module -Name $Name -ErrorAction Stop -WarningAction SilentlyContinue
}

# Returns the first non-empty value found by traversing one of the provided dotted paths.
function Get-ValueByPath {
    param(
        [Parameter(Mandatory = $true)]
        [object]$InputObject,

        [Parameter(Mandatory = $true)]
        [string[]]$Paths
    )

    foreach ($path in $Paths) {
        $segments = $path -split "\."
        $current = $InputObject
        $found = $true

        foreach ($segment in $segments) {
            if ($null -eq $current) {
                $found = $false
                break
            }

            if ($current -is [System.Collections.IDictionary] -and $current.Contains($segment)) {
                $current = $current[$segment]
                continue
            }

            $prop = $current.PSObject.Properties[$segment]
            if ($null -eq $prop) {
                $found = $false
                break
            }

            $current = $prop.Value
        }

        if ($found -and $null -ne $current -and -not [string]::IsNullOrWhiteSpace([string]$current)) {
            return $current
        }
    }

    return $null
}

# Normalizes raw connection-reference payloads into a count and semicolon-delimited label list.
function Convert-ToConnectionReferenceSummary {
    param(
        [Parameter(Mandatory = $false)]
        [object]$Entries
    )

    $labels = [System.Collections.Generic.List[string]]::new()

    $addLabel = {
        param([string]$Candidate)
        if ([string]::IsNullOrWhiteSpace($Candidate)) {
            return
        }

        if (-not $labels.Contains($Candidate)) {
            $labels.Add($Candidate) | Out-Null
        }
    }

    $buildLabel = {
        param(
            [object]$Entry,
            [string]$Fallback
        )

        $display = [string](Get-ValueByPath -InputObject $Entry -Paths @(
                "displayName",
                "connectionReferenceDisplayName",
                "properties.displayName",
                "Name",
                "name"
            ))

        $logicalName = [string](Get-ValueByPath -InputObject $Entry -Paths @(
                "connectionReferenceLogicalName",
                "logicalName",
                "ConnectionReferenceLogicalName",
                "id",
                "Id"
            ))

        $connector = [string](Get-ValueByPath -InputObject $Entry -Paths @(
                "api.name",
                "connectorName",
                "properties.apiName"
            ))

        if ([string]::IsNullOrWhiteSpace($display)) {
            if (-not [string]::IsNullOrWhiteSpace($logicalName)) {
                $display = $logicalName
            }
            elseif (-not [string]::IsNullOrWhiteSpace($Fallback)) {
                $display = $Fallback
            }
        }

        if ([string]::IsNullOrWhiteSpace($display)) {
            return $null
        }

        if (-not [string]::IsNullOrWhiteSpace($connector) -and ($display -notlike "*$connector*")) {
            return "$display ($connector)"
        }

        return $display
    }

    if ($null -eq $Entries) {
        return [PSCustomObject]@{
            Count = 0
            List  = $null
        }
    }

    $entryProperties = @($Entries.PSObject.Properties | Where-Object { $_.MemberType -in @("NoteProperty", "Property") })
    if ($entryProperties.Count -gt 0) {
        $singleEntryFields = @("displayName", "name", "id", "apiName", "connectionReferenceLogicalName")
        $isSingleEntry = $false
        foreach ($fieldName in $singleEntryFields) {
            if ($Entries.PSObject.Properties[$fieldName]) {
                $isSingleEntry = $true
                break
            }
        }

        if (-not $isSingleEntry) {
            foreach ($prop in $entryProperties) {
                if ($null -eq $prop.Value) {
                    continue
                }

                $label = & $buildLabel $prop.Value ([string]$prop.Name)
                & $addLabel $label
            }
        }
    }

    if ($Entries -is [System.Collections.IDictionary]) {
        foreach ($key in $Entries.Keys) {
            $entry = $Entries[$key]
            $label = & $buildLabel $entry ([string]$key)
            & $addLabel $label
        }
    }
    elseif ($Entries -is [System.Collections.IEnumerable] -and -not ($Entries -is [string])) {
        foreach ($entry in $Entries) {
            $label = & $buildLabel $entry $null
            & $addLabel $label
        }
    }
    else {
        $label = & $buildLabel $Entries $null
        & $addLabel $label
    }

    $listText = $null
    if ($labels.Count -gt 0) {
        $listText = [string]::Join("; ", $labels)
    }

    return [PSCustomObject]@{
        Count = $labels.Count
        List  = $listText
    }
}

# Extracts connection references directly from a Flow/App/Agent object when present.
function Get-ConnectionReferencesFromItem {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Item
    )

    $entries = Get-ValueByPath -InputObject $Item -Paths @(
        "ConnectionReferences",
        "connectionReferences",
        "properties.connectionReferences",
        "Internal.properties.connectionReferences"
    )

    return (Convert-ToConnectionReferenceSummary -Entries $entries)
}

# Fetches app connection references via cmdlet first, then falls back to object-level extraction.
function Get-AppConnectionReferences {
    param(
        [Parameter(Mandatory = $true)]
        [string]$EnvironmentId,

        [Parameter(Mandatory = $true)]
        [string]$AppId,

        [Parameter(Mandatory = $true)]
        [object]$AppItem
    )

    try {
        $references = Get-AdminPowerAppConnectionReferences -EnvironmentName $EnvironmentId -AppName $AppId
        $summary = Convert-ToConnectionReferenceSummary -Entries $references
        if ($summary.Count -gt 0) {
            return $summary
        }
    }
    catch {
    }

    return (Get-ConnectionReferencesFromItem -Item $AppItem)
}

# Fetches flow connection references via flow details first, then falls back to object-level extraction.
function Get-FlowConnectionReferences {
    param(
        [Parameter(Mandatory = $true)]
        [string]$EnvironmentId,

        [Parameter(Mandatory = $true)]
        [string]$FlowId,

        [Parameter(Mandatory = $true)]
        [object]$FlowItem
    )

    try {
        $flowDetails = Get-Flow -EnvironmentName $EnvironmentId -FlowName $FlowId
        if ($flowDetails -is [System.Collections.IEnumerable] -and -not ($flowDetails -is [string])) {
            $flowDetails = @($flowDetails) | Select-Object -First 1
        }

        if ($null -ne $flowDetails) {
            $summary = Convert-ToConnectionReferenceSummary -Entries (Get-ValueByPath -InputObject $flowDetails -Paths @(
                    "Internal.properties.connectionReferences",
                    "Internal.properties.definition.parameters.`$connections.defaultValue",
                    "ConnectionReferences",
                    "connectionReferences"
                ))

            if ($summary.Count -gt 0) {
                return $summary
            }
        }
    }
    catch {
    }

    return (Get-ConnectionReferencesFromItem -Item $FlowItem)
}

# Builds an agent tool/action summary by extracting a PAC template and parsing connectors/actions.
function Get-AgentToolSummary {
    param(
        [Parameter(Mandatory = $false)]
        [string]$EnvironmentUrl,

        [Parameter(Mandatory = $false)]
        [string]$AgentId
    )

    if ([string]::IsNullOrWhiteSpace($EnvironmentUrl) -or [string]::IsNullOrWhiteSpace($AgentId)) {
        return [PSCustomObject]@{
            Count = 0
            List  = $null
        }
    }

    if ($null -eq (Get-PacCommand)) {
        return [PSCustomObject]@{
            Count = 0
            List  = $null
        }
    }

    $tempFolder = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("pp_agent_tools_{0}" -f ([guid]::NewGuid().ToString("N")))
    $templateYaml = Join-Path -Path $tempFolder -ChildPath "agent_template.yaml"

    New-Item -Path $tempFolder -ItemType Directory -Force | Out-Null

    try {
        $extractOutput = Invoke-PacCommandSafe -Arguments @(
            "copilot",
            "extract-template",
            "--environment",
            $EnvironmentUrl,
            "--bot",
            $AgentId,
            "--templateFileName",
            $templateYaml,
            "--overwrite"
        ) -TimeoutSeconds 90

        if ($null -eq $extractOutput -or -not (Test-Path -Path $templateYaml)) {
            return [PSCustomObject]@{
                Count = 0
                List  = $null
            }
        }

        $toolLabels = [System.Collections.Generic.List[string]]::new()

        $addTool = {
            param([string]$Value)
            if ([string]::IsNullOrWhiteSpace($Value)) {
                return
            }

            if (-not $toolLabels.Contains($Value)) {
                $toolLabels.Add($Value) | Out-Null
            }
        }

        $templateJsonFile = Get-ChildItem -Path $tempFolder -Filter "*.json" -ErrorAction SilentlyContinue |
            Sort-Object -Property Name |
            Select-Object -First 1

        if ($null -ne $templateJsonFile) {
            try {
                $templateJson = Get-Content -Path $templateJsonFile.FullName -Raw | ConvertFrom-Json
                $connectors = @($templateJson.spec.connectors)

                foreach ($connector in $connectors) {
                    if ($connector -is [string]) {
                        & $addTool ("connector:{0}" -f $connector)
                        continue
                    }

                    $connectorName = [string](Get-ValueByPath -InputObject $connector -Paths @(
                            "displayName",
                            "name",
                            "id",
                            "connectorId"
                        ))

                    if (-not [string]::IsNullOrWhiteSpace($connectorName)) {
                        & $addTool ("connector:{0}" -f $connectorName)
                    }
                }
            }
            catch {
            }
        }

        try {
            $yamlText = Get-Content -Path $templateYaml -Raw
            $matches = [regex]::Matches($yamlText, "(?m)^\s*-\s*kind:\s*([A-Za-z0-9_]+)\s*$")
            foreach ($match in $matches) {
                $kind = [string]$match.Groups[1].Value
                if ([string]::IsNullOrWhiteSpace($kind)) {
                    continue
                }

                if ($kind -match "Invoke|Connector|Flow|Plugin|Mcp|Http|Search|Knowledge|Action") {
                    & $addTool ("action:{0}" -f $kind)
                }
            }
        }
        catch {
        }

        $listText = $null
        if ($toolLabels.Count -gt 0) {
            $listText = [string]::Join("; ", $toolLabels)
        }

        return [PSCustomObject]@{
            Count = $toolLabels.Count
            List  = $listText
        }
    }
    finally {
        Remove-Item -Path $tempFolder -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# Converts a date-like value into ISO-8601 text while preserving non-date values as strings.
function Convert-ToIsoDate {
    param(
        [Parameter(Mandatory = $false)]
        [object]$Value
    )

    if ($null -eq $Value) {
        return $null
    }

    try {
        return ([datetime]$Value).ToString("o")
    }
    catch {
        return [string]$Value
    }
}

# Resolves owner email/UPN candidates across known schema variants.
function Resolve-OwnerEmail {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Item
    )

    $candidate = Get-ValueByPath -InputObject $Item -Paths @(
        "Owner.email",
        "Owner.userPrincipalName",
        "Owner.upn",
        "properties.owner.email",
        "properties.owner.userPrincipalName",
        "properties.creator.email",
        "properties.creator.userPrincipalName",
        "Creator.email",
        "Creator.userPrincipalName",
        "CreatedBy.email",
        "CreatedBy.userPrincipalName",
        "UserDetails.email",
        "UserDetails.userPrincipalName"
    )

    return $candidate
}

# Resolves owner display name candidates across known schema variants.
function Resolve-OwnerDisplayName {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Item
    )

    return (Get-ValueByPath -InputObject $Item -Paths @(
            "Owner.displayName",
            "Owner.name",
            "properties.owner.displayName",
            "properties.creator.displayName",
            "Creator.displayName",
            "CreatedBy.displayName",
            "UserDetails.displayName"
        ))
}

# Extracts owner object IDs from multiple common payload shapes for flows/apps/agents.
function Get-OwnerObjectIdFromItem {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Item
    )

    return [string](Get-ValueByPath -InputObject $Item -Paths @(
            "Owner.id",
            "Owner.objectId",
            "Owner.userId",
            "Owner.aadObjectId",
            "owner.id",
            "owner.objectId",
            "owner.userId",
            "owner.aadObjectId",
            "CreatedBy.id",
            "CreatedBy.objectId",
            "CreatedBy.userId",
            "CreatedBy.aadObjectId",
            "createdBy.id",
            "createdBy.objectId",
            "createdBy.userId",
            "createdBy.aadObjectId",
            "Creator.id",
            "Creator.objectId",
            "Creator.userId",
            "creator.id",
            "creator.objectId",
            "creator.userId",
            "properties.owner.id",
            "properties.owner.objectId",
            "properties.owner.userId",
            "properties.creator.id",
            "properties.creator.objectId",
            "properties.creator.userId",
            "properties.createdBy.id",
            "properties.createdBy.objectId",
            "properties.createdBy.userId",
            "UserDetails.id",
            "userDetails.id",
            "_ownerid_value",
            "ownerid"
        ))
}

# Formats owner output as "Display <Email>" when both are available.
function Resolve-OwnerField {
    param(
        [Parameter(Mandatory = $false)]
        [string]$OwnerDisplayName,

        [Parameter(Mandatory = $false)]
        [string]$OwnerEmail
    )

    if (-not [string]::IsNullOrWhiteSpace($OwnerDisplayName) -and -not [string]::IsNullOrWhiteSpace($OwnerEmail) -and $OwnerDisplayName -ne $OwnerEmail) {
        return "$OwnerDisplayName <$OwnerEmail>"
    }

    if (-not [string]::IsNullOrWhiteSpace($OwnerEmail)) {
        return $OwnerEmail
    }

    return $OwnerDisplayName
}

# Pulls a GUID out of an input value, including free-form strings that contain GUID text.
function Get-GuidFromValue {
    param(
        [Parameter(Mandatory = $false)]
        [object]$Value
    )

    if ($null -eq $Value) {
        return $null
    }

    if ($Value -is [guid]) {
        return ([guid]$Value).ToString()
    }

    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    if ($text -match "(?i)[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}") {
        return $matches[0].ToLowerInvariant()
    }

    return $null
}

# Extracts an agent owner GUID from known owner/creator fields and fallback object blobs.
function Get-AgentOwnerGuidFromItem {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Item
    )

    $ownerCandidate = Get-ValueByPath -InputObject $Item -Paths @(
        "Owner.id",
        "Owner.objectId",
        "Owner.aadObjectId",
        "Owner.userId",
        "Owner.principalId",
        "owner.id",
        "owner.objectId",
        "owner.aadObjectId",
        "owner.userId",
        "owner.principalId",
        "ownerObjectId",
        "OwnerObjectId",
        "OwnerId",
        "ownerId",
        "_ownerid_value",
        "properties.owner.id",
        "properties.owner.objectId",
        "properties.owner.aadObjectId",
        "properties.owner.userId",
        "properties.owner.principalId",
        "properties.ownerObjectId",
        "properties.OwnerObjectId",
        "properties.OwnerId",
        "properties.ownerId",
        "properties._ownerid_value",
        "CreatedBy.id",
        "CreatedBy.objectId",
        "CreatedBy.aadObjectId",
        "createdBy.id",
        "createdBy.objectId",
        "createdBy.aadObjectId",
        "properties.createdBy.id",
        "properties.createdBy.objectId",
        "properties.creator.id",
        "properties.creator.objectId",
        "Creator.id",
        "Creator.objectId",
        "creator.id",
        "creator.objectId",
        "UserDetails.id",
        "userDetails.id"
    )

    $ownerGuid = Get-GuidFromValue -Value $ownerCandidate
    if (-not [string]::IsNullOrWhiteSpace($ownerGuid)) {
        return $ownerGuid
    }

    # Some cmdlets expose owner metadata as free-form strings (for example, resource IDs).
    foreach ($path in @("Owner", "owner", "CreatedBy", "createdBy", "properties.owner", "properties.createdBy")) {
        $candidateValue = Get-ValueByPath -InputObject $Item -Paths @($path)
        $parsedGuid = Get-GuidFromValue -Value $candidateValue
        if (-not [string]::IsNullOrWhiteSpace($parsedGuid)) {
            return $parsedGuid
        }
    }

    return $null
}

# Initializes Microsoft Graph modules/session used for Entra user lookups.
function Initialize-EntraLookup {
    if ($script:entraLookupInitialized) {
        return
    }

    $script:entraLookupInitialized = $true

    try {
        Ensure-Module -Name "Microsoft.Graph.Authentication" -SkipInstall:$SkipModuleInstall
        Ensure-Module -Name "Microsoft.Graph.Users" -SkipInstall:$SkipModuleInstall

        $context = Get-MgContext -ErrorAction SilentlyContinue
        if ($null -eq $context) {
            Connect-MgGraph -Scopes "User.Read.All" -NoWelcome | Out-Null
        }

        $script:entraLookupEnabled = $true
        Write-Host "Entra owner lookup enabled for agents (Microsoft Graph)." -ForegroundColor DarkCyan
    }
    catch {
        $script:entraLookupEnabled = $false
        if (-not $script:entraLookupFailureWarningShown) {
            Write-Warning "Could not initialize Entra owner lookup. Agent owner emails will remain unresolved when only GUIDs are available. $($_.Exception.Message)"
            $script:entraLookupFailureWarningShown = $true
        }
    }
}

# Resolves an Entra email/UPN from an object ID with in-memory caching.
function Resolve-EntraEmailFromObjectId {
    param(
        [Parameter(Mandatory = $false)]
        [string]$ObjectId
    )

    if ([string]::IsNullOrWhiteSpace($ObjectId)) {
        return $null
    }

    if ($script:entraUserEmailCache.ContainsKey($ObjectId)) {
        return $script:entraUserEmailCache[$ObjectId]
    }

    Initialize-EntraLookup
    if (-not $script:entraLookupEnabled) {
        $script:entraUserEmailCache[$ObjectId] = $null
        return $null
    }

    try {
        $user = Get-MgUser -UserId $ObjectId -Property Id,Mail,UserPrincipalName -ErrorAction Stop
        $resolvedEmail = [string]$user.Mail
        if ([string]::IsNullOrWhiteSpace($resolvedEmail)) {
            $resolvedEmail = [string]$user.UserPrincipalName
        }

        $script:entraUserEmailCache[$ObjectId] = $resolvedEmail
        return $resolvedEmail
    }
    catch {
        $script:entraUserEmailCache[$ObjectId] = $null
        return $null
    }
}

# Safely converts secure tokens/secrets into plaintext for HTTP authorization headers.
function Convert-SecretToPlainText {
    param(
        [Parameter(Mandatory = $false)]
        [object]$SecretValue
    )

    if ($null -eq $SecretValue) {
        return $null
    }

    if ($SecretValue -is [Security.SecureString]) {
        $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecretValue)
        try {
            return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
        }
        finally {
            if ($bstr -ne [IntPtr]::Zero) {
                [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
            }
        }
    }

    return [string]$SecretValue
}

# Gets and caches an access token for Dataverse API calls scoped to an instance URL.
function Get-DataverseAccessToken {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ResourceUrl
    )

    $normalizedUrl = $ResourceUrl.TrimEnd('/')

    if ($script:dataverseTokenCache.ContainsKey($normalizedUrl)) {
        return $script:dataverseTokenCache[$normalizedUrl]
    }

    if ($script:azAccountsAvailable) {
        try {
            $result = Get-AzAccessToken -ResourceUrl $normalizedUrl -ErrorAction Stop
            $token = Convert-SecretToPlainText -SecretValue $result.Token
            if ([string]::IsNullOrWhiteSpace($token) -and $result.PSObject.Properties["AccessToken"]) {
                $token = Convert-SecretToPlainText -SecretValue $result.AccessToken
            }
            if (-not [string]::IsNullOrWhiteSpace($token)) {
                $script:dataverseTokenCache[$normalizedUrl] = $token
                return $token
            }
        }
        catch { }
    }

    $script:dataverseTokenCache[$normalizedUrl] = $null
    return $null
}

# Queries Dataverse bots and returns a map of bot ID to owner/date metadata.
function Get-DataverseBotOwnerMap {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InstanceUrl
    )

    $map = @{}
    $normalizedUrl = $InstanceUrl.TrimEnd('/')
    $token = Get-DataverseAccessToken -ResourceUrl $normalizedUrl

    if ([string]::IsNullOrWhiteSpace($token)) {
        return $map
    }

    try {
        $headers = @{
            "Authorization"    = "Bearer $token"
            "Accept"           = "application/json"
            "OData-MaxVersion" = "4.0"
            "OData-Version"    = "4.0"
            "Prefer"           = "odata.include-annotations=`"OData.Community.Display.V1.FormattedValue`""
        }

        $uri = "$normalizedUrl/api/data/v9.2/bots?`$select=botid,name,_ownerid_value,createdon&`$top=5000"
        $response = Invoke-RestMethod -Uri $uri -Headers $headers -Method GET -ErrorAction Stop

        foreach ($bot in $response.value) {
            $botId = ([string]$bot.botid).ToLowerInvariant()
            if (-not [string]::IsNullOrWhiteSpace($botId)) {
                $ownerDisplayValue = $bot.PSObject.Properties["_ownerid_value@OData.Community.Display.V1.FormattedValue"]
                $map[$botId] = @{
                    OwnerSystemUserId = [string]$bot."_ownerid_value"
                    OwnerDisplayName  = if ($null -ne $ownerDisplayValue) { [string]$ownerDisplayValue.Value } else { $null }
                    CreatedOn         = [string]$bot.createdon
                }
            }
        }
    }
    catch { }

    return $map
}

# Resolves a Dataverse system user email and caches results by environment+user key.
function Resolve-DataverseSystemUserEmail {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InstanceUrl,

        [Parameter(Mandatory = $true)]
        [string]$SystemUserId
    )

    if ([string]::IsNullOrWhiteSpace($SystemUserId)) {
        return $null
    }

    $cacheKey = "$InstanceUrl|$SystemUserId"
    if ($script:dataverseSystemUserEmailCache.ContainsKey($cacheKey)) {
        return $script:dataverseSystemUserEmailCache[$cacheKey]
    }

    $normalizedUrl = $InstanceUrl.TrimEnd('/')
    $token = Get-DataverseAccessToken -ResourceUrl $normalizedUrl

    if ([string]::IsNullOrWhiteSpace($token)) {
        $script:dataverseSystemUserEmailCache[$cacheKey] = $null
        return $null
    }

    try {
        $headers = @{
            "Authorization"    = "Bearer $token"
            "Accept"           = "application/json"
            "OData-MaxVersion" = "4.0"
            "OData-Version"    = "4.0"
        }

        $uri = "$normalizedUrl/api/data/v9.2/systemusers($SystemUserId)?`$select=internalemailaddress,fullname"
        $user = Invoke-RestMethod -Uri $uri -Headers $headers -Method GET -ErrorAction Stop
        $email = [string]$user.internalemailaddress
        $script:dataverseSystemUserEmailCache[$cacheKey] = $email
        return $email
    }
    catch {
        $script:dataverseSystemUserEmailCache[$cacheKey] = $null
        return $null
    }
}

# Finds the first supported PowerShell cmdlet for retrieving Copilot Studio agents.
function Get-AgentCommand {
    $candidateCommands = @(
        "Get-AdminPowerVirtualAgent",
        "Get-AdminPowerVirtualAgents",
        "Get-AdminCopilotStudioAgent",
        "Get-AdminPowerAppCopilot"
    )

    foreach ($commandName in $candidateCommands) {
        $command = Get-Command -Name $commandName -ErrorAction SilentlyContinue
        if ($null -ne $command) {
            return $command
        }
    }

    return $null
}

# Returns the PAC CLI command if available in PATH.
function Get-PacCommand {
    return (Get-Command -Name "pac" -ErrorAction SilentlyContinue)
}

# Runs a PAC command in a background job and enforces a timeout to avoid hangs.
function Invoke-PacCommandSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Arguments,

        [Parameter(Mandatory = $false)]
        [int]$TimeoutSeconds = 45
    )

    $job = Start-Job -ScriptBlock {
        param([string[]]$InnerArguments)
        pac @InnerArguments 2>&1 | Out-String
    } -ArgumentList (, $Arguments)

    try {
        $completed = Wait-Job -Job $job -Timeout $TimeoutSeconds
        if ($null -eq $completed) {
            Stop-Job -Job $job -Force | Out-Null
            return $null
        }

        return [string](Receive-Job -Job $job)
    }
    finally {
        Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
    }
}

# Parses tabular `pac copilot list` output into structured name/id records.
function Parse-PacCopilotListOutput {
    param(
        [Parameter(Mandatory = $false)]
        [string]$OutputText
    )

    if ([string]::IsNullOrWhiteSpace($OutputText)) {
        return @()
    }

    $records = [System.Collections.Generic.List[object]]::new()
    $lines = $OutputText -split "`r?`n"
    $guidPattern = "[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}"

    foreach ($line in $lines) {
        $trimmed = $line.TrimEnd()
        if ([string]::IsNullOrWhiteSpace($trimmed)) {
            continue
        }

        if ($trimmed -match "^(Connected as|Connected to\.\.\.|Microsoft PowerPlatform CLI|Version:|Online documentation:|Feedback, Suggestions, Issues:|Name\s+Copilot ID|---+)") {
            continue
        }

        $idMatch = [regex]::Match($trimmed, $guidPattern)
        if (-not $idMatch.Success) {
            continue
        }

        $copilotId = $idMatch.Value
        $name = $trimmed.Substring(0, $idMatch.Index).Trim()

        $records.Add([PSCustomObject]@{
                Name = $name
                CopilotId = $copilotId
            }) | Out-Null
    }

    return $records
}

# Finds the first supported PowerShell cmdlet for listing environments.
function Get-EnvironmentCommand {
    $candidateCommands = @(
        "Get-AdminPowerAppEnvironment",
        "Get-AdminEnvironment"
    )

    foreach ($commandName in $candidateCommands) {
        $command = Get-Command -Name $commandName -ErrorAction SilentlyContinue
        if ($null -ne $command) {
            return $command
        }
    }

    return $null
}

# Adds a normalized artifact row into the in-memory inventory list.
function Add-InventoryRecord {
    param(
        [Parameter(Mandatory = $false)]
        [ValidateNotNull()]
        [System.Collections.Generic.List[object]]$Buffer,

        [Parameter(Mandatory = $true)]
        [string]$ArtifactType,

        [Parameter(Mandatory = $true)]
        [string]$Environment,

        [Parameter(Mandatory = $true)]
        [string]$EnvironmentId,

        [Parameter(Mandatory = $false)]
        [string]$Guid,

        [Parameter(Mandatory = $false)]
        [string]$ArtifactName,

        [Parameter(Mandatory = $false)]
        [string]$Owner,

        [Parameter(Mandatory = $false)]
        [string]$OwnerEmail,

        [Parameter(Mandatory = $false)]
        [string]$OwnerObjectId,

        [Parameter(Mandatory = $false)]
        [string]$DateCreated,

        [Parameter(Mandatory = $false)]
        [string]$ConnectionReferences,

        [Parameter(Mandatory = $false)]
        [int]$ConnectionReferenceCount = 0
    )

    if ($null -eq $Buffer) {
        throw "Buffer cannot be null."
    }

    $ownerToWrite = $Owner
    if ($ArtifactType -eq "Agent" -and -not [string]::IsNullOrWhiteSpace($OwnerObjectId)) {
        $ownerToWrite = $OwnerObjectId
    }

    $Buffer.Add([PSCustomObject]@{
            ArtifactType = $ArtifactType
            Environment  = $Environment
            EnvironmentId = $EnvironmentId
            GUID         = $Guid
            ArtifactName = $ArtifactName
            Owner        = $ownerToWrite
            OwnerEmail   = $OwnerEmail
            OwnerObjectId = $OwnerObjectId
            DateCreated  = $DateCreated
            ConnectionReferenceCount = $ConnectionReferenceCount
            ConnectionReferences = $ConnectionReferences
        }) | Out-Null
}

# Load required Power Platform modules before any data collection.
Ensure-Module -Name "Microsoft.PowerApps.Administration.PowerShell" -SkipInstall:$SkipModuleInstall
Ensure-Module -Name "Microsoft.PowerApps.PowerShell" -SkipInstall:$SkipModuleInstall

# Prepare Az context for Dataverse token-based enrichment (PAC fallback path).
try {
    Ensure-Module -Name "Az.Accounts" -SkipInstall:$SkipModuleInstall
    $azContext = Get-AzContext -ErrorAction SilentlyContinue
    if ($null -eq $azContext) {
        Write-Host "Signing in to Azure (Az.Accounts) for Dataverse token acquisition..." -ForegroundColor Cyan
        Write-Host "A device code will be displayed. Visit https://microsoft.com/devicelogin in your browser and enter the code." -ForegroundColor Yellow
        Connect-AzAccount -UseDeviceAuthentication -ErrorAction Stop | Out-Null
    }
    $script:azAccountsAvailable = $true
    Write-Host "Az.Accounts ready — Dataverse owner/date enrichment enabled for PAC fallback agents." -ForegroundColor DarkCyan
}
catch {
    Write-Warning "Az.Accounts module could not be loaded or authenticated. PAC fallback agents will have blank Owner/OwnerEmail/DateCreated. ($($_.Exception.Message))"
    $script:azAccountsAvailable = $false
}

Write-Host "Signing in to Power Platform..." -ForegroundColor Cyan
Add-PowerAppsAccount | Out-Null

# Discover tenant environments and cache environment URL lookups.
Write-Host "Collecting environments..." -ForegroundColor Cyan
$environmentCommand = Get-EnvironmentCommand
if ($null -eq $environmentCommand) {
    throw "No supported environment command was found. Expected one of: Get-AdminPowerAppEnvironment, Get-AdminEnvironment."
}

Write-Host "Using environment command: $($environmentCommand.Name)" -ForegroundColor DarkCyan
$environments = & $environmentCommand.Name

if (-not $environments) {
    throw "No environments found or you do not have sufficient permissions."
}

$inventory = [System.Collections.Generic.List[object]]::new()
$flowCount = 0
$appCount = 0
$agentCount = 0
$environmentUrlById = @{}

foreach ($environment in $environments) {
    $mappedEnvironmentId = [string](Get-ValueByPath -InputObject $environment -Paths @("EnvironmentName", "Name", "Id"))
    $mappedInstanceUrl = [string](Get-ValueByPath -InputObject $environment -Paths @("Internal.properties.linkedEnvironmentMetadata.instanceUrl"))
    if (-not [string]::IsNullOrWhiteSpace($mappedEnvironmentId)) {
        $environmentUrlById[$mappedEnvironmentId] = $mappedInstanceUrl
    }
}

$agentCommand = Get-AgentCommand
$pacCommand = Get-PacCommand
$agentCommandName = $null
$agentSupportsEnvironment = $false
$agentEnvParameter = $null
$usePacFallbackForAgents = $false
$pacFallbackOwnerDateMissingWarningShown = $false

# Decide whether to use native agent cmdlets or PAC CLI fallback.
if ($null -ne $agentCommand) {
    $agentCommandName = $agentCommand.Name
    foreach ($parameterName in @("EnvironmentName", "EnvironmentId")) {
        if ($agentCommand.Parameters.ContainsKey($parameterName)) {
            $agentSupportsEnvironment = $true
            $agentEnvParameter = $parameterName
            break
        }
    }

    Write-Host "Using agent command: $agentCommandName" -ForegroundColor DarkCyan
}
else {
    if ($null -ne $pacCommand) {
        $usePacFallbackForAgents = $true
        Write-Host "No supported PowerShell agent cmdlet was found. Using 'pac copilot list' for agent inventory." -ForegroundColor DarkYellow
    }
    else {
        Write-Warning "No supported agent command was found (PowerShell or PAC CLI). Agent inventory will be skipped."
    }
}

# Iterate environments and collect Flow/App/Agent artifacts.
foreach ($environment in $environments) {
    $environmentId = [string](Get-ValueByPath -InputObject $environment -Paths @("EnvironmentName", "Name", "Id"))
    $environmentDisplayName = [string](Get-ValueByPath -InputObject $environment -Paths @("DisplayName", "EnvironmentDisplayName"))
    $environmentInstanceUrl = [string](Get-ValueByPath -InputObject $environment -Paths @("Internal.properties.linkedEnvironmentMetadata.instanceUrl"))

    if ([string]::IsNullOrWhiteSpace($environmentDisplayName)) {
        $environmentDisplayName = $environmentId
    }

    Write-Host "Processing environment: $environmentDisplayName" -ForegroundColor Gray

    try {
        $flows = Get-AdminFlow -EnvironmentName $environmentId

        foreach ($flow in $flows) {
            $ownerEmail = [string](Resolve-OwnerEmail -Item $flow)
            $ownerDisplay = [string](Resolve-OwnerDisplayName -Item $flow)
            $ownerObjectId = [string](Get-OwnerObjectIdFromItem -Item $flow)

            if ([string]::IsNullOrWhiteSpace($ownerEmail) -and -not [string]::IsNullOrWhiteSpace($ownerObjectId)) {
                $ownerEmail = [string](Resolve-EntraEmailFromObjectId -ObjectId $ownerObjectId)

                if ([string]::IsNullOrWhiteSpace($ownerEmail) -and -not [string]::IsNullOrWhiteSpace($environmentInstanceUrl)) {
                    $ownerEmail = [string](Resolve-DataverseSystemUserEmail -InstanceUrl $environmentInstanceUrl -SystemUserId $ownerObjectId)
                }
            }

            $ownerValue = Resolve-OwnerField -OwnerDisplayName $ownerDisplay -OwnerEmail $ownerEmail
            if ([string]::IsNullOrWhiteSpace($ownerValue) -and -not [string]::IsNullOrWhiteSpace($ownerObjectId)) {
                $ownerValue = $ownerObjectId
            }

            $guid = [string](Get-ValueByPath -InputObject $flow -Paths @("FlowName", "Name", "Id", "properties.workflowEntityId"))
            $artifactName = [string](Get-ValueByPath -InputObject $flow -Paths @("DisplayName", "FlowDisplayName", "properties.displayName", "Name"))
            $created = Convert-ToIsoDate (Get-ValueByPath -InputObject $flow -Paths @("CreatedTime", "CreatedDateTime", "properties.createdTime", "properties.createdDateTime"))
            $connectionReferences = $null
            $connectionReferenceCount = 0

            if ($IncludeConnectionReferences) {
                $flowRefSummary = Get-FlowConnectionReferences -EnvironmentId $environmentId -FlowId $guid -FlowItem $flow
                $connectionReferences = [string]$flowRefSummary.List
                $connectionReferenceCount = [int]$flowRefSummary.Count
            }

            Add-InventoryRecord -Buffer $inventory -ArtifactType "Flow" -Environment $environmentDisplayName -EnvironmentId $environmentId -Guid $guid -ArtifactName $artifactName -Owner $ownerValue -OwnerEmail $ownerEmail -OwnerObjectId $ownerObjectId -DateCreated $created -ConnectionReferences $connectionReferences -ConnectionReferenceCount $connectionReferenceCount
            $flowCount++
        }
    }
    catch {
        Write-Warning "Failed to collect flows for environment '$environmentDisplayName'. $($_.Exception.Message)"
    }

    try {
        $apps = Get-AdminPowerApp -EnvironmentName $environmentId

        foreach ($app in $apps) {
            $ownerEmail = [string](Resolve-OwnerEmail -Item $app)
            $ownerDisplay = [string](Resolve-OwnerDisplayName -Item $app)
            $ownerObjectId = [string](Get-OwnerObjectIdFromItem -Item $app)

            if ([string]::IsNullOrWhiteSpace($ownerEmail) -and -not [string]::IsNullOrWhiteSpace($ownerObjectId)) {
                $ownerEmail = [string](Resolve-EntraEmailFromObjectId -ObjectId $ownerObjectId)

                if ([string]::IsNullOrWhiteSpace($ownerEmail) -and -not [string]::IsNullOrWhiteSpace($environmentInstanceUrl)) {
                    $ownerEmail = [string](Resolve-DataverseSystemUserEmail -InstanceUrl $environmentInstanceUrl -SystemUserId $ownerObjectId)
                }
            }

            $ownerValue = Resolve-OwnerField -OwnerDisplayName $ownerDisplay -OwnerEmail $ownerEmail
            if ([string]::IsNullOrWhiteSpace($ownerValue) -and -not [string]::IsNullOrWhiteSpace($ownerObjectId)) {
                $ownerValue = $ownerObjectId
            }

            $guid = [string](Get-ValueByPath -InputObject $app -Paths @("AppName", "Name", "Id"))
            $artifactName = [string](Get-ValueByPath -InputObject $app -Paths @("DisplayName", "AppDisplayName", "properties.displayName", "Name"))
            $created = Convert-ToIsoDate (Get-ValueByPath -InputObject $app -Paths @("CreatedTime", "CreatedDateTime", "properties.createdTime", "properties.createdDateTime"))
            $connectionReferences = $null
            $connectionReferenceCount = 0

            if ($IncludeConnectionReferences) {
                $appRefSummary = Get-AppConnectionReferences -EnvironmentId $environmentId -AppId $guid -AppItem $app
                $connectionReferences = [string]$appRefSummary.List
                $connectionReferenceCount = [int]$appRefSummary.Count
            }

            Add-InventoryRecord -Buffer $inventory -ArtifactType "App" -Environment $environmentDisplayName -EnvironmentId $environmentId -Guid $guid -ArtifactName $artifactName -Owner $ownerValue -OwnerEmail $ownerEmail -OwnerObjectId $ownerObjectId -DateCreated $created -ConnectionReferences $connectionReferences -ConnectionReferenceCount $connectionReferenceCount
            $appCount++
        }
    }
    catch {
        Write-Warning "Failed to collect apps for environment '$environmentDisplayName'. $($_.Exception.Message)"
    }

    if ($null -ne $agentCommand -and $agentSupportsEnvironment) {
        try {
            $agentArgs = @{}
            $agentArgs[$agentEnvParameter] = $environmentId
            $agents = & $agentCommandName @agentArgs

            foreach ($agent in $agents) {
                $ownerEmail = [string](Resolve-OwnerEmail -Item $agent)
                $ownerDisplay = [string](Resolve-OwnerDisplayName -Item $agent)
                $ownerGuid = Get-AgentOwnerGuidFromItem -Item $agent

                if ([string]::IsNullOrWhiteSpace($ownerEmail) -and -not [string]::IsNullOrWhiteSpace($ownerGuid)) {
                    $ownerEmail = [string](Resolve-EntraEmailFromObjectId -ObjectId $ownerGuid)
                }

                $ownerValue = Resolve-OwnerField -OwnerDisplayName $ownerDisplay -OwnerEmail $ownerEmail
                if ([string]::IsNullOrWhiteSpace($ownerValue) -and -not [string]::IsNullOrWhiteSpace($ownerGuid)) {
                    $ownerValue = $ownerGuid
                }

                $guid = [string](Get-ValueByPath -InputObject $agent -Paths @("BotId", "CopilotId", "AgentId", "Name", "Id"))
                $artifactName = [string](Get-ValueByPath -InputObject $agent -Paths @("DisplayName", "Name", "properties.displayName", "BotName"))
                $created = Convert-ToIsoDate (Get-ValueByPath -InputObject $agent -Paths @("CreatedTime", "CreatedDateTime", "properties.createdTime", "properties.createdDateTime", "CreatedOn"))
                $connectionReferences = $null
                $connectionReferenceCount = 0

                if ($IncludeConnectionReferences) {
                    $agentRefSummary = Get-ConnectionReferencesFromItem -Item $agent
                    if ($agentRefSummary.Count -gt 0) {
                        $connectionReferences = [string]$agentRefSummary.List
                        $connectionReferenceCount = [int]$agentRefSummary.Count
                    }
                    else {
                        $agentToolSummary = Get-AgentToolSummary -EnvironmentUrl $environmentInstanceUrl -AgentId $guid
                        $connectionReferences = [string]$agentToolSummary.List
                        $connectionReferenceCount = [int]$agentToolSummary.Count
                    }
                }

                Add-InventoryRecord -Buffer $inventory -ArtifactType "Agent" -Environment $environmentDisplayName -EnvironmentId $environmentId -Guid $guid -ArtifactName $artifactName -Owner $ownerValue -OwnerEmail $ownerEmail -OwnerObjectId $ownerGuid -DateCreated $created -ConnectionReferences $connectionReferences -ConnectionReferenceCount $connectionReferenceCount
                $agentCount++
            }
        }
        catch {
            Write-Warning "Failed to collect agents for environment '$environmentDisplayName'. $($_.Exception.Message)"
        }
    }

    if ($usePacFallbackForAgents) {
        try {
            $instanceUrl = [string](Get-ValueByPath -InputObject $environment -Paths @("Internal.properties.linkedEnvironmentMetadata.instanceUrl"))
            if ([string]::IsNullOrWhiteSpace($instanceUrl)) {
                continue
            }

            $pacOutput = Invoke-PacCommandSafe -Arguments @("copilot", "list", "--environment", $instanceUrl) -TimeoutSeconds 45
            if ($null -eq $pacOutput) {
                Write-Warning "PAC agent query timed out for environment '$environmentDisplayName'."
                continue
            }

            $pacAgents = @(Parse-PacCopilotListOutput -OutputText $pacOutput)

            # Attempt to enrich owner/date via Dataverse Web API (requires Az.Accounts for token acquisition)
            $dvBotOwnerMap = Get-DataverseBotOwnerMap -InstanceUrl $instanceUrl
            $dvEnrichmentAvailable = $dvBotOwnerMap.Count -gt 0

            foreach ($agent in $pacAgents) {
                $agentGuid = [string]$agent.CopilotId
                $ownerEmail = $null
                $ownerDisplayName = $null
                $ownerObjectId = $null
                $created = $null

                if ($dvEnrichmentAvailable) {
                    $dvEntry = $dvBotOwnerMap[$agentGuid.ToLowerInvariant()]
                    if ($null -ne $dvEntry) {
                        $ownerDisplayName = [string]$dvEntry.OwnerDisplayName
                        $created = Convert-ToIsoDate $dvEntry.CreatedOn

                        $ownerObjectId = [string]$dvEntry.OwnerSystemUserId
                        if (-not [string]::IsNullOrWhiteSpace($ownerObjectId)) {
                            $ownerEmail = Resolve-DataverseSystemUserEmail -InstanceUrl $instanceUrl -SystemUserId $ownerObjectId
                        }
                    }
                }

                $ownerValue = Resolve-OwnerField -OwnerDisplayName $ownerDisplayName -OwnerEmail $ownerEmail
                if ([string]::IsNullOrWhiteSpace($ownerValue) -and -not [string]::IsNullOrWhiteSpace($ownerObjectId)) {
                    $ownerValue = $ownerObjectId
                }

                $connectionReferences = $null
                $connectionReferenceCount = 0

                if ($IncludeConnectionReferences) {
                    $agentToolSummary = Get-AgentToolSummary -EnvironmentUrl $instanceUrl -AgentId $agentGuid
                    $connectionReferences = [string]$agentToolSummary.List
                    $connectionReferenceCount = [int]$agentToolSummary.Count
                }

                Add-InventoryRecord -Buffer $inventory -ArtifactType "Agent" -Environment $environmentDisplayName -EnvironmentId $environmentId -Guid $agentGuid -ArtifactName ([string]$agent.Name) -Owner $ownerValue -OwnerEmail $ownerEmail -OwnerObjectId $ownerObjectId -DateCreated $created -ConnectionReferences $connectionReferences -ConnectionReferenceCount $connectionReferenceCount
                $agentCount++
            }

            if (@($pacAgents).Count -gt 0 -and -not $dvEnrichmentAvailable -and -not $pacFallbackOwnerDateMissingWarningShown) {
                Write-Warning "PAC fallback owner/date enrichment via Dataverse was not available (Az.Accounts module with active login may be required). Agent Owner/OwnerEmail/DateCreated may be blank."
                $pacFallbackOwnerDateMissingWarningShown = $true
            }
        }
        catch {
            Write-Warning "Failed to collect agents through PAC for environment '$environmentDisplayName'. $($_.Exception.Message)"
        }
    }
}

# Handle agent cmdlets that only support tenant-level enumeration (no environment filter).
if ($null -ne $agentCommand -and -not $agentSupportsEnvironment) {
    try {
        $agents = & $agentCommandName

        foreach ($agent in $agents) {
            $environmentDisplayName = [string](Get-ValueByPath -InputObject $agent -Paths @("EnvironmentDisplayName", "EnvironmentName", "EnvironmentId"))
            $environmentId = [string](Get-ValueByPath -InputObject $agent -Paths @("EnvironmentName", "EnvironmentId"))

            if ([string]::IsNullOrWhiteSpace($environmentDisplayName)) {
                $environmentDisplayName = $environmentId
            }

            $ownerEmail = [string](Resolve-OwnerEmail -Item $agent)
            $ownerDisplay = [string](Resolve-OwnerDisplayName -Item $agent)
            $ownerGuid = Get-AgentOwnerGuidFromItem -Item $agent

            if ([string]::IsNullOrWhiteSpace($ownerEmail) -and -not [string]::IsNullOrWhiteSpace($ownerGuid)) {
                $ownerEmail = [string](Resolve-EntraEmailFromObjectId -ObjectId $ownerGuid)
            }

            $ownerValue = Resolve-OwnerField -OwnerDisplayName $ownerDisplay -OwnerEmail $ownerEmail
            if ([string]::IsNullOrWhiteSpace($ownerValue) -and -not [string]::IsNullOrWhiteSpace($ownerGuid)) {
                $ownerValue = $ownerGuid
            }

            $guid = [string](Get-ValueByPath -InputObject $agent -Paths @("BotId", "CopilotId", "AgentId", "Name", "Id"))
            $artifactName = [string](Get-ValueByPath -InputObject $agent -Paths @("DisplayName", "Name", "properties.displayName", "BotName"))
            $created = Convert-ToIsoDate (Get-ValueByPath -InputObject $agent -Paths @("CreatedTime", "CreatedDateTime", "properties.createdTime", "properties.createdDateTime", "CreatedOn"))
            $connectionReferences = $null
            $connectionReferenceCount = 0

            if ($IncludeConnectionReferences) {
                $agentRefSummary = Get-ConnectionReferencesFromItem -Item $agent
                if ($agentRefSummary.Count -gt 0) {
                    $connectionReferences = [string]$agentRefSummary.List
                    $connectionReferenceCount = [int]$agentRefSummary.Count
                }
                else {
                    $agentEnvUrl = $null
                    if (-not [string]::IsNullOrWhiteSpace($environmentId) -and $environmentUrlById.ContainsKey($environmentId)) {
                        $agentEnvUrl = [string]$environmentUrlById[$environmentId]
                    }

                    $agentToolSummary = Get-AgentToolSummary -EnvironmentUrl $agentEnvUrl -AgentId $guid
                    $connectionReferences = [string]$agentToolSummary.List
                    $connectionReferenceCount = [int]$agentToolSummary.Count
                }
            }

            Add-InventoryRecord -Buffer $inventory -ArtifactType "Agent" -Environment $environmentDisplayName -EnvironmentId $environmentId -Guid $guid -ArtifactName $artifactName -Owner $ownerValue -OwnerEmail $ownerEmail -OwnerObjectId $ownerGuid -DateCreated $created -ConnectionReferences $connectionReferences -ConnectionReferenceCount $connectionReferenceCount
            $agentCount++
        }
    }
    catch {
        Write-Warning "Failed to collect agents with command '$agentCommandName'. $($_.Exception.Message)"
    }
}

# Ensure output directory exists, then export a stable, sorted CSV.
$outputDirectory = Split-Path -Path $OutputCsvPath -Parent
if (-not [string]::IsNullOrWhiteSpace($outputDirectory) -and -not (Test-Path -Path $outputDirectory)) {
    New-Item -Path $outputDirectory -ItemType Directory -Force | Out-Null
}

$inventory |
Sort-Object -Property ArtifactType, Environment, GUID |
Export-Csv -Path $OutputCsvPath -NoTypeInformation -Encoding UTF8

Write-Host "Inventory complete." -ForegroundColor Green
Write-Host "Flows:  $flowCount" -ForegroundColor Green
Write-Host "Apps:   $appCount" -ForegroundColor Green
Write-Host "Agents: $agentCount" -ForegroundColor Green
Write-Host "Total:  $($inventory.Count)" -ForegroundColor Green
Write-Host "CSV:    $OutputCsvPath" -ForegroundColor Green