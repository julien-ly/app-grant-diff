#Requires -Version 7.0
#Requires -Modules Microsoft.Graph.Applications

<#
.SYNOPSIS
    Compares the application permissions declared on an app registration against
    the app role assignments actually granted to its service principal, and
    reports the gap.

.DESCRIPTION
    This tool does not administer, modify, recommend or decide anything.
    It establishes what a tenant declares, what it grants, and whether the two
    agree. A gap is a state to examine, not a fault.

    Three independent dimensions, reported separately.

    Conclusion:
      GapEstablished          At least one permission is under-covered or
                              over-covered.
      NoGapEstablished        Every principal is resolved, the manifest is
                              complete, and every permission is in its expected
                              state.
      CoverageNotDemonstrable No gap is proven and the assessment could not
                              cover the whole population.

    Assessment completeness, a property of the manifest:
      Complete   Manifest declared complete and fully resolved.
      Partial    Manifest incomplete, entries unresolved, or grant holders that
                 the reference does not allow judging. Reasons are listed and
                 name the principals concerned.
      Absent     No manifest provided.

    Observation freshness, a property of the grant snapshot:
      NotDemonstrated  Microsoft Graph documents replication delays on app role
                       assignments and exposes no convergence indicator.

    Freshness is never asserted as good. The tool has no way to establish that
    the snapshot it read has converged, so it says so rather than implying it.

    ENUMERATION AXIS: service principals, not app registrations.
    Starting from registrations can never reach the principals that hold grants
    without a local registration, and that is the population this tool exists to
    make visible. A paginated $expand establishes which principals hold grants;
    individual reads are then issued only where a zero carries a conclusion.

    $EXPAND ESTABLISHES PRESENCE, NEVER COMPLETENESS.
    For Entra resources deriving from directoryObject, $expand returns at most
    20 items of the expanded relationship and no @odata.nextLink. Navigation
    properties never carry an odata.count annotation either. A principal holding
    more than 20 assignments is therefore silently under-reported, with nothing
    in the payload to signal it. Every non-empty expanded collection is re-read
    individually before use. See README, "The $expand limit".

    ABSENCE OF A PROPERTY IS NOT ZERO.
    $expand omits appRoleAssignments entirely for principals that hold none. A
    principal whose grant state was not observed receives no evaluation row: it
    degrades assessment completeness and raises a diagnostic.

    Application permissions only. Delegated permissions are out of scope:
    dynamic consent, consent type and multiple scopes in a single value make the
    gap a normal state there and would produce false positives.

.PARAMETER IntentPath
    Optional. Path to an intent manifest. Without it, grants are reported as
    Observed and the conclusion cannot exceed CoverageNotDemonstrable, because
    requiredResourceAccess carries no exhaustiveness flag and over-coverage
    cannot be established from tenant data alone. See README, "Intent manifest".

.PARAMETER OutputPath
    Optional. Path of the JSON report. Defaults to report.json in the current
    directory.

.PARAMETER PassThru
    Emit the report object on the pipeline in addition to writing the file.

.PARAMETER Quiet
    Suppress the console summary. The file and -PassThru are unaffected.

.EXAMPLE
    Connect-MgGraph -Scopes 'Application.Read.All'
    .\Get-AppGrantDiff.ps1

.EXAMPLE
    .\Get-AppGrantDiff.ps1 -IntentPath .\samples\intent.json -OutputPath .\report.json

.NOTES
    Read-only. Requires Application.Read.All, the least privileged permission documented for every endpoint it reads.
    MIT. https://github.com/julien-ly/app-grant-diff
#>

[CmdletBinding()]
param(
    [string] $IntentPath,
    [string] $OutputPath,
    [switch] $PassThru,
    [switch] $Quiet
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$ToolVersion  = '1.0.0'
$NullRoleId   = '00000000-0000-0000-0000-000000000000'

if (-not $OutputPath) { $OutputPath = Join-Path (Get-Location) 'report.json' }

function Write-Line { param([string]$Text) if (-not $Quiet) { Write-Host $Text } }

# ── Session ─────────────────────────────────────────────────────────────────

# Mixed SDK versions fail with an assembly load error that names the assembly
# and not the cause. Say the cause here rather than leave it to the README.
$authLoaded = @(Get-Module Microsoft.Graph.Authentication)
if ($authLoaded.Count -gt 1) {
    throw "Several versions of Microsoft.Graph.Authentication are loaded ($(($authLoaded.Version) -join ', ')). An assembly cannot be unloaded from a running process: start a new one, and install the SDK submodules at a single matching version."
}
if ($authLoaded.Count -eq 1) {
    $appsLoaded = @(Get-Module Microsoft.Graph.Applications)
    if ($appsLoaded.Count -eq 1 -and $appsLoaded[0].Version -ne $authLoaded[0].Version) {
        Write-Warning "Microsoft.Graph.Authentication $($authLoaded[0].Version) and Microsoft.Graph.Applications $($appsLoaded[0].Version) are loaded at different versions. If a call fails on an assembly conflict, this is the cause."
    }
}

$context = Get-MgContext
if (-not $context) {
    throw "Not connected. Run: Connect-MgGraph -Scopes 'Application.Read.All'"
}
$tenantId = $context.TenantId
Write-Line "Tenant : $tenantId"

# ── Helpers ─────────────────────────────────────────────────────────────────

function Get-Field {
    # Invoke-MgGraphRequest may return hashtables or PSObjects depending on the
    # SDK version. Dot access works on both, Select-Object only on the second.
    # Read explicitly rather than assume the shape.
    param($Item, [string]$Name)
    if ($null -eq $Item) { return $null }
    if ($Item -is [System.Collections.IDictionary]) {
        if ($Item.Contains($Name)) { return $Item[$Name] }
        return $null
    }
    $p = $Item.PSObject.Properties[$Name]
    if ($p) { return $p.Value }
    return $null
}

function Get-EntrySignature {
    # Two entries for the same principal are identical only if they say the same
    # thing. Comparing on appId alone would call a contradiction a duplicate.
    param($Entry)
    $perms = @($Entry.Permissions | ForEach-Object { "$($_.Key)=$($_.ExpectedGranted)" } | Sort-Object) -join ';'
    "$($Entry.Complete)|$($Entry.PermissionsProvided)|$perms"
}

$diagnostics = [System.Collections.Generic.List[object]]::new()
function Add-Diagnostic {
    param([string]$Severity, [string]$Code, [string]$Object, [string]$Message)
    $diagnostics.Add([pscustomobject]@{
        severity = $Severity; code = $Code; object = $Object; message = $Message
    })
}

# ── 1. Intent manifest ──────────────────────────────────────────────────────
#    Contract validation. An evidence tool does not infer its own reference.

$intentStatus      = 'Absent'
$intentComplete    = $false
$intentDescription = $null
$intentAsOf        = $null
$intentHash        = $null
$intentDeclared    = 0
$intentNotResolved = 0
$intentDuplicates  = 0
$intentByAppId     = @{}

if ($IntentPath) {
    if (-not (Test-Path -LiteralPath $IntentPath)) {
        # Echo back the resolved path and the working directory. A relative path
        # repeated verbatim tells the reader nothing they did not already type.
        $resolved = [System.IO.Path]::GetFullPath($IntentPath, (Get-Location).Path)
        throw "Intent manifest not found. Given: $IntentPath. Resolved to: $resolved. Working directory: $((Get-Location).Path)."
    }

    $intentHash = (Get-FileHash -LiteralPath $IntentPath -Algorithm SHA256).Hash
    $raw = Get-Content -LiteralPath $IntentPath -Raw -Encoding utf8 | ConvertFrom-Json

    if ($null -eq $raw.PSObject.Properties['complete']) {
        throw "The manifest carries no 'complete' field. It is mandatory and never inferred."
    }
    # PowerShell casts any non-empty string to $true, so [bool]"false" is $true.
    # A JSON string where a boolean belongs would silently invert an
    # exhaustiveness claim. Check the type; never coerce.
    if ($raw.complete -isnot [bool]) {
        throw "The manifest declares 'complete' as $($raw.complete.GetType().Name), not a JSON boolean. A string is not coerced: PowerShell reads ""false"" as true."
    }
    $intentComplete    = $raw.complete
    $intentStatus      = 'Available'
    $intentDescription = if ($raw.PSObject.Properties['description']) { $raw.description } else { $null }
    $intentAsOf        = if ($raw.PSObject.Properties['asOf']) { $raw.asOf } else { $null }

    foreach ($p in @($raw.principals)) {
        $intentDeclared++
        $label = if ($p.PSObject.Properties['displayName']) { $p.displayName } else { $p.appId }

        if ($null -eq $p.PSObject.Properties['complete']) {
            throw "Entry '$label' carries no 'complete' field. It is mandatory per principal."
        }
        if ($p.complete -isnot [bool]) {
            $intentNotResolved++
            Add-Diagnostic 'Warning' 'IntentEntryMalformed' $label `
                "'complete' must be a JSON boolean, not $($p.complete.GetType().Name). Entry skipped: a string is not coerced."
            continue
        }
        if ($null -eq $p.PSObject.Properties['appId'] -or -not $p.appId) {
            $intentNotResolved++
            Add-Diagnostic 'Warning' 'IntentEntryNotResolved' $label 'The entry carries no usable appId.'
            continue
        }

        $permissions = @()
        $malformed   = $null
        if ($p.PSObject.Properties['permissions']) {
            foreach ($q in @($p.permissions)) {
                if ($null -eq $q.PSObject.Properties['expectedGranted'] -or $q.expectedGranted -isnot [bool]) {
                    $malformed = "$($q.resourceAppId)/$($q.value)"
                    break
                }
                $permissions += [pscustomobject]@{
                    ResourceAppId   = $q.resourceAppId
                    Value           = $q.value
                    ExpectedGranted = $q.expectedGranted
                    Key             = "$($q.resourceAppId)/$($q.value)"
                }
            }
        }
        if ($malformed) {
            # A manifest malformed about a principal cannot serve as its
            # reference. Skipping the whole entry is the conservative reading.
            $intentNotResolved++
            Add-Diagnostic 'Warning' 'IntentEntryMalformed' $label `
                "Permission '$malformed' must declare 'expectedGranted' as a JSON boolean. Entry skipped."
            continue
        }

        $candidate = [pscustomobject]@{
            DisplayName         = $label
            Complete            = $p.complete
            Permissions         = @($permissions)
            PermissionsProvided = [bool]$p.PSObject.Properties['permissions']
        }

        if ($intentByAppId.ContainsKey($p.appId)) {
            # A manifest that contradicts itself about a principal cannot serve
            # as a reference for it. Keeping the first silently would let an
            # invalid manifest produce a confident claim, decided by line order.
            if ((Get-EntrySignature $intentByAppId[$p.appId]) -ne (Get-EntrySignature $candidate)) {
                throw "Intent manifest declares contradictory expectations for '$label' ($($p.appId))."
            }
            $intentDuplicates++
            Add-Diagnostic 'Info' 'IntentEntryDuplicate' $label 'The principal appears more than once with identical content.'
            continue
        }

        $intentByAppId[$p.appId] = $candidate
    }
    # Only file-level resolution is known at this point. Resolution against the
    # tenant happens later; reporting it here would print a figure that a later
    # line contradicts.
    Write-Line "Manifest : $intentDeclared entries loaded, $($intentByAppId.Count) usable"
} else {
    Write-Line 'Manifest : none. The conclusion cannot exceed CoverageNotDemonstrable.'
}

# ── 2. Granted side: paginated $expand ──────────────────────────────────────

Write-Line ''
Write-Line 'Reading assignments'

$grantsBySpId  = @{}   # spObjectId -> assignments, key absent when not observed
$spBySpId      = @{}
$spByAppId     = @{}
$nestedSignals = @{}
$expandCount   = @{}

$uri = 'https://graph.microsoft.com/v1.0/servicePrincipals?$select=id,appId,displayName,appOwnerOrganizationId&$expand=appRoleAssignments'
$sw = [System.Diagnostics.Stopwatch]::StartNew()
$pages = 0
$next = $uri
while ($next) {
    $page = Invoke-MgGraphRequest -Method GET -Uri $next -OutputType PSObject
    $pages++
    foreach ($v in @(Get-Field -Item $page -Name 'value')) {
        $id    = Get-Field -Item $v -Name 'id'
        $appId = Get-Field -Item $v -Name 'appId'
        $sp = [pscustomobject]@{
            Id          = $id
            AppId       = $appId
            DisplayName = Get-Field -Item $v -Name 'displayName'
            OwnerOrgId  = Get-Field -Item $v -Name 'appOwnerOrganizationId'
        }
        $spBySpId[$id] = $sp
        if ($appId) { $spByAppId[$appId] = $sp }

        $assignments = Get-Field -Item $v -Name 'appRoleAssignments'
        if ($null -ne $assignments) {
            $grantsBySpId[$id] = @(@($assignments) | ForEach-Object {
                [pscustomobject]@{
                    Id         = Get-Field -Item $_ -Name 'id'
                    AppRoleId  = Get-Field -Item $_ -Name 'appRoleId'
                    ResourceId = Get-Field -Item $_ -Name 'resourceId'
                }
            })
            # Does Graph announce the rest of the nested collection? If it does,
            # the truncation is signalled and the reader is at fault. If it does
            # not, it is silent. The difference is observed, not assumed.
            $nestedLink  = Get-Field -Item $v -Name 'appRoleAssignments@odata.nextLink'
            $nestedCount = Get-Field -Item $v -Name 'appRoleAssignments@odata.count'
            if ($nestedLink -or $nestedCount) {
                $nestedSignals[$id] = [pscustomobject]@{ nextLink = $nestedLink; count = $nestedCount }
            }
        }
    }
    $next = Get-Field -Item $page -Name '@odata.nextLink'
}
$sw.Stop()
Write-Line "$($spBySpId.Count) principals, $pages pages, $([math]::Round($sw.Elapsed.TotalSeconds,1)) s"
Write-Line "$($grantsBySpId.Count) principals with the relationship rendered"

# ── 3. Declared side: local app registrations ───────────────────────────────

$declaredByAppId = @{}
foreach ($app in @(Get-MgApplication -All -Property 'id,appId,displayName,requiredResourceAccess')) {
    $declared = @()
    foreach ($resource in @($app.RequiredResourceAccess)) {
        foreach ($access in @($resource.ResourceAccess)) {
            if ($access.Type -eq 'Role') {
                $declared += [pscustomobject]@{ ResourceAppId = $resource.ResourceAppId; RoleId = $access.Id }
            }
        }
    }
    if ($declaredByAppId.ContainsKey($app.AppId)) {
        Add-Diagnostic 'Warning' 'DuplicateApplication' $app.DisplayName 'Several local registrations carry the same appId.'
        continue
    }
    $declaredByAppId[$app.AppId] = [pscustomobject]@{
        ObjectId = $app.Id; DisplayName = $app.DisplayName; Declared = @($declared)
    }
}
Write-Line "$($declaredByAppId.Count) local registrations"

# ── 4. Role name resolution ─────────────────────────────────────────────────

$rolesByAppId = @{}
function Get-RoleValue {
    param([string]$ResourceAppId, [string]$RoleId)
    if ($RoleId -eq $NullRoleId) { return '(assigned without a specific role)' }
    if (-not $rolesByAppId.ContainsKey($ResourceAppId)) {
        $found = @(Get-MgServicePrincipal -Filter "appId eq '$ResourceAppId'" -Property 'id,appId,displayName,appRoles')
        $rolesByAppId[$ResourceAppId] = if ($found.Count -eq 1) { $found[0] } else { $null }
    }
    $sp = $rolesByAppId[$ResourceAppId]
    if (-not $sp) { return "(role $RoleId on resource $ResourceAppId absent from the tenant)" }
    $match = @($sp.AppRoles | Where-Object { $_.Id -eq $RoleId })
    if ($match.Count -eq 0) { return "(role $RoleId not exposed by $($sp.DisplayName))" }
    $match[0].Value
}

function Get-ResourceAppId {
    param([string]$ResourceSpId)
    if (-not $spBySpId.ContainsKey($ResourceSpId)) { return "(sp $ResourceSpId unknown)" }
    $spBySpId[$ResourceSpId].AppId
}

# ── 5. Scope, and targeted individual reads ─────────────────────────────────

$scope = [System.Collections.Generic.HashSet[string]]::new()
foreach ($id in $grantsBySpId.Keys) { if (@($grantsBySpId[$id]).Count -gt 0) { [void]$scope.Add($id) } }
foreach ($appId in $declaredByAppId.Keys) {
    if ($declaredByAppId[$appId].Declared.Count -gt 0 -and $spByAppId.ContainsKey($appId)) {
        [void]$scope.Add($spByAppId[$appId].Id)
    }
}
foreach ($appId in $intentByAppId.Keys) {
    if ($spByAppId.ContainsKey($appId)) { [void]$scope.Add($spByAppId[$appId].Id) }
    else {
        # An entry that names a principal absent from the tenant is unresolved,
        # exactly like one that carries no usable appId. It must degrade
        # over-coverage demonstrability: a manifest claiming exhaustiveness
        # while pointing at something that does not exist is not fully resolved.
        $intentNotResolved++
        Add-Diagnostic 'Warning' 'IntentPrincipalAbsent' $intentByAppId[$appId].DisplayName 'The manifest principal does not exist in the tenant. Counted as unresolved.'
    }
}

$zeroConfirmingReads = 0
$completenessReads   = 0
$truncated           = 0
$notObserved = [System.Collections.Generic.List[string]]::new()
# Principals whose grants are PRESENT but whose completeness could not be
# verified, because the individual re-read failed. Their observed grants are
# facts; their absences are not.
$grantsUnverified = [System.Collections.Generic.HashSet[string]]::new()

foreach ($id in @($scope)) {

    $alreadyRendered = $grantsBySpId.ContainsKey($id)
    if ($alreadyRendered) {
        $expandCount[$id] = @($grantsBySpId[$id]).Count
        if ($expandCount[$id] -eq 0) { continue }
        $completenessReads++
    } else {
        $zeroConfirmingReads++
    }

    try {
        $full = @(@(Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $id -All) | ForEach-Object {
            [pscustomobject]@{ Id = $_.Id; AppRoleId = $_.AppRoleId; ResourceId = $_.ResourceId }
        })

        if ($alreadyRendered -and $full.Count -ne $expandCount[$id]) {
            $truncated++
            $signalled = $nestedSignals.ContainsKey($id)
            # Both sets are in hand at the moment the gap is detected. Reporting
            # only how many were missing leaves the reader with an unidentified
            # absence: five roles could be read-only or Directory.ReadWrite.All.
            # Establishing a gap without naming its scope is the same failure the
            # tool reports elsewhere, applied to its own diagnostic.
            $renderedIds = @($grantsBySpId[$id] | ForEach-Object { $_.Id })
            $missing = @($full | Where-Object { $_.Id -notin $renderedIds } | ForEach-Object {
                Get-RoleValue (Get-ResourceAppId $_.ResourceId) $_.AppRoleId
            })
            Add-Diagnostic 'Warning' 'ExpandCollectionTruncated' $spBySpId[$id].DisplayName `
                ("`$expand returned $($expandCount[$id]) assignments out of $($full.Count) actual. " +
                 "Missing: $($missing -join ', '). " +
                 $(if ($signalled) { "Graph announced the rest: $($nestedSignals[$id] | ConvertTo-Json -Compress)." }
                   else { 'No continuation link and no announced count in the payload. Documented behaviour: $expand returns at most 20 items for an expanded relationship on a directoryObject-derived resource, with no @odata.nextLink. See learn.microsoft.com/graph/query-parameters#expand.' }))
        }

        $grantsBySpId[$id] = $full
    } catch {
        if ($alreadyRendered) {
            # The expanded collection stands as evidence of presence. It cannot
            # stand as evidence of absence: $expand caps expanded relationships,
            # so a permission missing from it may exist and not have been
            # returned. Keeping the grants and evaluating absences against them
            # would turn an unverified state into an observed one.
            [void]$grantsUnverified.Add($id)
            Add-Diagnostic 'Warning' 'CompletenessNotVerified' $spBySpId[$id].DisplayName `
                "Re-read failed, the `$expand result could not be verified, so its completeness is unknown: $($_.Exception.Message)"
        } else {
            $notObserved.Add($id)
            Add-Diagnostic 'Error' 'GrantStateNotObserved' $spBySpId[$id].DisplayName `
                "Assignments could not be read: $($_.Exception.Message)"
        }
    }
}
Write-Line "$($scope.Count) principals in scope"
Write-Line "$zeroConfirmingReads reads to confirm a zero, $completenessReads to verify completeness"
if ($truncated -gt 0 -and -not $Quiet) { Write-Warning "$truncated `$expand collection(s) truncated." }

# ── 6. Evaluation matrix ────────────────────────────────────────────────────

$evaluation          = [System.Collections.Generic.List[object]]::new()
$completenessReasons = [System.Collections.Generic.List[string]]::new()
$notJudgeable        = [System.Collections.Generic.List[string]]::new()
$withoutRows         = [System.Collections.Generic.List[object]]::new()
$correctlyEmptyCount = 0

if ($intentStatus -eq 'Absent')                              { $completenessReasons.Add('No intent manifest was provided.') }
if ($intentStatus -eq 'Available' -and -not $intentComplete) { $completenessReasons.Add('The manifest declares itself not complete.') }
if ($intentNotResolved -gt 0)                                { $completenessReasons.Add("$intentNotResolved manifest entries could not be resolved.") }

foreach ($spId in @($scope)) {
    $sp = $spBySpId[$spId]
    $appId = $sp.AppId

    if ($notObserved.Contains($spId)) {
        $completenessReasons.Add("The grant state of '$($sp.DisplayName)' was not observed.")
        continue
    }

    $entry        = if ($appId -and $intentByAppId.ContainsKey($appId))   { $intentByAppId[$appId] }   else { $null }
    $registration = if ($appId -and $declaredByAppId.ContainsKey($appId)) { $declaredByAppId[$appId] } else { $null }

    # requiredResourceAccess is a POSITIVE-ONLY declaration. It has no way to
    # express "explicitly not expected": the absence of an entry there means
    # "not declared", not "excluded". A negative expectation can therefore only
    # come from the manifest, including for a principal that has a registration.
    # Without this rule CorrectExclusion is unreachable as soon as a
    # registration exists, and the state is never produced.

    $expected         = @{}
    $declaredPositive = @{}

    if ($registration) {
        foreach ($d in $registration.Declared) {
            $value = Get-RoleValue $d.ResourceAppId $d.RoleId
            $declaredPositive["$($d.ResourceAppId)/$value"] = $true
            $expected["$($d.ResourceAppId)/$value"] = $true
        }
    }

    if ($entry -and $entry.PermissionsProvided) {
        foreach ($p in $entry.Permissions) {
            if (-not $registration) {
                $expected[$p.Key] = $p.ExpectedGranted
            }
            elseif (-not $p.ExpectedGranted) {
                if ($declaredPositive.ContainsKey($p.Key)) {
                    Add-Diagnostic 'Warning' 'ManifestDeclarationMismatch' $sp.DisplayName `
                        "The manifest declares '$($p.Key)' explicitly not expected while the registration declares it. The registration prevails; the contradiction is reported."
                } else {
                    $expected[$p.Key] = $false
                }
            }
        }

        if ($registration) {
            $manifestPositive = @{}
            foreach ($p in $entry.Permissions) { if ($p.ExpectedGranted) { $manifestPositive[$p.Key] = $true } }
            if ($manifestPositive.Count -gt 0) {
                $registrationOnly = @($declaredPositive.Keys | Where-Object { -not $manifestPositive.ContainsKey($_) })
                $manifestOnly     = @($manifestPositive.Keys | Where-Object { -not $declaredPositive.ContainsKey($_) })
                if ($registrationOnly.Count -gt 0 -or $manifestOnly.Count -gt 0) {
                    Add-Diagnostic 'Warning' 'ManifestDeclarationMismatch' $sp.DisplayName `
                        ("Manifest and requiredResourceAccess diverge on positive expectations. Registration only: $($registrationOnly -join ', '). Manifest only: $($manifestOnly -join ', ').")
                }
            }
        }
    }

    $observed = @{}
    foreach ($g in @($grantsBySpId[$spId])) {
        $resourceAppId = Get-ResourceAppId $g.ResourceId
        $value = Get-RoleValue $resourceAppId $g.AppRoleId
        $observed["$resourceAppId/$value"] = $true
    }

    # Over-coverage is demonstrable only when an exhaustive reference exists for
    # THIS principal. Without a manifest, without an entry, or with an entry that
    # does not claim to be exhaustive, an unmatched grant stays NotInManifest, or
    # Observed when there is no manifest at all.
    $overDemonstrable = ($intentStatus -eq 'Available') -and $entry -and $entry.Complete -and ($intentNotResolved -eq 0)

    $producedNotInManifest = $false

    # When completeness is unverified, only rows carrying an observed grant are
    # emitted. Every state that rests on an absence - UnderCoverage and
    # CorrectExclusion - is withheld and moved to assessment.reasons, which is
    # where things the engine could not establish belong.
    $completenessUnverified = $grantsUnverified.Contains($spId)
    $withheld = [System.Collections.Generic.List[string]]::new()
    $rowsBefore = $evaluation.Count

    foreach ($key in @(@($expected.Keys) + @($observed.Keys) | Select-Object -Unique)) {
        $exp = if ($expected.ContainsKey($key)) { [bool]$expected[$key] } else { $null }
        $obs = $observed.ContainsKey($key)

        if ($completenessUnverified -and -not $obs) { $withheld.Add($key); continue }

        $state =
            if     ($exp -eq $true  -and $obs)      { 'CorrectCoverage' }
            elseif ($exp -eq $true  -and -not $obs) { 'UnderCoverage' }
            elseif ($exp -eq $false -and -not $obs) { 'CorrectExclusion' }
            elseif ($exp -eq $false -and $obs)      { if ($overDemonstrable) { 'OverCoverage' } else { 'NotInManifest' } }
            elseif ($intentStatus -eq 'Absent')     { 'Observed' }
            elseif ($overDemonstrable)              { 'OverCoverage' }
            else                                    { 'NotInManifest' }

        if ($state -eq 'NotInManifest') { $producedNotInManifest = $true }

        $parts = $key -split '/', 2
        $evaluation.Add([pscustomobject]@{
            principalId       = $spId
            principalAppId    = $appId
            displayName       = $sp.DisplayName
            resourceAppId     = $parts[0]
            permission        = $parts[1]
            expectedGranted   = $exp
            observedGranted   = $obs
            state             = $state
            source            = if ($null -eq $exp) { 'grant' } elseif ($registration) { 'registration' } else { 'intent' }
            localRegistration = [bool]$registration
        })
    }

    # A principal in scope that produced no row at all is invisible in the
    # evaluation, and a reader cannot tell it from one that was never reached.
    # Report the facts the engine holds about it - no taxonomy. Whether an object
    # is a resource rather than a client is not among them: the engine reads
    # neither appRoles nor appRoleAssignedTo, so saying so would be an inference
    # drawn from its name.
    if ($evaluation.Count -eq $rowsBefore) {
        $expectedHasSource = $entry -and ($entry.PermissionsProvided -or $registration)
        $isCorrectlyEmpty  = ($intentStatus -eq 'Available') -and $entry -and $entry.Complete `
                             -and $expectedHasSource -and ($expected.Count -eq 0) -and ($observed.Count -eq 0)
        if ($isCorrectlyEmpty) { $correctlyEmptyCount++ }
        $withoutRows.Add([pscustomobject]@{
            displayName         = $sp.DisplayName
            principalAppId      = $appId
            localRegistration   = [bool]$registration
            declaredPermissions = $declaredPositive.Count
            observedGrants      = $observed.Count
            inManifest          = [bool]$entry
            manifestComplete    = if ($entry) { $entry.Complete } else { $null }
            correctlyEmpty      = $isCorrectlyEmpty
        })
    }

    if ($completenessUnverified) {
        $detail = if ($withheld.Count -gt 0) { " Withheld: $($withheld -join ', ')." } else { '' }
        $completenessReasons.Add("The completeness of the grants of '$($sp.DisplayName)' could not be verified, so no absence was established for it.$detail")
    }

    # Name a principal only when it actually produced an unjudgeable row. A
    # degraded reference does not make every grant holder unjudged: a principal
    # whose every assignment matches a declaration is judged CorrectCoverage,
    # and listing it here would overstate what the sentence describes.
    if ($intentStatus -eq 'Available' -and $producedNotInManifest) {
        $notJudgeable.Add($sp.DisplayName)
        if ($intentComplete -and -not $entry) {
            Add-Diagnostic 'Warning' 'ManifestClaimsCompleteButOmitsGrantHolder' $sp.DisplayName `
                'The manifest declares itself complete and does not contain this principal, which holds assignments. The exhaustiveness claim is contradicted by the tenant.'
        }
    }
}

# ── 7. Three axes ───────────────────────────────────────────────────────────

# summary has two tiers because its counters do not count the same thing, and do
# not cover the same population. perPermission counts evaluation rows, one per
# permission, across every principal read. perEntry counts manifest ENTRIES, and
# only exists when a manifest was supplied. Flattening the two would let a reader,
# or an aggregator, sum eight numbers into a total that designates nothing.
#
# Neither perEntry counter is a row state: both count entries that produced no row
# at all. They are carried anyway, because summary is where a reader counts what
# happened. NotResolved: the entry named something that could not be resolved, so
# nothing was evaluated. CorrectlyEmpty: the entry declared itself complete, what
# it expected was empty, and nothing was granted - an affirmation the engine
# verified. Without them the totals read as full coverage at one end and as a
# shorter population at the other.

$states = @('CorrectCoverage','CorrectExclusion','UnderCoverage','OverCoverage','NotInManifest','Observed')
$perPermission = [ordered]@{}
foreach ($s in $states) { $perPermission[$s] = @($evaluation | Where-Object { $_.state -eq $s }).Count }

# Without a manifest, perEntry is meaningless rather than zero: there are no
# entries to count. Two zeros would read as two passed verifications.
$perEntry = if ($intentStatus -eq 'Available') {
    [pscustomobject][ordered]@{ NotResolved = $intentNotResolved; CorrectlyEmpty = $correctlyEmptyCount }
} else { $null }

if ($notJudgeable.Count -gt 0) {
    $names = @($notJudgeable | Sort-Object -Unique)
    $excerpt = if ($names.Count -le 5) { $names -join ', ' } else { (@($names)[0..4] -join ', ') + ", and $($names.Count - 5) more" }
    $completenessReasons.Add("$($names.Count) principal(s) hold assignments the reference does not allow judging: $excerpt.")
}

$reasons = @($completenessReasons | Select-Object -Unique)
$completeness = if     ($intentStatus -eq 'Absent') { 'Absent' }
                elseif ($reasons.Count -gt 0)       { 'Partial' }
                else                                { 'Complete' }

$observation = [pscustomobject]@{
    replicationIndicator = 'NotExposed'
    expandCompleteness   = if ($truncated -gt 0) { 'TruncatedAndRepaired' } else { 'NotContradicted' }
    freshness            = 'NotDemonstrated'
    note                 = 'Microsoft Graph documents replication delays on app role assignments and exposes no convergence indicator. That the snapshot read is current has not been established.'
}

$hasGap = ($perPermission['UnderCoverage'] -gt 0) -or ($perPermission['OverCoverage'] -gt 0)

$conclusion =
    if ($hasGap) {
        $parts = @()
        if ($perPermission['UnderCoverage'] -gt 0) { $parts += "$($perPermission['UnderCoverage']) permission(s) declared and not granted" }
        if ($perPermission['OverCoverage']  -gt 0) { $parts += "$($perPermission['OverCoverage']) permission(s) granted and not expected" }
        [pscustomobject]@{
            result            = 'GapEstablished'
            detail            = ($parts -join '. ') + '.'
            gapBreakCondition = 'The manifest is out of date or wrong about what is expected for these principals, or the assignment snapshot was not current when the report was generated.'
            impactBoundary    = 'Delegated permissions, user consent, directory role membership and managed identities are not evaluated. A gap here does not establish what the principal actually accesses.'
        }
    }
    elseif ($completeness -ne 'Complete') {
        [pscustomobject]@{
            result            = 'CoverageNotDemonstrable'
            detail            = 'No gap was established, and the assessment could not cover the whole population. ' + ($reasons -join ' ')
            gapBreakCondition = $null
            impactBoundary    = $null
        }
    }
    else {
        [pscustomobject]@{
            result            = 'NoGapEstablished'
            detail            = 'No gap observed. The manifest was complete and fully resolved. The freshness of the assignments was not independently established.'
            gapBreakCondition = 'The manifest is out of date, or it claims exhaustiveness while omitting part of the real population, or the assignments had not finished converging when read.'
            impactBoundary    = $null
        }
    }

# ── 8. Report ───────────────────────────────────────────────────────────────

$report = [pscustomobject]@{
    metadata = [pscustomobject]@{
        toolVersion       = $ToolVersion
        generatedAt       = (Get-Date).ToUniversalTime().ToString('o')
        powerShellVersion = $PSVersionTable.PSVersion.ToString()
        tenantId          = $tenantId
        disclaimer        = 'Conclusions are bounded to application permissions. Delegated permissions are not evaluated.'
    }
    intentSource = [pscustomobject]@{
        status              = $intentStatus
        path                = $IntentPath
        description         = $intentDescription
        asOf                = $intentAsOf
        sha256              = $intentHash
        complete            = if ($intentStatus -eq 'Available') { $intentComplete } else { $null }
        declaredObjectCount = $intentDeclared
        resolvedObjectCount = $intentByAppId.Count
        notResolved         = $intentNotResolved
        duplicatesIgnored   = $intentDuplicates
    }
    read = [pscustomobject]@{
        principalsTotal          = $spBySpId.Count
        expandPages              = $pages
        expandDurationMs         = $sw.ElapsedMilliseconds
        relationRendered         = $expandCount.Count
        scopedPrincipals         = $scope.Count
        zeroConfirmingReads      = $zeroConfirmingReads
        completenessReads        = $completenessReads
        expandTruncated          = $truncated
        grantStateNotObserved    = $notObserved.Count
        grantCompletenessUnverified = $grantsUnverified.Count
        evaluatedWithoutRows        = $withoutRows.ToArray()
        grantHoldersNotJudgeable = @($notJudgeable | Sort-Object -Unique).Count
    }
    evaluation  = $evaluation.ToArray()
    summary     = [pscustomobject]@{
        perPermission = [pscustomobject]$perPermission
        perEntry      = $perEntry
    }
    assessment  = [pscustomobject]@{ completeness = $completeness; reasons = $reasons }
    observation = $observation
    conclusion  = $conclusion
    diagnostics = $diagnostics.ToArray()
}

$report | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $OutputPath -Encoding utf8

if (-not $Quiet) {
    Write-Host ''
    $evaluation | Format-Table displayName, permission, expectedGranted, observedGranted, state -AutoSize
    Write-Host "completeness : $completeness"
    foreach ($r in $reasons) { Write-Host "  - $r" }
    Write-Host "freshness    : $($observation.freshness)"
    Write-Host "conclusion   : $($conclusion.result)"
    Write-Host "               $($conclusion.detail)"
    if ($diagnostics.Count -gt 0) {
        Write-Host ''
        Write-Host 'Diagnostics:'
        $diagnostics | Format-Table severity, code, object, message -AutoSize -Wrap
    }
    Write-Host ''
    Write-Host "Report : $OutputPath"
}

if ($PassThru) { $report }
