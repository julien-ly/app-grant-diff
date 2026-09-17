# app-grant-diff

Compares the application permissions declared on an app registration against the app role assignments actually granted to its service principal, and reports the gap.

Read-only. It does not administer, modify, recommend or decide anything. A gap is a state to examine, not a fault: Microsoft asks developers to list permissions in `requiredResourceAccess` so that the admin consent flow works, but an `appRoleAssignment` can be created directly through Graph without going through the declaration.

## What it shows

A four-state matrix crossing the expected permission against the observed grant.

**Correct coverage.** The permission is expected and granted.
**Under-coverage.** The permission is expected and not granted. The application will fail on a function, and nothing signals it before the incident.
**Over-coverage.** The permission is granted and not expected. The principal retains an active access the registration no longer mentions, and a review that reads the registration cannot see it.
**Correct exclusion.** The permission is explicitly not expected and not granted.

Two further states exist for grants the matrix cannot classify. **Not in manifest**, when a grant holder is absent from a manifest that does not claim exhaustiveness, or carries an entry that does not. **Observed**, when no manifest was provided at all.

## Two tiers in summary

`summary` has two tiers, because its counters do not count the same thing and do not cover the same population.

**`summary.perPermission`** counts evaluation rows, one per permission, across every principal read: the six states above.

**`summary.perEntry`** counts manifest *entries*, and is null when no manifest was supplied. Neither of its counters is a row state; both count entries that produced no evaluation row at all.

**`NotResolved`** — the entry named a principal absent from the tenant, or carried no usable `appId`. Nothing was evaluated for it.
**`CorrectlyEmpty`** — the entry declared itself complete, what it expected was empty, and nothing was granted. An affirmation the engine verified, not a silence.

Both are carried because `summary` is where a reader counts what happened, not where rows are summed. Without `NotResolved` the totals read as full coverage while some entries were never evaluated. Without `CorrectlyEmpty` a principal verified as expecting nothing and holding nothing counts nowhere, and the totals read as if it had not been part of the evaluation.

The two tiers are separate objects rather than eight flat keys so that nothing sums by accident. A total across all of them would designate nothing.

`CorrectlyEmpty` is restricted to entries that make the affirmation: `complete: true`, an expected set that is empty, and a source for that expected set — either `permissions` supplied, or a local registration the entry vouches for. `complete: false` with an empty `permissions` affirms nothing at all: it lists expectations without claiming exhaustiveness, and lists none. It is indistinguishable from having no entry.

The restriction is what makes the counter mean something where the tool is otherwise blind. A third-party principal with no local registration and no assignments can produce no fact: no readable declaration, no grant, so no row. The manifest is the only possible source of declared state for it, and "nothing is expected" is the only statement that can be made about it.

An unresolved entry degrades the **whole report**, not only its own principal. Over-coverage is established only against a fully resolved reference, so one stale line in a manifest of a thousand suppresses `OverCoverage` across the tenant. `complete: false` on a principal, by contrast, degrades that principal alone. The rule is deliberate, since a manifest naming something that does not exist does not describe the population it claims to describe; the asymmetry between the two is the part that surprises.

Application permissions only. Delegated permissions are out of scope: dynamic consent, consent type and multiple scopes in a single value make the gap a normal state there and would produce false positives.

## Why the gap is not hypothetical

Microsoft documents that removing a permission from an app registration does not revoke permissions already granted, and that revocation must be done manually. The documented removal procedure requires two actions in two separate blades: remove the permission under App registrations, then have an administrator revoke it under Enterprise applications. Over-coverage is therefore the normal consequence of a documented procedure carried out halfway.

The portal reports the "granted but not configured" state one application at a time, never at tenant scale. Existing scripts and reports list what is consented; none compares it with what is declared.

## Three axes, reported separately

A proven gap, an incomplete manifest and a grant snapshot of unknown freshness are three different facts. The report keeps them apart.

**Conclusion.**
`GapEstablished` at least one permission is under-covered or over-covered.
`NoGapEstablished` the manifest is complete, every entry resolved, and every permission in its expected state.
`CoverageNotDemonstrable` no gap was proven, and the assessment could not cover the whole population.

**Assessment completeness**, a property of the manifest.
`Complete` manifest declared complete and fully resolved.
`Partial` manifest incomplete, entries unresolved, or grant holders the reference does not allow judging. Reasons are listed and name the principals concerned, so that one unjudged principal is not read as three hundred.
`Absent` no manifest provided.

**Observation freshness**, a property of the grant snapshot.
`NotDemonstrated` Microsoft Graph documents replication delays on app role assignments and exposes no convergence indicator.

Freshness is never asserted as good. The tool has no way to establish that the snapshot it read has converged, so it says so rather than implying it.

A gap that has been proven does not disappear because the rest of the population could not be assessed. `GapEstablished` with `Partial` completeness is a normal and useful result. Conversely `NoGapEstablished` reads as an absence of gap in the retrieved snapshot, not as a guarantee about a converged one.

All conclusions are bounded to application permissions. The report never states that a principal is over-privileged, only that a grant is not matched by a declaration. Delegated permissions, user consent, directory role membership and managed identities are not evaluated.

## Reading the report

Read `assessment.reasons` before the states.

An absence of `OverCoverage` means one of two things: no grant is unmatched, or the reference did not allow establishing that any grant is unmatched. The row states alone do not separate them. `assessment` does, and it names the principals concerned rather than only counting them, so that one unjudged principal is not read as three hundred.

Reading the states first and concluding "no gap" is the same failure this tool exists to detect, silence taken for absence, applied to its own output.

## What would make this redundant

A consent report would not be enough: it lists what is granted, never the difference with what is declared. Neither would a permission inventory: the same list, read from the other side.

The tool becomes redundant when the platform natively compares, at tenant scale, the application permissions declared on a registration against the app role assignments held by its service principal, and reports both directions of the difference along with the completeness of the assessment.

## Requirements

PowerShell 7 or later. Microsoft Graph PowerShell SDK module `Microsoft.Graph.Applications`, which pulls `Microsoft.Graph.Authentication`. Install them at a single matching version; mixed versions of the SDK fail to load, and the failure names an assembly rather than the cause.

```powershell
$v = (Find-Module Microsoft.Graph.Authentication -Repository PSGallery).Version
Install-Module Microsoft.Graph.Authentication -RequiredVersion $v -Scope CurrentUser
Install-Module Microsoft.Graph.Applications   -RequiredVersion $v -Scope CurrentUser
```

**Read-only.** `Application.Read.All`, `Directory.Read.All`. Nothing more; the tool never writes.

## Usage

```powershell
Connect-MgGraph -Scopes 'Application.Read.All'

# Without a manifest: reports grants as Observed and concludes at best
# CoverageNotDemonstrable
.\Get-AppGrantDiff.ps1

# With a manifest: produces the four-state matrix and a conclusion
.\Get-AppGrantDiff.ps1 -IntentPath .\samples\intent.json

# Write the report elsewhere, and keep the object
.\Get-AppGrantDiff.ps1 -IntentPath .\intent.json -OutputPath .\report.json -PassThru
```

## Intent manifest

The manifest does not replace `requiredResourceAccess`. The expected side remains the declaration carried by the registration: comparing it with the grants is the point of the tool. The manifest supplies the two things Entra does not carry.

**`complete`, per principal.** Nothing in Entra states whether a registration lists everything it is supposed to have. Without that flag, a grant without a declaration cannot be established as over-coverage: it stays `NotInManifest`. This is not a design preference. `requiredResourceAccess` has no exhaustiveness semantics, so over-coverage is not demonstrable from tenant data alone.

**`permissions`, for principals without a local registration.** There, Entra carries no declared state at all, and the manifest is the only place one can exist. This is what makes the hardest population to audit by hand evaluable: a service principal that holds grants and has no app registration in the tenant.

For a principal that does have a registration, `permissions` is optional. If present it does not override `requiredResourceAccess`: it is compared against it, and a disagreement raises a `ManifestDeclarationMismatch` diagnostic. A manifest that contradicts a registration is an audit signal, not a conflict to resolve silently.

One exception, and it is structural. `requiredResourceAccess` is a positive-only declaration: the absence of an entry means "not declared", not "excluded". A negative expectation can therefore only come from the manifest, and it is honoured even when a registration exists. Without that rule `CorrectExclusion` is unreachable as soon as a registration exists.

`complete` is mandatory at both levels, must be a JSON boolean, and is never inferred. A missing field raises an error; a wrong type is refused rather than coerced, because PowerShell reads the non-empty string `"false"` as `$true` and would silently invert the claim. The same check applies to `expectedGranted`.

A principal declared twice with different content raises an error rather than resolving by line order: a manifest that contradicts itself about a principal cannot serve as its reference. Two identical entries are inert and reported as `IntentEntryDuplicate`. There are two levels because a tenant holds many principals, not one group: `complete` on the manifest means the list of principals is exhaustive, `complete` on a principal means what is expected for that principal is complete.

Principals are keyed by `appId`, not `displayName`. An `appId` is stable across tenants and unambiguous; display names are not, and Microsoft first-party principals duplicate them.

See `samples/intent.json`.

## The $expand limit

The tool enumerates service principals rather than app registrations. Starting from registrations can never reach the principals that hold grants without a local registration, and that is the population this tool exists to make visible.

Reading assignments principal by principal costs one call each. A single paginated `$expand=appRoleAssignments` is roughly thirty times faster. It also has two properties that make it unusable on its own.

**Absence of the property is not zero.** `$expand` omits `appRoleAssignments` entirely for principals that hold none, rather than returning an empty array. A reader that treats the absence as zero converts an unobserved state into an observed one.

**$expand establishes presence, never completeness.** For Entra resources deriving from `directoryObject`, `$expand` returns at most 20 items of the expanded relationship and no `@odata.nextLink`. Navigation properties never carry an `odata.count` annotation either. A principal holding more than 20 assignments is silently under-reported, with nothing in the payload to signal it. Verified in a lab tenant: a principal holding 25 assignments returned 20, the five missing ones scattered through the series rather than at the end.

The documented behaviour is at [learn.microsoft.com/graph/query-parameters#expand](https://learn.microsoft.com/graph/query-parameters#expand). It is not hidden; it is published on a query-parameter page that few people read while building a tenant-wide permission inventory.

The tool therefore uses `$expand` to establish which principals hold grants, then re-reads individually every non-empty expanded collection, and issues a targeted read wherever a zero carries a conclusion. On a tenant of 305 principals that is four pages plus nine individual reads, against 305 in a naive enumeration. A repaired truncation is reported as `ExpandCollectionTruncated` and surfaced in `observation.expandCompleteness`.

`ExpandCollectionTruncated` an expanded collection was incomplete and was repaired by an individual read. The diagnostic names the permissions that `$expand` had omitted: how many were missing does not tell a reader whether they were read-only roles or `Directory.ReadWrite.All`.
`read.evaluatedWithoutRows` lists the principals in scope that produced no evaluation row, with the facts the engine holds about each: local registration, declared count, observed count, presence in the manifest. No taxonomy. Whether an object is a resource rather than a client is not among those facts — the engine reads neither `appRoles` nor `appRoleAssignedTo` — so saying so would be an inference drawn from its name.

## Diagnostics

`CompletenessNotVerified` the re-read failed, so the completeness of the expanded collection is unknown. Its grants still count as present; no absence is established for that principal, so `UnderCoverage` and `CorrectExclusion` are withheld and named in `assessment.reasons` instead.
`GrantStateNotObserved` assignments could not be read for a principal in scope.
`ManifestClaimsCompleteButOmitsGrantHolder` the manifest declares itself complete and omits a principal that holds assignments. The exhaustiveness claim is contradicted by the tenant.
`ManifestDeclarationMismatch` the manifest and `requiredResourceAccess` disagree.
`IntentEntryMalformed` an entry declared `complete` or `expectedGranted` as something other than a JSON boolean, and was skipped.
`IntentEntryNotResolved`, `IntentEntryDuplicate`, `IntentPrincipalAbsent` manifest entries that could not be used.
`DuplicateApplication` several local registrations carry the same `appId`.

## Limits

Application permissions only.

`appRoleId` may be the null GUID `00000000-0000-0000-0000-000000000000`, which is valid and means the principal is assigned to the resource without a specific role. It is named rather than treated as an unknown role.

Graph does not enforce which side an assignment is addressed through: creation and revocation are both accepted against either the client or the resource collection. Nothing about the collection used can be inferred from the object.

The tool reads. It does not revoke, does not recommend a revocation, and does not rank findings by severity.

## Structure

```
Get-AppGrantDiff.ps1     the analyzer, read-only
samples/intent.json      manifest format, hand-written example
samples/lab-intent.json  the manifest the shipped report actually ran against
samples/output.json      a real report, not a reconstruction
lab/                     scripts that build a test corpus in a test tenant, write
```

`samples/output.json` records the SHA256 of `samples/lab-intent.json` as the
provenance of that run. The pair verifies itself:

```powershell
foreach ($f in '.\samples\lab-intent.json', '.\samples\output.json') {
    if (-not (Test-Path -LiteralPath $f)) { throw "missing: $f" }
}
$hash = (Get-FileHash .\samples\lab-intent.json -Algorithm SHA256).Hash
$report = Get-Content .\samples\output.json -Raw | ConvertFrom-Json
"$($report.intentSource.sha256 -eq $hash) - expected True, hash $hash"
```

The existence check is not ceremony. Without it, a missing file makes both sides
of the comparison `$null`, and the check prints `True` while measuring nothing.
A test that passes when its inputs are absent is the failure this tool exists to
report, turned on the tool itself.

## Licence

MIT.
