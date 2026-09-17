#requires -Version 7.0
<#
    agd-lab-analyse.ps1

    LECTURE SEULE. Remplace agd-lab-verifie.ps1.

    Émet le schéma et le vocabulaire de scope-gap : evaluation[] avec les états
    CorrectCoverage / UnderCoverage / OverCoverage / CorrectExclusion /
    NotInManifest / Observed, et trois axes de rapport séparés,
    assessment.completeness, observation.freshness, conclusion.result.

    AXE D'ÉNUMÉRATION : les principaux de service, pas les inscriptions.
    Une inscription ne peut pas mener aux principaux qui portent des grants sans
    inscription locale, or c'est la population que l'outil existe pour rendre
    visible. Une pagination $expand donne l'ensemble des porteurs, puis une
    lecture individuelle est faite uniquement là où un zéro porte une conclusion.

    RÈGLE CARDINALE : $expand omet la propriété appRoleAssignments pour les
    principaux qui n'en portent pas. Absence de propriété n'est pas zéro. Un
    principal dont l'état accordé n'a pas été observé ne reçoit aucune ligne :
    il dégrade assessment.completeness et produit un diagnostic.

    Sans -IntentPath : completeness Absent, les grants sortent en Observed, et
    la conclusion ne peut pas dépasser CoverageNotDemonstrable.
#>

[CmdletBinding()]
param(
    [string] $OutputDir,
    [string] $IntentPath,
    [string] $RegistrePath,
    [string] $Colonne
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$TOOL_VERSION = '0.1.0'
$ROLE_NUL     = '00000000-0000-0000-0000-000000000000'
$GRAPH_APPID  = '00000003-0000-0000-c000-000000000000'

if (-not $OutputDir) {
    $base = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
    $OutputDir = Join-Path $base 'agd-lab-out'
}
if (-not (Test-Path -LiteralPath $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }
if (-not $RegistrePath) { $RegistrePath = Join-Path $OutputDir 'agd-lab-registre.csv' }
Write-Host "Sortie  : $OutputDir"

# ── Garde-fou modules ───────────────────────────────────────────────────────

$loaded = @(Get-Module Microsoft.Graph.Authentication)
if ($loaded.Count -gt 1) { throw "Plusieurs versions de Microsoft.Graph.Authentication chargees. Processus neuf requis." }
$authAvailable = @(Get-Module -ListAvailable Microsoft.Graph.Authentication | Sort-Object Version -Descending)
$appsAvailable = @(Get-Module -ListAvailable Microsoft.Graph.Applications   | Sort-Object Version -Descending)
if ($authAvailable.Count -eq 0 -or $appsAvailable.Count -eq 0) { throw "Modules Graph absents du disque." }
$appsVersions = @($appsAvailable.Version)
$common = @($authAvailable.Version | Where-Object { $_ -in $appsVersions } | Sort-Object -Descending) | Select-Object -First 1
if (-not $common) { throw "Aucune version commune entre Authentication et Applications." }
if ($loaded.Count -eq 1 -and $loaded[0].Version -ne $common) {
    throw "Microsoft.Graph.Authentication $($loaded[0].Version) deja chargee, cible $common. Processus neuf requis."
}
Import-Module Microsoft.Graph.Authentication -RequiredVersion $common -Force
Import-Module Microsoft.Graph.Applications   -RequiredVersion $common -Force
Write-Host "Modules Graph pinnes en $common"

Connect-MgGraph -Scopes 'Application.Read.All','Directory.Read.All' -ContextScope Process -NoWelcome
$tenantId = (Get-MgContext).TenantId
Write-Host "Tenant  : $tenantId"

# ── Helpers de forme, éprouvés ──────────────────────────────────────────────

function Get-Champ {
    param($Item, [string]$Nom)
    if ($null -eq $Item) { return $null }
    if ($Item -is [System.Collections.IDictionary]) {
        if ($Item.Contains($Nom)) { return $Item[$Nom] }
        return $null
    }
    $p = $Item.PSObject.Properties[$Nom]
    if ($p) { return $p.Value }
    return $null
}

$diagnostics = [System.Collections.Generic.List[object]]::new()
function Add-Diagnostic {
    param([string]$Severity, [string]$Code, [string]$Object, [string]$Message)
    $diagnostics.Add([pscustomobject]@{
        severity = $Severity; code = $Code; object = $Object; message = $Message
    })
}

# ── 1. Manifeste d'intention ────────────────────────────────────────────────
#    Validation de contrat. Un outil de preuve n'infère pas sa propre référence.

$intentStatus       = 'Absent'
$intentComplete     = $false
$intentDescription  = $null
$intentAsOf         = $null
$intentHash         = $null
$intentDeclared     = 0
$intentNotResolved  = 0
$intentDuplicates   = 0
$intentParAppId     = @{}

if ($IntentPath) {
    if (-not (Test-Path -LiteralPath $IntentPath)) {
        # Renvoyer le chemin resolu et le repertoire courant. Repeter tel quel le
        # chemin relatif qu'on vient de recevoir n'apprend rien au lecteur.
        $resolu = [System.IO.Path]::GetFullPath($IntentPath, (Get-Location).Path)
        throw "Manifeste introuvable. Fourni : $IntentPath. Resolu en : $resolu. Repertoire courant : $((Get-Location).Path)."
    }

    $intentHash = (Get-FileHash -LiteralPath $IntentPath -Algorithm SHA256).Hash
    $raw = Get-Content -LiteralPath $IntentPath -Raw -Encoding utf8 | ConvertFrom-Json

    if ($null -eq $raw.PSObject.Properties['complete']) {
        throw "Le manifeste ne porte pas de champ 'complete'. Il est obligatoire et n'est jamais infere."
    }
    $intentComplete    = [bool]$raw.complete
    $intentStatus      = 'Available'
    $intentDescription = if ($raw.PSObject.Properties['description']) { $raw.description } else { $null }
    $intentAsOf        = if ($raw.PSObject.Properties['asOf']) { $raw.asOf } else { $null }

    foreach ($p in @($raw.principals)) {
        $intentDeclared++
        $label = if ($p.PSObject.Properties['displayName']) { $p.displayName } else { $p.appId }

        if ($null -eq $p.PSObject.Properties['complete']) {
            throw "L'entree '$label' ne porte pas de champ 'complete'. Il est obligatoire par principal."
        }
        if ($null -eq $p.PSObject.Properties['appId'] -or -not $p.appId -or $p.appId -match '^<') {
            $intentNotResolved++
            Add-Diagnostic 'Warning' 'IntentEntryNotResolved' $label "L'entree ne porte pas d'appId exploitable."
            continue
        }

        $perms = @()
        if ($p.PSObject.Properties['permissions']) {
            foreach ($q in @($p.permissions)) {
                $perms += [pscustomobject]@{
                    ResourceAppId   = $q.resourceAppId
                    Value           = $q.value
                    ExpectedGranted = [bool]$q.expectedGranted
                    Cle             = "$($q.resourceAppId)/$($q.value)"
                }
            }
        }

        if ($intentParAppId.ContainsKey($p.appId)) {
            $intentDuplicates++
            Add-Diagnostic 'Info' 'IntentEntryDuplicate' $label "Le principal apparait plus d'une fois dans le manifeste."
            continue
        }

        $intentParAppId[$p.appId] = [pscustomobject]@{
            DisplayName      = $label
            Complete         = [bool]$p.complete
            Permissions      = @($perms)
            PermissionsFournies = [bool]$p.PSObject.Properties['permissions']
        }
    }
    # Seule la resolution au niveau du fichier est connue ici. La resolution
    # contre le tenant vient plus tard : l'annoncer maintenant afficherait un
    # chiffre qu'une ligne ulterieure contredit.
    Write-Host "Manifeste : $intentDeclared entree(s) chargee(s), $($intentParAppId.Count) exploitable(s)"
} else {
    Write-Host "Manifeste : aucun. La conclusion ne pourra pas depasser CoverageNotDemonstrable."
}

# ── 2. Côté accordé : pagination $expand ────────────────────────────────────
#    Forme select_puis_expand, la plus rapide des trois mesurees.

Write-Host "`nLecture des attributions"

$grantsParSpId    = @{}   # spObjectId -> tableau d'attributions, ou absent si non observe
$spParSpId        = @{}
$spParAppId       = @{}
$signauxImbriques = @{}   # spObjectId -> nextLink / count annonces sur la collection developpee
$expandCompte     = @{}   # spObjectId -> nombre rendu par $expand, avant verification

$uri = 'https://graph.microsoft.com/v1.0/servicePrincipals?$select=id,appId,displayName,appOwnerOrganizationId&$expand=appRoleAssignments'
$sw = [System.Diagnostics.Stopwatch]::StartNew()
$pages = 0
$next = $uri
while ($next) {
    $page = Invoke-MgGraphRequest -Method GET -Uri $next -OutputType PSObject
    $pages++
    foreach ($v in @(Get-Champ -Item $page -Nom 'value')) {
        $id    = Get-Champ -Item $v -Nom 'id'
        $appId = Get-Champ -Item $v -Nom 'appId'
        $sp = [pscustomobject]@{
            Id          = $id
            AppId       = $appId
            DisplayName = Get-Champ -Item $v -Nom 'displayName'
            OwnerOrgId  = Get-Champ -Item $v -Nom 'appOwnerOrganizationId'
        }
        $spParSpId[$id] = $sp
        if ($appId) { $spParAppId[$appId] = $sp }

        # Absence de propriete n'est pas zero. On n'enregistre que ce qui est rendu.
        $ass = Get-Champ -Item $v -Nom 'appRoleAssignments'
        if ($null -ne $ass) {
            $grantsParSpId[$id] = @(@($ass) | ForEach-Object {
                [pscustomobject]@{
                    Id         = Get-Champ -Item $_ -Nom 'id'
                    AppRoleId  = Get-Champ -Item $_ -Nom 'appRoleId'
                    ResourceId = Get-Champ -Item $_ -Nom 'resourceId'
                }
            })
            # Graph annonce-t-il la suite de la collection imbriquee ? Si oui la
            # troncature est signalee et c'est le lecteur qui l'ignore. Si non
            # elle est muette. La difference se constate, elle ne se suppose pas.
            $lienImbrique = Get-Champ -Item $v -Nom 'appRoleAssignments@odata.nextLink'
            $compteImbrique = Get-Champ -Item $v -Nom 'appRoleAssignments@odata.count'
            if ($lienImbrique -or $compteImbrique) {
                $signauxImbriques[$id] = [pscustomobject]@{
                    nextLink = $lienImbrique; count = $compteImbrique
                }
            }
        }
    }
    $next = Get-Champ -Item $page -Nom '@odata.nextLink'
}
$sw.Stop()
Write-Host "$($spParSpId.Count) principaux, $pages page(s), $([math]::Round($sw.Elapsed.TotalSeconds,1)) s"
Write-Host "$($grantsParSpId.Count) principal(aux) avec relation rendue"

# ── 3. Côté déclaré : inscriptions locales ──────────────────────────────────

$declareParAppId = @{}
foreach ($app in @(Get-MgApplication -All -Property 'id,appId,displayName,requiredResourceAccess')) {
    $d = @()
    foreach ($r in @($app.RequiredResourceAccess)) {
        foreach ($a in @($r.ResourceAccess)) {
            if ($a.Type -eq 'Role') {
                $d += [pscustomobject]@{ ResourceAppId = $r.ResourceAppId; RoleId = $a.Id }
            }
        }
    }
    if ($declareParAppId.ContainsKey($app.AppId)) {
        Add-Diagnostic 'Warning' 'DuplicateApplication' $app.DisplayName "Plusieurs inscriptions locales portent le meme appId."
        continue
    }
    $declareParAppId[$app.AppId] = [pscustomobject]@{
        ObjectId = $app.Id; DisplayName = $app.DisplayName; Declare = @($d)
    }
}
Write-Host "$($declareParAppId.Count) inscription(s) locale(s)"

# ── 4. Périmètre, et lectures individuelles ciblées ─────────────────────────
#    Un principal entre dans le perimetre s'il porte des grants, s'il declare des
#    permissions applicatives, ou s'il figure au manifeste. Une lecture
#    individuelle n'est faite que la ou un zero porterait une conclusion.

$perimetre = [System.Collections.Generic.HashSet[string]]::new()
foreach ($id in $grantsParSpId.Keys) { if (@($grantsParSpId[$id]).Count -gt 0) { [void]$perimetre.Add($id) } }
foreach ($appId in $declareParAppId.Keys) {
    if ($declareParAppId[$appId].Declare.Count -gt 0 -and $spParAppId.ContainsKey($appId)) {
        [void]$perimetre.Add($spParAppId[$appId].Id)
    }
}
foreach ($appId in $intentParAppId.Keys) {
    if ($spParAppId.ContainsKey($appId)) { [void]$perimetre.Add($spParAppId[$appId].Id) }
    else {
        # Une entree designant un principal absent du tenant est NON RESOLUE,
        # au meme titre qu'une entree sans appId exploitable. Elle doit degrader
        # overDemontrable : un manifeste qui revendique l'exhaustivite en
        # designant l'inexistant n'est pas entierement resolu.
        $intentNotResolved++
        Add-Diagnostic 'Warning' 'IntentPrincipalAbsent' $intentParAppId[$appId].DisplayName "Le principal du manifeste n'existe pas dans le tenant. Compte comme non resolu."
    }
}

# REGLE : $expand etablit la PRESENCE d'attributions, jamais leur COMPLETUDE.
#
# Documente : pour les ressources Entra derivant de directoryObject, $expand rend
# typiquement au maximum 20 elements de la relation developpee, et sans
# @odata.nextLink. Le servicePrincipal en derive.
# learn.microsoft.com/graph/query-parameters#expand
#
# Verifie en tenant le 08/09/2026 : un principal portant 25 attributions n'en a
# rendu que 20 par $expand, les cinq manquantes dispersees dans la serie, sans
# lien de continuation ni compte annonce.
#
# La limite est donc connue et la charge utile ne la signale pas localement. La
# seule defense est de relire. Toute collection developpee non vide est relue
# individuellement avant usage, et le cout reste borne au nombre de principaux
# porteurs, quatre sur trois cent quatre ici.

$lecturesCiblees   = 0
$lecturesCompletude = 0
$tronquees         = 0
$nonObserves = [System.Collections.Generic.List[string]]::new()
# Principaux dont les attributions sont PRESENTES mais dont la completude n'a
# pas pu etre verifiee, la relecture ayant echoue. Leurs grants observes sont
# des faits, leurs absences n'en sont pas.
$grantsNonVerifies = [System.Collections.Generic.HashSet[string]]::new()

foreach ($id in @($perimetre)) {

    $dejaRendu = $grantsParSpId.ContainsKey($id)
    if ($dejaRendu) {
        $expandCompte[$id] = @($grantsParSpId[$id]).Count
        if ($expandCompte[$id] -eq 0) { continue }
        $lecturesCompletude++
    } else {
        $lecturesCiblees++
    }

    try {
        $g = @(Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $id -All)
        $complet = @($g | ForEach-Object {
            [pscustomobject]@{ Id = $_.Id; AppRoleId = $_.AppRoleId; ResourceId = $_.ResourceId }
        })

        if ($dejaRendu -and $complet.Count -ne $expandCompte[$id]) {
            $tronquees++
            $signale = $signauxImbriques.ContainsKey($id)
            Add-Diagnostic 'Warning' 'ExpandCollectionTruncated' $spParSpId[$id].DisplayName `
                ("`$expand a rendu $($expandCompte[$id]) attribution(s) sur $($complet.Count) reelles. " +
                 $(if ($signale) { "Graph a signale la suite : $($signauxImbriques[$id] | ConvertTo-Json -Compress)." }
                   else { "Aucun lien de continuation ni compte annonce dans la charge utile. Comportement documente : `$expand rend typiquement au maximum 20 elements pour une relation developpee sur une ressource derivant de directoryObject, sans @odata.nextLink. Voir learn.microsoft.com/graph/query-parameters#expand." }))
        }

        $grantsParSpId[$id] = $complet
    } catch {
        if ($dejaRendu) {
            # La collection developpee vaut preuve de PRESENCE. Elle ne vaut pas
            # preuve d'absence : $expand plafonne les relations developpees, donc
            # une permission absente du resultat peut exister sans avoir ete
            # rendue. Garder les grants et evaluer les absences contre eux
            # transformerait un etat non verifie en etat observe.
            [void]$grantsNonVerifies.Add($id)
            Add-Diagnostic 'Warning' 'CompletenessNotVerified' $spParSpId[$id].DisplayName `
                "Relecture impossible, le resultat `$expand n'a pas pu etre verifie, sa completude est inconnue : $($_.Exception.Message)"
        } else {
            $nonObserves.Add($id)
            Add-Diagnostic 'Error' 'GrantStateNotObserved' $spParSpId[$id].DisplayName `
                "Lecture des attributions impossible : $($_.Exception.Message)"
        }
    }
}
Write-Host "$($perimetre.Count) principal(aux) dans le perimetre"
Write-Host "$lecturesCiblees lecture(s) pour confirmer un zero, $lecturesCompletude pour verifier la completude"
if ($tronquees -gt 0) { Write-Warning "$tronquees collection(s) `$expand tronquee(s)." }

# ── 5. Résolution des noms de rôles ─────────────────────────────────────────

$rolesParAppId = @{}
function Get-ValeurRole {
    param([string]$ResourceAppId, [string]$RoleId)
    if ($RoleId -eq $ROLE_NUL) { return '(attribue sans role specifique)' }
    if (-not $rolesParAppId.ContainsKey($ResourceAppId)) {
        $r = @(Get-MgServicePrincipal -Filter "appId eq '$ResourceAppId'" -Property 'id,appId,displayName,appRoles')
        $rolesParAppId[$ResourceAppId] = if ($r.Count -eq 1) { $r[0] } else { $null }
    }
    $sp = $rolesParAppId[$ResourceAppId]
    if (-not $sp) { return "(role $RoleId sur ressource $ResourceAppId absente du tenant)" }
    $m = @($sp.AppRoles | Where-Object { $_.Id -eq $RoleId })
    if ($m.Count -eq 0) { return "(role $RoleId non expose par $($sp.DisplayName))" }
    $m[0].Value
}

function Get-AppIdRessource {
    param([string]$ResourceSpId)
    if (-not $spParSpId.ContainsKey($ResourceSpId)) { return "(sp $ResourceSpId inconnu)" }
    $spParSpId[$ResourceSpId].AppId
}

# ── 6. Matrice d'évaluation ─────────────────────────────────────────────────

$evaluation = [System.Collections.Generic.List[object]]::new()
$completenessReasons = [System.Collections.Generic.List[string]]::new()
$nonJugeables = [System.Collections.Generic.List[string]]::new()
$sansLigne    = [System.Collections.Generic.List[object]]::new()
$correctlyEmptyCount = 0

if ($intentStatus -eq 'Absent')                      { $completenessReasons.Add('Aucun manifeste d''intention n''a ete fourni.') }
if ($intentStatus -eq 'Available' -and -not $intentComplete) { $completenessReasons.Add('Le manifeste se declare non exhaustif.') }
if ($intentNotResolved -gt 0)                        { $completenessReasons.Add("$intentNotResolved entree(s) du manifeste n'ont pas pu etre resolues.") }

foreach ($spId in @($perimetre)) {
    $sp = $spParSpId[$spId]
    $appId = $sp.AppId

    if ($nonObserves.Contains($spId)) {
        $completenessReasons.Add("L'etat accorde de '$($sp.DisplayName)' n'a pas ete observe.")
        continue
    }

    $entree = if ($appId -and $intentParAppId.ContainsKey($appId)) { $intentParAppId[$appId] } else { $null }
    $inscription = if ($appId -and $declareParAppId.ContainsKey($appId)) { $declareParAppId[$appId] } else { $null }

    # Côté attendu. requiredResourceAccess quand une inscription locale existe,
    # le manifeste quand il n'y en a pas. Le manifeste n'ecrase jamais
    # l'inscription : un desaccord est un diagnostic, pas une resolution.
    # requiredResourceAccess est une declaration POSITIVE UNIQUEMENT. Il n'a aucun
    # moyen d'exprimer "explicitement non attendu" : l'absence d'une entree y
    # signifie "non declare", pas "exclu". Une attente negative ne peut donc
    # venir que du manifeste, y compris pour un principal qui a une inscription.
    # Sans cette regle, CorrectExclusion est inatteignable des qu'une inscription
    # existe, et l'etat n'est jamais produit.

    $attendu = @{}
    $declarePositif = @{}

    if ($inscription) {
        foreach ($d in $inscription.Declare) {
            $v = Get-ValeurRole $d.ResourceAppId $d.RoleId
            $declarePositif["$($d.ResourceAppId)/$v"] = $true
            $attendu["$($d.ResourceAppId)/$v"] = $true
        }
    }

    if ($entree -and $entree.PermissionsFournies) {
        foreach ($p in $entree.Permissions) {
            if (-not $inscription) {
                $attendu[$p.Cle] = $p.ExpectedGranted
            }
            elseif (-not $p.ExpectedGranted) {
                if ($declarePositif.ContainsKey($p.Cle)) {
                    Add-Diagnostic 'Warning' 'ManifestDeclarationMismatch' $sp.DisplayName `
                        "Le manifeste declare '$($p.Cle)' explicitement non attendu alors que l'inscription la declare. L'inscription prevaut, la contradiction est signalee."
                } else {
                    $attendu[$p.Cle] = $false
                }
            }
        }

        # Divergence sur les attentes positives, entre manifeste et inscription.
        if ($inscription) {
            $manifPositif = @{}
            foreach ($p in $entree.Permissions) { if ($p.ExpectedGranted) { $manifPositif[$p.Cle] = $true } }
            if ($manifPositif.Count -gt 0) {
                $seulInscription = @($declarePositif.Keys | Where-Object { -not $manifPositif.ContainsKey($_) })
                $seulManifeste   = @($manifPositif.Keys   | Where-Object { -not $declarePositif.ContainsKey($_) })
                if ($seulInscription.Count -gt 0 -or $seulManifeste.Count -gt 0) {
                    Add-Diagnostic 'Warning' 'ManifestDeclarationMismatch' $sp.DisplayName `
                        ("Le manifeste et requiredResourceAccess divergent sur les attentes positives. Inscription seule : $($seulInscription -join ', '). Manifeste seul : $($seulManifeste -join ', ').")
                }
            }
        }
    }

    # Côté observé.
    $observe = @{}
    foreach ($g in @($grantsParSpId[$spId])) {
        $ra = Get-AppIdRessource $g.ResourceId
        $v  = Get-ValeurRole $ra $g.AppRoleId
        $observe["$ra/$v"] = $true
    }

    # L'over-coverage n'est demontrable que si une reference exhaustive existe
    # pour CE principal. Sans manifeste, ou sans entree, ou entree non exhaustive,
    # un grant sans contrepartie reste NotInManifest, ou Observed s'il n'y a
    # aucun manifeste du tout.
    $overDemontrable = ($intentStatus -eq 'Available') -and $entree -and $entree.Complete -and ($intentNotResolved -eq 0)

    $aProduitNotInManifest = $false

    # Completude non verifiee : seules les lignes portant un grant observe sont
    # emises. Tout etat reposant sur une absence, UnderCoverage et
    # CorrectExclusion, est retenu et bascule dans assessment.reasons.
    $completudeNonVerifiee = $grantsNonVerifies.Contains($spId)
    $retenues = [System.Collections.Generic.List[string]]::new()
    $lignesAvant = $evaluation.Count

    foreach ($cle in @(@($attendu.Keys) + @($observe.Keys) | Select-Object -Unique)) {
        $exp = if ($attendu.ContainsKey($cle)) { [bool]$attendu[$cle] } else { $null }
        $obs = $observe.ContainsKey($cle)

        if ($completudeNonVerifiee -and -not $obs) { $retenues.Add($cle); continue }

        $state =
            if ($exp -eq $true  -and $obs)      { 'CorrectCoverage' }
            elseif ($exp -eq $true  -and -not $obs) { 'UnderCoverage' }
            elseif ($exp -eq $false -and -not $obs) { 'CorrectExclusion' }
            elseif ($exp -eq $false -and $obs)      { if ($overDemontrable) { 'OverCoverage' } else { 'NotInManifest' } }
            elseif ($intentStatus -eq 'Absent')     { 'Observed' }
            elseif ($overDemontrable)               { 'OverCoverage' }
            else                                    { 'NotInManifest' }

        if ($state -eq 'NotInManifest') { $aProduitNotInManifest = $true }

        $parts = $cle -split '/', 2
        $evaluation.Add([pscustomobject]@{
            principalId     = $spId
            principalAppId  = $appId
            displayName     = $sp.DisplayName
            resourceAppId   = $parts[0]
            permission      = $parts[1]
            expectedGranted = $exp
            observedGranted = $obs
            state           = $state
            source          = if ($null -eq $exp) { 'grant' } elseif ($inscription) { 'registration' } else { 'intent' }
            localRegistration = [bool]$inscription
        })
    }

    # Un principal du perimetre qui ne produit aucune ligne est invisible dans
    # l'evaluation, et le lecteur ne peut pas le distinguer d'un principal jamais
    # atteint. On rapporte les faits detenus, sans taxonomie. Qu'un objet soit une
    # ressource et non un client n'en fait pas partie : le moteur ne lit ni les
    # appRoles ni appRoleAssignedTo, le dire serait une inference tiree du nom.
    if ($evaluation.Count -eq $lignesAvant) {
        $attenduAUneSource = $entree -and ($entree.PermissionsFournies -or $inscription)
        $estCorrectlyEmpty = ($intentStatus -eq 'Available') -and $entree -and $entree.Complete `
                             -and $attenduAUneSource -and ($attendu.Count -eq 0) -and ($observe.Count -eq 0)
        if ($estCorrectlyEmpty) { $correctlyEmptyCount++ }
        $sansLigne.Add([pscustomobject]@{
            displayName         = $sp.DisplayName
            principalAppId      = $appId
            localRegistration   = [bool]$inscription
            declaredPermissions = $declarePositif.Count
            observedGrants      = $observe.Count
            inManifest          = [bool]$entree
            manifestComplete    = if ($entree) { $entree.Complete } else { $null }
            correctlyEmpty      = $estCorrectlyEmpty
        })
    }

    if ($completudeNonVerifiee) {
        $detail = if ($retenues.Count -gt 0) { " Retenues : $($retenues -join ', ')." } else { '' }
        $completenessReasons.Add("La completude des attributions de '$($sp.DisplayName)' n'a pas pu etre verifiee, aucune absence n'a donc ete etablie pour lui.$detail")
    }

    # Ne nommer un principal que s'il a reellement produit une ligne non
    # jugeable. Une reference degradee ne rend pas tout porteur de grants non
    # juge : un principal dont chaque attribution correspond a une declaration
    # sort en CorrectCoverage, et le lister ici surestimerait la phrase.
    if ($intentStatus -eq 'Available' -and $aProduitNotInManifest) {
        $nonJugeables.Add($sp.DisplayName)
        if ($intentComplete -and -not $entree) {
            Add-Diagnostic 'Warning' 'ManifestClaimsCompleteButOmitsGrantHolder' $sp.DisplayName `
                "Le manifeste se declare exhaustif et ne contient pas ce principal, qui porte des attributions. La revendication d'exhaustivite est contredite par le tenant."
        }
    }
}

# ── 7. Trois axes ───────────────────────────────────────────────────────────

# summary a deux etages : ses compteurs ne comptent pas la meme chose et ne
# couvrent pas la meme population. perPermission compte des lignes d'evaluation,
# une par permission, sur tous les principaux lus. perEntry compte des ENTREES de
# manifeste, et n'existe que si un manifeste a ete fourni. A plat, un lecteur ou
# un agregateur pourrait sommer huit nombres en un total qui ne designe rien.
#
# Aucun des deux compteurs perEntry n'est un etat de ligne : tous deux comptent
# des entrees qui n'ont produit aucune ligne. Ils figurent quand meme, parce que
# summary est la table ou un lecteur compte ce qui s'est passe.

$etats = @('CorrectCoverage','CorrectExclusion','UnderCoverage','OverCoverage','NotInManifest','Observed')
$perPermission = [ordered]@{}
foreach ($e in $etats) { $perPermission[$e] = @($evaluation | Where-Object { $_.state -eq $e }).Count }

# Sans manifeste, perEntry est depourvu de sens et non vide de valeur : il n'y a
# aucune entree a compter. Deux zeros se liraient comme deux verifications passees.
$perEntry = if ($intentStatus -eq 'Available') {
    [pscustomobject][ordered]@{ NotResolved = $intentNotResolved; CorrectlyEmpty = $correctlyEmptyCount }
} else { $null }

if ($nonJugeables.Count -gt 0) {
    $noms = @($nonJugeables | Sort-Object -Unique)
    $extrait = if ($noms.Count -le 5) { $noms -join ', ' } else { (@($noms)[0..4] -join ', ') + ", et $($noms.Count - 5) autre(s)" }
    $completenessReasons.Add("$($noms.Count) principal(aux) portent des attributions que la reference ne permet pas de juger : $extrait.")
}

$reasons = @($completenessReasons | Select-Object -Unique)
$completeness = if ($intentStatus -eq 'Absent') { 'Absent' }
                elseif ($reasons.Count -gt 0)   { 'Partial' }
                else                            { 'Complete' }

# Axe 2. Graph documente des delais de replication sur appRoleAssignments et
# n'expose aucun indicateur de convergence. La fraicheur n'est donc jamais
# affirmee : elle n'est pas demontree.
$observation = [pscustomobject]@{
    replicationIndicator = 'NotExposed'
    expandCompleteness   = if ($tronquees -gt 0) { 'TruncatedAndRepaired' } else { 'NotContradicted' }
    freshness            = 'NotDemonstrated'
    note                 = 'Graph signale des delais de replication sur les attributions et n''expose aucun indicateur de convergence. Que l''instantane lu soit a jour n''a pas ete etabli.'
}

$conclusion =
    if ($hasGap) {
        $parts = @()
        if ($perPermission['UnderCoverage'] -gt 0) { $parts += "$($perPermission['UnderCoverage']) permission(s) declaree(s) et non accordee(s)" }
        if ($perPermission['OverCoverage']  -gt 0) { $parts += "$($perPermission['OverCoverage']) permission(s) accordee(s) et non attendue(s)" }
        [pscustomobject]@{
            result            = 'GapEstablished'
            detail            = ($parts -join '. ') + '.'
            gapBreakCondition = 'Le manifeste est perime ou se trompe sur ce qui est attendu pour ces principaux, ou l''instantane des attributions n''etait pas a jour a la generation.'
            impactBoundary    = 'Les permissions deleguees, les consentements utilisateur, l''appartenance a des roles d''annuaire et les identites managees ne sont pas evalues. Un ecart ici n''etablit pas ce a quoi le principal accede reellement.'
        }
    }
    elseif ($completeness -ne 'Complete') {
        [pscustomobject]@{
            result            = 'CoverageNotDemonstrable'
            detail            = 'Aucun ecart n''a ete etabli, et l''evaluation n''a pas pu couvrir toute la population. ' + ($reasons -join ' ')
            gapBreakCondition = $null
            impactBoundary    = $null
        }
    }
    else {
        [pscustomobject]@{
            result            = 'NoGapEstablished'
            detail            = 'Aucun ecart observe. Le manifeste etait exhaustif et entierement resolu. La fraicheur des attributions n''a pas ete etablie independamment.'
            gapBreakCondition = 'Le manifeste est perime, ou il se declare exhaustif en omettant une partie de la population reelle, ou les attributions n''avaient pas fini de converger a la lecture.'
            impactBoundary    = $null
        }
    }

# ── 8. Sortie ───────────────────────────────────────────────────────────────

$rapport = [pscustomobject]@{
    metadata = [pscustomobject]@{
        toolVersion       = $TOOL_VERSION
        generatedAt       = (Get-Date).ToUniversalTime().ToString('o')
        powerShellVersion = $PSVersionTable.PSVersion.ToString()
        tenantId          = $tenantId
        disclaimer        = 'Les conclusions sont bornees aux permissions applicatives. Les permissions deleguees ne sont pas evaluees.'
    }
    intentSource = [pscustomobject]@{
        status              = $intentStatus
        path                = $IntentPath
        description         = $intentDescription
        asOf                = $intentAsOf
        sha256              = $intentHash
        complete            = if ($intentStatus -eq 'Available') { $intentComplete } else { $null }
        declaredObjectCount = $intentDeclared
        resolvedObjectCount = $intentParAppId.Count
        notResolved         = $intentNotResolved
        duplicatesIgnored   = $intentDuplicates
    }
    read = [pscustomobject]@{
        principalsTotal      = $spParSpId.Count
        expandPages          = $pages
        expandDurationMs     = $sw.ElapsedMilliseconds
        relationRendered     = $expandCompte.Count
        scopedPrincipals     = $perimetre.Count
        zeroConfirmingReads  = $lecturesCiblees
        completenessReads    = $lecturesCompletude
        expandTruncated      = $tronquees
        grantStateNotObserved = $nonObserves.Count
        grantCompletenessUnverified = $grantsNonVerifies.Count
        evaluatedWithoutRows        = $sansLigne.ToArray()
        grantHoldersNotJudgeable = @($nonJugeables | Sort-Object -Unique).Count
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

$jsonPath = Join-Path $OutputDir 'agd-lab-rapport.json'
$rapport | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $jsonPath -Encoding utf8

$evaluation | Format-Table displayName, permission, expectedGranted, observedGranted, state -AutoSize
Write-Host "completeness : $completeness"
if ($reasons.Count -gt 0) { foreach ($r in $reasons) { Write-Host "  - $r" } }
Write-Host "freshness    : $($observation.freshness)"
Write-Host "conclusion   : $($conclusion.result)"
Write-Host "               $($conclusion.detail)"
if ($diagnostics.Count -gt 0) {
    Write-Host "`nDiagnostics :"
    $diagnostics | Format-Table severity, code, object, message -AutoSize -Wrap
}

# ── 9. Confrontation au registre ────────────────────────────────────────────
#    Le registre reste la ligne de base ecrite a la main. Il n'est pas derive
#    de ce script, sans quoi les deux partageraient la meme erreur.

if (Test-Path -LiteralPath $RegistrePath) {
    $colonne = if ($Colonne) { $Colonne }
               elseif ($intentStatus -eq 'Available') { 'state_attendu_avec_manifeste' }
               else { 'state_attendu_sans_manifeste' }
    $registre = @(Import-Csv -LiteralPath $RegistrePath)
    if ($registre.Count -gt 0 -and $registre[0].PSObject.Properties[$colonne]) {
        Write-Host "`nConfrontation au registre, colonne $colonne :"
        $verdicts = foreach ($r in $registre) {
            $obs = @($evaluation | Where-Object { $_.displayName -eq $r.nom })
            $etatsObs = @($obs | ForEach-Object { $_.state } | Select-Object -Unique)
            $observe = if ($etatsObs.Count -eq 0) { '(aucune ligne)' } else { $etatsObs -join '+' }
            [pscustomobject]@{
                cas = $r.cas; nom = $r.nom; attendu = $r.$colonne; observe = $observe
                verdict = if ($observe -eq $r.$colonne) { 'conforme' } else { 'ECART' }
            }
        }
        @($verdicts) | Format-Table cas, nom, attendu, observe, verdict -AutoSize
        $ecarts = @(@($verdicts) | Where-Object { $_.verdict -ne 'conforme' })
        if ($ecarts.Count -gt 0) { Write-Warning "$($ecarts.Count) cas non conforme(s)." }
    } else {
        Write-Warning "Le registre ne porte pas de colonne '$colonne'. Confrontation sautee."
    }
}

Write-Host "`nRapport : $jsonPath"
