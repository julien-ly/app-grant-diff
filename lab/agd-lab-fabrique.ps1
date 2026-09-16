#requires -Version 7.0
<#
    agd-lab-fabrique.ps1

    Fabrique les objets de test de app-grant-diff et produit le registre.
    ECRIT dans le tenant. A n'executer que dans un tenant de lab.

    Sequence attendue :
      pwsh -NoProfile
      .\agd-lab-preflight.ps1     (bloc separe, ci-dessous en commentaire)
      .\agd-lab-fabrique.ps1

    Sortie : agd-lab-registre.csv et agd-lab-registre.json
             agd-lab-observations.json  (faits bruts a examiner, pas des conclusions)
#>

[CmdletBinding()]
param(
    [string] $Prefix     = 'agd-lab',
    [string] $OutputDir,
    [int]    $MaxRetries = 12,
    [int]    $RetryDelay = 5
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# ---------------------------------------------------------------------------
# -1. SORTIE : RESOLUE ET TESTEE EN PREMIER
#     Le repertoire par defaut suit le script, pas le repertoire courant :
#     une session lancee depuis C:\Windows\System32 ne doit pas decider ou
#     s'ecrit le registre. Le test d'ecriture passe avant Connect-MgGraph
#     pour qu'un chemin non inscriptible ne consomme pas une authentification
#     interactive et n'echoue pas apres avoir cree des objets dans le tenant.
# ---------------------------------------------------------------------------

if (-not $OutputDir) {
    $base = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
    $OutputDir = Join-Path $base 'agd-lab-out'
}

try {
    if (-not (Test-Path -LiteralPath $OutputDir)) {
        New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
    }
    $probe = Join-Path $OutputDir '.agd-write-test'
    Set-Content -LiteralPath $probe -Value 'ok' -Encoding utf8
    Remove-Item -LiteralPath $probe -Force
} catch {
    throw "Repertoire de sortie inutilisable : $OutputDir. $($_.Exception.Message) Relancer avec -OutputDir sur un chemin inscriptible."
}

Write-Host "Sortie  : $OutputDir"

# ---------------------------------------------------------------------------
# 0. PREFLIGHT MODULES
#    A executer en session vierge AVANT ce script. Laisse ici pour reference.
#
#    # Ce qui est charge dans la session courante :
#    Get-Module Microsoft.Graph* | Select-Object Name, Version
#
#    # Ce qui est present sur disque, toutes versions, avec le chemin.
#    # C'est cette commande qui revele le side-by-side, pas Get-InstalledModule,
#    # qui ne connait que ce que PowerShellGet a enregistre.
#    Get-Module -ListAvailable Microsoft.Graph* |
#        Select-Object Name, Version, Path | Sort-Object Name, Version
#
#    # Desinstallation : meta-module d'abord, Authentication en dernier.
#    Uninstall-Module Microsoft.Graph -AllVersions -ErrorAction SilentlyContinue
#    Get-InstalledModule Microsoft.Graph.* |
#        Where-Object Name -ne 'Microsoft.Graph.Authentication' |
#        Uninstall-Module -AllVersions
#    Uninstall-Module Microsoft.Graph.Authentication -AllVersions
#
#    # Uninstall-Module echoue sur ce qui n'a pas ete pose par PowerShellGet.
#    # Verifier les residus sur disque avant de reinstaller :
#    $env:PSModulePath.Split([IO.Path]::PathSeparator) |
#        ForEach-Object { Get-ChildItem $_ -Filter 'Microsoft.Graph*' -EA SilentlyContinue }
#
#    # Reinstallation : deux sous-modules seulement, version identique.
#    $v = (Find-Module Microsoft.Graph.Authentication -Repository PSGallery).Version
#    Install-Module Microsoft.Graph.Authentication -RequiredVersion $v -Scope CurrentUser -Repository PSGallery
#    Install-Module Microsoft.Graph.Applications   -RequiredVersion $v -Scope CurrentUser -Repository PSGallery
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# 1. GARDE-FOU DE SESSION
#    Le conflit vient de l'ordre de chargement : la premiere version de
#    Microsoft.Graph.Authentication liee dans le processus gagne. On la charge
#    donc explicitement en premier, a une version pinnee, et on refuse de
#    continuer si une autre version est deja en memoire.
# ---------------------------------------------------------------------------

$loaded = @(Get-Module Microsoft.Graph.Authentication)
if ($loaded.Count -gt 1) {
    throw "Plusieurs versions de Microsoft.Graph.Authentication chargees : $(($loaded.Version) -join ', '). Relancer dans un processus neuf."
}

$authAvailable = @(Get-Module -ListAvailable Microsoft.Graph.Authentication |
                   Sort-Object Version -Descending)
$appsAvailable = @(Get-Module -ListAvailable Microsoft.Graph.Applications |
                   Sort-Object Version -Descending)

if ($authAvailable.Count -eq 0) { throw "Microsoft.Graph.Authentication absent du disque." }
if ($appsAvailable.Count -eq 0) { throw "Microsoft.Graph.Applications absent du disque." }

# Version commune la plus haute presente pour les deux modules.
$appsVersions = @($appsAvailable.Version)
$common = @($authAvailable.Version |
            Where-Object { $_ -in $appsVersions } |
            Sort-Object -Descending) | Select-Object -First 1

if (-not $common) {
    $msg  = "Aucune version commune. Authentication : $(($authAvailable.Version) -join ', '). "
    $msg += "Applications : $(($appsAvailable.Version) -join ', '). "
    $msg += "C'est la cause directe du conflit d'assembly."
    throw $msg
}

if ($loaded.Count -eq 1 -and $loaded[0].Version -ne $common) {
    throw "Microsoft.Graph.Authentication $($loaded[0].Version) deja chargee, version cible $common. Relancer dans un processus neuf."
}

Import-Module Microsoft.Graph.Authentication -RequiredVersion $common -Force
Import-Module Microsoft.Graph.Applications   -RequiredVersion $common -Force

Write-Host "Modules Graph pinnes en $common"

# ---------------------------------------------------------------------------
# 2. CONNEXION
#    ContextScope Process : plusieurs tenants ouverts en parallele.
#    Scopes conformes a la doc de reference pour la fabrication d'appRoleAssignment.
# ---------------------------------------------------------------------------

Connect-MgGraph -Scopes 'Application.ReadWrite.All','AppRoleAssignment.ReadWrite.All' `
                -ContextScope Process -NoWelcome

$ctx = Get-MgContext
Write-Host "Tenant  : $($ctx.TenantId)"
Write-Host "Compte  : $($ctx.Account)"

$observations = [System.Collections.Generic.List[object]]::new()
function Add-Observation {
    param([string]$Cas, [string]$Fait, $Valeur)
    $observations.Add([pscustomobject]@{
        horodatage = (Get-Date).ToUniversalTime().ToString('o')
        cas        = $Cas
        fait       = $Fait
        valeur     = $Valeur
    })
}

# ---------------------------------------------------------------------------
# 3. RESSOURCE ET ROLES
#    resourceAppId cote declare est un appId. resourceId cote grant est un
#    objectId de principal de service. Les deux ne sont pas interchangeables :
#    c'est la jointure que le moteur devra faire.
# ---------------------------------------------------------------------------

$graphAppId = '00000003-0000-0000-c000-000000000000'
$graphSp = Get-MgServicePrincipal -Filter "appId eq '$graphAppId'" -Property 'id,appId,displayName,appRoles'
if (-not $graphSp) { throw "Principal de service Microsoft Graph introuvable dans le tenant." }

function Get-AppRoleId {
    param([string]$Value)
    $role = $graphSp.AppRoles | Where-Object { $_.Value -eq $Value -and $_.AllowedMemberTypes -contains 'Application' }
    if (-not $role) { throw "Role applicatif '$Value' introuvable sur le SP Microsoft Graph." }
    $role.Id
}

$roleUserReadAll  = Get-AppRoleId 'User.Read.All'
$roleGroupReadAll = Get-AppRoleId 'Group.Read.All'

Add-Observation 'commun' 'graph_sp_object_id' $graphSp.Id
Add-Observation 'commun' 'graph_app_id'       $graphSp.AppId
Add-Observation 'commun' 'role_user_read_all' $roleUserReadAll
Add-Observation 'commun' 'role_group_read_all' $roleGroupReadAll

# ---------------------------------------------------------------------------
# 4. HELPERS
# ---------------------------------------------------------------------------

function New-LabApp {
    param(
        [string]   $Suffix,
        [string[]] $DeclaredRoleIds = @()
    )
    $name = "$Prefix-$Suffix"

    $rra = @()
    if ($DeclaredRoleIds.Count -gt 0) {
        $rra = @(
            @{
                resourceAppId  = $graphAppId
                resourceAccess = @($DeclaredRoleIds | ForEach-Object { @{ id = $_; type = 'Role' } })
            }
        )
    }

    $body = @{
        displayName            = $name
        signInAudience         = 'AzureADMyOrg'
        requiredResourceAccess = $rra
    }

    $app = New-MgApplication -BodyParameter $body
    $sp  = New-MgServicePrincipal -AppId $app.AppId

    [pscustomobject]@{
        Nom         = $name
        AppId       = $app.AppId
        AppObjectId = $app.Id
        SpObjectId  = $sp.Id
    }
}

function Grant-LabRole {
    param([string]$ClientSpId, [string]$RoleId)

    # ServicePrincipalId doit valoir l'objectId du SP RESSOURCE, pas du client.
    # C'est la faute la plus courante et elle ne produit pas d'erreur lisible.
    $params = @{
        principalId = $ClientSpId
        resourceId  = $graphSp.Id
        appRoleId   = $RoleId
    }
    New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $graphSp.Id -BodyParameter $params
}

function Wait-ForAssignmentCount {
    <#
        Relit jusqu'a obtenir le compte attendu. Journalise la latence.
        L'acceptation de l'ecriture ne prouve pas l'etat : la doc Graph
        signale des delais de replication sur ces collections.

        Retour sans virgule. Les sites d'appel enveloppent dans @(), ce qui
        donne le bon compte pour zero, un et n elements. Avec ,$tableau le
        site d'appel recoit un tableau imbrique et .Count vaut toujours 1,
        y compris quand la collection reelle est vide.
    #>
    param([string]$Cas, [string]$ClientSpId, [int]$Expected)

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    for ($i = 0; $i -lt $MaxRetries; $i++) {
        $assignments = @(Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ClientSpId -All)
        if ($assignments.Count -eq $Expected) {
            $sw.Stop()
            Add-Observation $Cas 'latence_lecture_ms' $sw.ElapsedMilliseconds
            Add-Observation $Cas 'tentatives_lecture' ($i + 1)
            return $assignments
        }
        Start-Sleep -Seconds $RetryDelay
    }
    $sw.Stop()
    Add-Observation $Cas 'lecture_non_convergente' "attendu=$Expected observe=$($assignments.Count) apres $($sw.Elapsed)"
    return $assignments
}

function Get-DeclaredRoles {
    param([string]$AppObjectId)
    # requiredResourceAccess n'est pas toujours retourne par defaut sur les
    # endpoints de liste. On le demande explicitement.
    $app = Get-MgApplication -ApplicationId $AppObjectId -Property 'id,appId,displayName,requiredResourceAccess'
    $out = @()
    foreach ($r in @($app.RequiredResourceAccess)) {
        foreach ($a in @($r.ResourceAccess)) {
            if ($a.Type -eq 'Role') {
                $out += [pscustomobject]@{ ResourceAppId = $r.ResourceAppId; RoleId = $a.Id }
            }
        }
    }
    $out
}

$spCache = @{}
function Resolve-RoleName {
    # Le repli du cas C decouvre des attributions dont la ressource n'est pas
    # forcement Microsoft Graph. Resoudre le role sur le SP ressource porte par
    # l'attribution, et rendre un libelle explicite plutot que rien quand la
    # resolution echoue : un nom absent ne doit pas passer pour un role vide.
    param($Assignment)

    $rid = $Assignment.ResourceId
    if (-not $spCache.ContainsKey($rid)) {
        try {
            $spCache[$rid] = Get-MgServicePrincipal -ServicePrincipalId $rid `
                             -Property 'id,appId,displayName,appRoles'
        } catch {
            $spCache[$rid] = $null
        }
    }

    $sp = $spCache[$rid]
    if (-not $sp) { return "(role $($Assignment.AppRoleId) sur ressource $rid non lisible)" }

    $role = @($sp.AppRoles | Where-Object { $_.Id -eq $Assignment.AppRoleId })
    if ($role.Count -eq 0) { return "(role $($Assignment.AppRoleId) non expose par $($sp.DisplayName))" }

    $role[0].Value
}

# ---------------------------------------------------------------------------
# 5. FABRICATION
# ---------------------------------------------------------------------------

$registre = [System.Collections.Generic.List[object]]::new()

# --- Cas A : declare non accorde -------------------------------------------
Write-Host "`nCas A : declare non accorde"
$a = New-LabApp -Suffix 'a-declare-non-accorde' -DeclaredRoleIds @($roleUserReadAll)
Add-Observation 'A' 'app_object_id' $a.AppObjectId
Add-Observation 'A' 'sp_object_id'  $a.SpObjectId
$aGrants = @(Wait-ForAssignmentCount -Cas 'A' -ClientSpId $a.SpObjectId -Expected 0)

$registre.Add([pscustomobject]@{
    cas                       = 'A'
    nom                       = $a.Nom
    app_id                    = $a.AppId
    app_object_id             = $a.AppObjectId
    sp_object_id              = $a.SpObjectId
    declare                   = 'User.Read.All'
    accorde                   = ''
    state_attendu             = 'declared_not_granted'
    evaluation_status_attendu = 'evaluated'
})

# --- Cas B : accorde non declare -------------------------------------------
# Declaration puis grant puis retrait de la declaration. La doc Microsoft pose
# que le retrait de la permission dans l'inscription ne revoque pas le grant.
# Ce script ne le suppose pas : il le relit apres coup.
Write-Host "`nCas B : accorde non declare"
$b = New-LabApp -Suffix 'b-accorde-non-declare' -DeclaredRoleIds @($roleGroupReadAll)
Grant-LabRole -ClientSpId $b.SpObjectId -RoleId $roleGroupReadAll | Out-Null
$bBefore = @(Wait-ForAssignmentCount -Cas 'B' -ClientSpId $b.SpObjectId -Expected 1)
Add-Observation 'B' 'grants_avant_retrait_declaration' $bBefore.Count

Update-MgApplication -ApplicationId $b.AppObjectId -BodyParameter @{ requiredResourceAccess = @() }

$bDeclaredAfter = @(Get-DeclaredRoles -AppObjectId $b.AppObjectId)
$bAfter = @(Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $b.SpObjectId -All)
Add-Observation 'B' 'declaration_apres_retrait' $bDeclaredAfter.Count
Add-Observation 'B' 'grants_apres_retrait'      $bAfter.Count
Add-Observation 'B' 'grant_survit_au_retrait'   ($bAfter.Count -eq 1 -and $bDeclaredAfter.Count -eq 0)

$registre.Add([pscustomobject]@{
    cas                       = 'B'
    nom                       = $b.Nom
    app_id                    = $b.AppId
    app_object_id             = $b.AppObjectId
    sp_object_id              = $b.SpObjectId
    declare                   = ''
    accorde                   = 'Group.Read.All'
    state_attendu             = 'granted_not_declared'
    evaluation_status_attendu = 'evaluated'
})

# --- Cas C : retire ---------------------------------------------------------
# Le cas C se fabriquait ici en supprimant l'inscription pour garder le
# principal. Ce comportement n'existe pas : Learn documente que supprimer
# l'inscription dans son tenant d'origine supprime aussi le principal, et le
# tenant l'a confirme (couple appId 465b8e3b dans directory/deletedItems).
# Le cas C est desormais instancie par agd-lab-cas-c-v2.ps1, qui attribue un
# role de lab inerte a un principal dont appOwnerOrganizationId differe du
# tenant. Ne pas reintroduire ce bloc : chaque passage creait un second couple
# agd-lab-c-* qui rend la preuve de cascade ambigue.

# --- Cas D : alignement complet --------------------------------------------
Write-Host "`nCas D : alignement complet"
$d = New-LabApp -Suffix 'd-aligne' -DeclaredRoleIds @($roleUserReadAll)
Grant-LabRole -ClientSpId $d.SpObjectId -RoleId $roleUserReadAll | Out-Null
$dGrants = @(Wait-ForAssignmentCount -Cas 'D' -ClientSpId $d.SpObjectId -Expected 1)
Add-Observation 'D' 'grants' $dGrants.Count

$registre.Add([pscustomobject]@{
    cas                       = 'D'
    nom                       = $d.Nom
    app_id                    = $d.AppId
    app_object_id             = $d.AppObjectId
    sp_object_id              = $d.SpObjectId
    declare                   = 'User.Read.All'
    accorde                   = 'User.Read.All'
    state_attendu             = 'aligned'
    evaluation_status_attendu = 'evaluated'
})

# --- Cas E : controle negatif ----------------------------------------------
# Zero declare, zero accorde. Ce zero doit etre distinguable d'un module qui
# n'a rien evalue : state porte la conclusion, evaluation_status porte le fait
# que le moteur a pu conclure.
Write-Host "`nCas E : controle negatif"
$e = New-LabApp -Suffix 'e-controle-negatif'
$eDeclared = @(Get-DeclaredRoles -AppObjectId $e.AppObjectId)
$eGrants   = @(Wait-ForAssignmentCount -Cas 'E' -ClientSpId $e.SpObjectId -Expected 0)
Add-Observation 'E' 'declare' $eDeclared.Count
Add-Observation 'E' 'accorde' $eGrants.Count

$registre.Add([pscustomobject]@{
    cas                       = 'E'
    nom                       = $e.Nom
    app_id                    = $e.AppId
    app_object_id             = $e.AppObjectId
    sp_object_id              = $e.SpObjectId
    declare                   = ''
    accorde                   = ''
    state_attendu             = 'aligned_empty'
    evaluation_status_attendu = 'evaluated'
})

# ---------------------------------------------------------------------------
# 6. CONTRADICTION DOCUMENTAIRE A TRANCHER
#    La page de reference pose ServicePrincipalId = ResourceId a la creation,
#    puis passe l'objectId du client dans l'exemple de revocation. Les deux ne
#    peuvent pas etre vrais. On teste la forme "client" sur un grant jetable.
# ---------------------------------------------------------------------------

Write-Host "`nTest de la forme de revocation"
$z = New-LabApp -Suffix 'z-jetable' -DeclaredRoleIds @($roleUserReadAll)
$zAssign = Grant-LabRole -ClientSpId $z.SpObjectId -RoleId $roleUserReadAll
Wait-ForAssignmentCount -Cas 'Z' -ClientSpId $z.SpObjectId -Expected 1 | Out-Null

$formeClient = $null
try {
    Remove-MgServicePrincipalAppRoleAssignedTo -ServicePrincipalId $z.SpObjectId `
        -AppRoleAssignmentId $zAssign.Id -ErrorAction Stop
    $formeClient = 'acceptee'
} catch {
    $formeClient = "refusee : $($_.Exception.Message)"
}
Add-Observation 'Z' 'revocation_approleassignedto_avec_id_client' $formeClient

if ($formeClient -ne 'acceptee') {
    try {
        Remove-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $z.SpObjectId `
            -AppRoleAssignmentId $zAssign.Id -ErrorAction Stop
        Add-Observation 'Z' 'revocation_approleassignments_avec_id_client' 'acceptee'
    } catch {
        Add-Observation 'Z' 'revocation_approleassignments_avec_id_client' "refusee : $($_.Exception.Message)"
    }
}

# Acceptation != effet. On relit.
$zApres = @(Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $z.SpObjectId -All)
Add-Observation 'Z' 'grants_apres_revocation' $zApres.Count

Remove-MgApplication -ApplicationId $z.AppObjectId -ErrorAction SilentlyContinue
Remove-MgServicePrincipal -ServicePrincipalId $z.SpObjectId -ErrorAction SilentlyContinue

# ---------------------------------------------------------------------------
# 7. SORTIE
# ---------------------------------------------------------------------------

$csvPath  = Join-Path $OutputDir 'agd-lab-registre.csv'
$jsonPath = Join-Path $OutputDir 'agd-lab-registre.json'
$obsPath  = Join-Path $OutputDir 'agd-lab-observations.json'

$registre | Export-Csv -Path $csvPath -NoTypeInformation -Encoding utf8
$registre | ConvertTo-Json -Depth 6 | Set-Content -Path $jsonPath -Encoding utf8
$observations | ConvertTo-Json -Depth 6 | Set-Content -Path $obsPath -Encoding utf8

Write-Host "`nRegistre     : $csvPath"
Write-Host "Registre     : $jsonPath"
Write-Host "Observations : $obsPath"
$registre | Format-Table cas, nom, declare, accorde, state_attendu, evaluation_status_attendu -AutoSize

Write-Host "`nNettoyage : Get-MgApplication -Filter `"startswith(displayName,'$Prefix')`""
Write-Host "            Get-MgServicePrincipal -Filter `"startswith(displayName,'$Prefix')`""
