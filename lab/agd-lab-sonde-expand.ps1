#requires -Version 7.0
<#
    agd-lab-sonde-expand.ps1

    LECTURE SEULE. Ne cree, ne modifie et ne supprime rien.

    Deux questions, mesurees et non supposees :

    1. Combien coute l'axe /servicePrincipals ? Un appel par principal, ou un
       seul appel pagine avec $expand=appRoleAssignments ?

    2. Si $expand fonctionne, rend-il la meme chose que la lecture individuelle ?
       Les collections developpees sont souvent plafonnees. Une troncature
       silencieuse ferait sous-declarer les attributions par app-grant-diff,
       sans erreur, sur la population meme qu'il doit rendre visible.

    La reference est la lecture individuelle, principal par principal. $expand
    est l'hypothese a valider contre elle, jamais l'inverse.

    LIMITE CONNUE : ce tenant ne contient probablement aucun principal portant
    plus de quelques attributions. La sonde detecte donc une divergence de
    presence, pas un seuil de troncature. Tester le seuil demanderait un
    principal portant vingt-cinq attributions ou plus, ce qui suppose des
    ecritures et sort du perimetre de ce script.
#>

[CmdletBinding()]
param(
    [string] $OutputDir
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

if (-not $OutputDir) {
    $base = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
    $OutputDir = Join-Path $base 'agd-lab-out'
}
if (-not (Test-Path -LiteralPath $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }
Write-Host "Sortie  : $OutputDir"

# --- Garde-fou modules ------------------------------------------------------

$loaded = @(Get-Module Microsoft.Graph.Authentication)
if ($loaded.Count -gt 1) { throw "Plusieurs versions de Microsoft.Graph.Authentication chargees. Processus neuf requis." }
$authAvailable = @(Get-Module -ListAvailable Microsoft.Graph.Authentication | Sort-Object Version -Descending)
$appsAvailable = @(Get-Module -ListAvailable Microsoft.Graph.Applications   | Sort-Object Version -Descending)
if ($authAvailable.Count -eq 0 -or $appsAvailable.Count -eq 0) { throw "Modules Graph absents du disque." }
$appsVersions = @($appsAvailable.Version)
$common = @($authAvailable.Version | Where-Object { $_ -in $appsVersions } | Sort-Object -Descending) |
          Select-Object -First 1
if (-not $common) { throw "Aucune version commune entre Authentication et Applications." }
if ($loaded.Count -eq 1 -and $loaded[0].Version -ne $common) {
    throw "Microsoft.Graph.Authentication $($loaded[0].Version) deja chargee, cible $common. Processus neuf requis."
}
Import-Module Microsoft.Graph.Authentication -RequiredVersion $common -Force
Import-Module Microsoft.Graph.Applications   -RequiredVersion $common -Force
Write-Host "Modules Graph pinnes en $common"

Connect-MgGraph -Scopes 'Application.Read.All','Directory.Read.All' -ContextScope Process -NoWelcome
Write-Host "Tenant  : $((Get-MgContext).TenantId)"

# --- Observations -----------------------------------------------------------

$obsPath = Join-Path $OutputDir 'agd-lab-observations.json'
$observations = [System.Collections.Generic.List[object]]::new()
if (Test-Path -LiteralPath $obsPath) {
    foreach ($o in @(Get-Content -LiteralPath $obsPath -Raw | ConvertFrom-Json)) { $observations.Add($o) }
}
function Add-Observation {
    param([string]$Cas, [string]$Fait, $Valeur)
    $observations.Add([pscustomobject]@{
        horodatage = (Get-Date).ToUniversalTime().ToString('o')
        cas        = $Cas
        fait       = $Fait
        valeur     = $Valeur
        source     = 'agd-lab-sonde-expand'
    })
}

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

# ---------------------------------------------------------------------------
# A. REFERENCE : LECTURE INDIVIDUELLE
#    C'est la verite contre laquelle $expand sera juge, et la mesure du cout
#    reel de l'axe /servicePrincipals.
# ---------------------------------------------------------------------------

Write-Host "`nA. Lecture individuelle, principal par principal"

$sps = @(Get-MgServicePrincipal -All -Property 'id,appId,displayName' | Sort-Object DisplayName)
Write-Host "$($sps.Count) principaux"

$reference = @{}
$appels = 0
$echecs = [System.Collections.Generic.List[object]]::new()
$swRef = [System.Diagnostics.Stopwatch]::StartNew()

for ($i = 0; $i -lt $sps.Count; $i++) {
    $sp = $sps[$i]
    if ($i % 25 -eq 0) { Write-Host "  $i / $($sps.Count)" }
    try {
        $g = @(Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -All)
        $appels++
        $reference[$sp.Id] = @($g | ForEach-Object { $_.Id })
    } catch {
        $appels++
        $echecs.Add([pscustomobject]@{ id = $sp.Id; nom = $sp.DisplayName; erreur = $_.Exception.Message })
        $reference[$sp.Id] = $null   # non observable, distinct de zero
    }
}
$swRef.Stop()

$avecGrants = @($reference.Keys | Where-Object { $null -ne $reference[$_] -and $reference[$_].Count -gt 0 })
$nonObservables = @($reference.Keys | Where-Object { $null -eq $reference[$_] })

Add-Observation 'sonde' 'reference_principaux'        $sps.Count
Add-Observation 'sonde' 'reference_appels'            $appels
Add-Observation 'sonde' 'reference_duree_ms'          $swRef.ElapsedMilliseconds
Add-Observation 'sonde' 'reference_avec_grants'       $avecGrants.Count
Add-Observation 'sonde' 'reference_non_observables'   $nonObservables.Count
foreach ($e in $echecs) { Add-Observation 'sonde' 'reference_echec_lecture' $e }

Write-Host "$appels appels, $([math]::Round($swRef.Elapsed.TotalSeconds,1)) s, $($avecGrants.Count) principaux porteurs d'attributions"
if ($nonObservables.Count -gt 0) {
    Write-Warning "$($nonObservables.Count) principal(aux) non observable(s). Ce n'est pas zero attribution, c'est une absence de mesure."
}

# ---------------------------------------------------------------------------
# B. HYPOTHESE : $expand
#    Trois formes, parce qu'aucune documentation consultee ne confirme que le
#    point de terminaison de liste accepte cette relation. On essaie, on note
#    ce qui repond, on ne suppose rien.
# ---------------------------------------------------------------------------

Write-Host "`nB. Formes d'appel avec `$expand"

$formes = @(
    [pscustomobject]@{ nom = 'expand_seul'
        uri = 'https://graph.microsoft.com/v1.0/servicePrincipals?$expand=appRoleAssignments' }
    [pscustomobject]@{ nom = 'select_puis_expand'
        uri = 'https://graph.microsoft.com/v1.0/servicePrincipals?$select=id,appId,displayName&$expand=appRoleAssignments' }
    [pscustomobject]@{ nom = 'expand_avec_select_imbrique'
        uri = 'https://graph.microsoft.com/v1.0/servicePrincipals?$expand=appRoleAssignments($select=id,appRoleId,resourceId,principalId)' }
)

$formeRetenue = $null
$expand = @{}

foreach ($f in $formes) {
    Write-Host "  $($f.nom)"
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $pages = 0
    $collecte = @{}
    $ok = $true
    $erreur = ''
    try {
        $next = $f.uri
        while ($next) {
            $page = Invoke-MgGraphRequest -Method GET -Uri $next -OutputType PSObject
            $pages++
            foreach ($v in @(Get-Champ -Item $page -Nom 'value')) {
                $id = Get-Champ -Item $v -Nom 'id'
                $ass = Get-Champ -Item $v -Nom 'appRoleAssignments'
                if ($null -eq $ass) {
                    $collecte[$id] = $null      # relation non rendue : non observable
                } else {
                    $collecte[$id] = @(@($ass) | ForEach-Object { Get-Champ -Item $_ -Nom 'id' })
                }
            }
            $next = Get-Champ -Item $page -Nom '@odata.nextLink'
        }
    } catch {
        $ok = $false
        $erreur = $_.Exception.Message
    }
    $sw.Stop()

    $rendus = if ($ok) { @($collecte.Keys | Where-Object { $null -ne $collecte[$_] }).Count } else { 0 }

    Add-Observation 'sonde' "expand_$($f.nom)" ([pscustomobject]@{
        uri                          = $f.uri
        accepte                      = $ok
        erreur                       = $erreur
        pages                        = $pages
        duree_ms                     = $sw.ElapsedMilliseconds
        principaux_rendus            = $collecte.Count
        relation_effectivement_rendue = $rendus
    })

    if ($ok) {
        Write-Host "    accepte, $pages page(s), $([math]::Round($sw.Elapsed.TotalSeconds,1)) s, relation rendue pour $rendus / $($collecte.Count)"
        # Une forme acceptee qui ne rend la relation pour personne n'est pas
        # une forme qui marche : Graph a ignore le $expand sans le dire.
        if ($rendus -gt 0 -and -not $formeRetenue) { $formeRetenue = $f.nom; $expand = $collecte }
    } else {
        Write-Host "    refuse : $erreur"
    }
}

Add-Observation 'sonde' 'expand_forme_retenue' $(if ($formeRetenue) { $formeRetenue } else { 'aucune' })

# ---------------------------------------------------------------------------
# C. CONFRONTATION
# ---------------------------------------------------------------------------

if (-not $formeRetenue) {
    Write-Warning "`nAucune forme `$expand exploitable. L'axe /servicePrincipals coute un appel par principal."
    Add-Observation 'sonde' 'conclusion' 'expand_inexploitable_cout_lineaire'
} else {
    Write-Host "`nC. Confrontation de '$formeRetenue' a la lecture individuelle"

    $lignes = foreach ($sp in $sps) {
        $ref = $reference[$sp.Id]
        $exp = if ($expand.ContainsKey($sp.Id)) { $expand[$sp.Id] } else { $null }

        $verdict =
            if ($null -eq $ref)                       { 'reference_non_observable' }
            elseif ($null -eq $exp)                   { 'expand_non_observable' }
            elseif (@($ref).Count -ne @($exp).Count)  { 'ECART_COMPTE' }
            elseif (@(@($ref) | Where-Object { $_ -notin @($exp) }).Count -gt 0) { 'ECART_CONTENU' }
            else                                      { 'concordant' }

        [pscustomobject]@{
            nom              = $sp.DisplayName
            sp_object_id     = $sp.Id
            app_id           = $sp.AppId
            reference_compte = if ($null -eq $ref) { '' } else { @($ref).Count }
            expand_compte    = if ($null -eq $exp) { '' } else { @($exp).Count }
            verdict          = $verdict
        }
    }
    $lignes = @($lignes)

    $ecarts = @($lignes | Where-Object { $_.verdict -like 'ECART*' })
    $absents = @($lignes | Where-Object { $_.verdict -eq 'expand_non_observable' })

    Add-Observation 'sonde' 'confrontation_total'   $lignes.Count
    Add-Observation 'sonde' 'confrontation_ecarts'  $ecarts.Count
    Add-Observation 'sonde' 'confrontation_absents' $absents.Count
    foreach ($e in $ecarts) { Add-Observation 'sonde' 'ecart_expand' $e }

    $lignes | Where-Object { $_.verdict -ne 'concordant' -or $_.reference_compte -ne 0 } |
        Format-Table nom, reference_compte, expand_compte, verdict -AutoSize

    if ($ecarts.Count -eq 0 -and $absents.Count -eq 0) {
        Add-Observation 'sonde' 'conclusion' 'expand_concordant_sur_ce_tenant'
        Write-Host "Concordance complete sur ce tenant."
        Write-Host "Portee : aucun principal n'y porte assez d'attributions pour eprouver un plafond."
    } else {
        Add-Observation 'sonde' 'conclusion' 'expand_divergent'
        Write-Warning "$($ecarts.Count) ecart(s), $($absents.Count) principal(aux) absent(s) du resultat `$expand."
    }

    $confPath = Join-Path $OutputDir 'agd-lab-sonde-expand.csv'
    $lignes | Export-Csv -LiteralPath $confPath -NoTypeInformation -Encoding utf8
    Write-Host "Confrontation : $confPath"
}

$observations | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $obsPath -Encoding utf8
Write-Host "Observations  : $obsPath"
