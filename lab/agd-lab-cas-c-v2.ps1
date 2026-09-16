#requires -Version 7.0
<#
    agd-lab-cas-c-v2.ps1

    Remplace agd-lab-cas-c.ps1 et le bloc "Cas C" de agd-lab-fabrique.ps1,
    qui reposait sur une survie du principal apres suppression de l'application.
    Ce comportement n'existe pas : Learn documente que supprimer l'inscription
    dans son tenant d'origine supprime aussi le principal correspondant.

    ETAPE 1  Recuperer les objets du premier passage dans directory/deletedItems
             et prouver la cascade. Le script de fabrication a conclu sans
             enregistrer l'identifiant de l'objet dont il parlait ; on le
             retrouve par displayName tant que la retention de 30 jours court.

    ETAPE 2  Lister les principaux sans objet application local, dans un CSV
             et non dans une table console tronquee.

    ETAPE 3  Instancier le cas C sans conferer aucun privilege reel : une
             application de lab expose son propre role applicatif, et c'est ce
             role qui est attribue au principal choisi. Aucun service
             n'interprete ce role.

    Sans -TargetSpId ni -TargetAppId, les etapes 1 et 2 s'executent et le
    script s'arrete sans rien ecrire dans le tenant.
#>

[CmdletBinding()]
param(
    [string] $OutputDir,
    [string] $Prefix     = 'agd-lab',
    [string] $TargetSpId,
    [string] $TargetAppId,
    [int]    $MaxRetries = 12,
    [int]    $RetryDelay = 5
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

if (-not $OutputDir) {
    $base = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
    $OutputDir = Join-Path $base 'agd-lab-out'
}
if (-not (Test-Path -LiteralPath $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}
Write-Host "Sortie  : $OutputDir"

# --- Garde-fou modules ------------------------------------------------------

$loaded = @(Get-Module Microsoft.Graph.Authentication)
if ($loaded.Count -gt 1) {
    throw "Plusieurs versions de Microsoft.Graph.Authentication chargees : $(($loaded.Version) -join ', '). Relancer dans un processus neuf."
}
$authAvailable = @(Get-Module -ListAvailable Microsoft.Graph.Authentication | Sort-Object Version -Descending)
$appsAvailable = @(Get-Module -ListAvailable Microsoft.Graph.Applications   | Sort-Object Version -Descending)
if ($authAvailable.Count -eq 0) { throw "Microsoft.Graph.Authentication absent du disque." }
if ($appsAvailable.Count -eq 0) { throw "Microsoft.Graph.Applications absent du disque." }
$appsVersions = @($appsAvailable.Version)
$common = @($authAvailable.Version | Where-Object { $_ -in $appsVersions } | Sort-Object -Descending) |
          Select-Object -First 1
if (-not $common) { throw "Aucune version commune entre Authentication et Applications." }
if ($loaded.Count -eq 1 -and $loaded[0].Version -ne $common) {
    throw "Microsoft.Graph.Authentication $($loaded[0].Version) deja chargee, cible $common. Relancer dans un processus neuf."
}
Import-Module Microsoft.Graph.Authentication -RequiredVersion $common -Force
Import-Module Microsoft.Graph.Applications   -RequiredVersion $common -Force
Write-Host "Modules Graph pinnes en $common"

Connect-MgGraph -Scopes 'Application.ReadWrite.All','AppRoleAssignment.ReadWrite.All','Directory.Read.All' `
                -ContextScope Process -NoWelcome
$ctx = Get-MgContext
$tenantId = $ctx.TenantId
Write-Host "Tenant  : $tenantId"

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
    })
}
function Save-Observations {
    $observations | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $obsPath -Encoding utf8
}

function Get-Champ {
    # Invoke-MgGraphRequest peut rendre des tables de hachage ou des PSObject
    # selon la version et l'OutputType. L'acces par point marche sur les deux,
    # Select-Object ne marche que sur le second. On lit donc explicitement au
    # lieu de supposer la forme.
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

function ConvertTo-ObjetPlat {
    # type est une cle de l'objet des sa construction, pas un membre ajoute
    # apres coup : un objet mal forme se voit a la lecture, pas au premier
    # Where-Object qui trebuche dessus.
    param($Item, [string[]]$Champs, [string]$Type = '')
    $h = [ordered]@{ type = $Type }
    foreach ($c in $Champs) { $h[$c] = Get-Champ -Item $Item -Nom $c }
    [pscustomobject]$h
}

function Get-GraphCollection {
    # Pagination explicite. Le filtre serveur sur deletedItems est partiel,
    # Learn recommande de filtrer cote client sur cette collection.
    param(
        [string]   $Uri,
        [string]   $Etiquette = '',
        [string[]] $Champs = @('id','displayName','appId','deletedDateTime')
    )
    $items = [System.Collections.Generic.List[object]]::new()
    $formeBrute = $null
    $next = $Uri
    while ($next) {
        $page = Invoke-MgGraphRequest -Method GET -Uri $next -OutputType PSObject
        $valeurs = @(Get-Champ -Item $page -Nom 'value')
        foreach ($v in $valeurs) {
            if (-not $formeBrute) { $formeBrute = $v.GetType().FullName }
            $items.Add((ConvertTo-ObjetPlat -Item $v -Champs $Champs -Type $Etiquette))
        }
        $next = Get-Champ -Item $page -Nom '@odata.nextLink'
    }
    if ($formeBrute) { Add-Observation 'C' 'forme_brute_elements_graph' $formeBrute }
    # Pas de virgule devant le retour. Avec ,$tableau le site d'appel qui fait
    # @(...) recoit un tableau imbrique : un seul element, qui est le tableau.
    # Where-Object voit alors le tableau et non les objets, $_.propriete
    # declenche l'enumeration de membres, et -like sur un tableau rend les
    # elements correspondants au lieu d'un booleen, donc filtre toujours vrai.
    $items.ToArray()
}

function Test-Forme {
    # Controle d'integrite explicite. Sous StrictMode, une propriete manquante
    # fait echouer la premiere expression qui la touche, et le message designe
    # cette expression, pas la cause. On verifie donc la forme au moment ou on
    # la produit.
    param($Objets, [string[]]$Attendus, [string]$Etiquette)

    # Detecter l'imbrication avant de parler de proprietes manquantes : un
    # element qui est lui-meme une collection produit exactement le meme
    # symptome, toutes les proprietes absentes, pour une cause differente.
    $imbriques = @(@($Objets) | Where-Object {
        $_ -is [System.Collections.IEnumerable] -and $_ -isnot [string]
    })
    Add-Observation 'C' 'elements_imbriques' ([pscustomobject]@{
        source = $Etiquette; total = @($Objets).Count; imbriques = $imbriques.Count
    })
    if ($imbriques.Count -gt 0) {
        throw "Collection imbriquee dans '$Etiquette' : $($imbriques.Count) element(s) sur $(@($Objets).Count) sont eux-memes des collections."
    }

    $manquants = [System.Collections.Generic.List[string]]::new()
    foreach ($o in @($Objets)) {
        foreach ($a in $Attendus) {
            if (-not $o.PSObject.Properties[$a]) { $manquants.Add($a) }
        }
    }
    $resume = [pscustomobject]@{
        source              = $Etiquette
        objets              = @($Objets).Count
        proprietes_absentes = @($manquants | Sort-Object -Unique) -join ', '
    }
    Add-Observation 'C' 'controle_forme' $resume
    if ($manquants.Count -gt 0) {
        throw "Objets mal formes dans '$Etiquette' : proprietes absentes $($resume.proprietes_absentes)."
    }
}

# ---------------------------------------------------------------------------
# ETAPE 1 : RECUPERER LES OBJETS SUPPRIMES, PROUVER LA CASCADE
# ---------------------------------------------------------------------------

Write-Host "`nEtape 1 : objets supprimes"

# La cascade se prouve par un COUPLE, pas par deux comptes non nuls.
# Une inscription supprimee et un principal supprime partageant le meme appId,
# quand seule l'inscription a ete supprimee explicitement, etablissent la
# cascade. Deux objets sans lien commun ne l'etablissent pas, meme s'ils sont
# tous deux presents. Le controle negatif est ici agd-lab-z-jetable, dont les
# deux objets ont ete supprimes separement : son couple doit exister sans
# jamais servir a conclure.

$delSp  = @()
$delApp = @()
$lectureDelOk = $true
try {
    $delSp  = @(Get-GraphCollection -Uri 'https://graph.microsoft.com/v1.0/directory/deletedItems/microsoft.graph.servicePrincipal' -Etiquette 'servicePrincipal')
    $delApp = @(Get-GraphCollection -Uri 'https://graph.microsoft.com/v1.0/directory/deletedItems/microsoft.graph.application'      -Etiquette 'application')
} catch {
    $lectureDelOk = $false
    Add-Observation 'C' 'lecture_deleteditems' "echec : $($_.Exception.Message)"
    Write-Warning "Lecture de deletedItems impossible : $($_.Exception.Message)"
}

if ($lectureDelOk) {
    Add-Observation 'C' 'deleteditems_sp_total'  $delSp.Count
    Add-Observation 'C' 'deleteditems_app_total' $delApp.Count

    $attendus = @('type','id','displayName','appId','deletedDateTime')
    Test-Forme -Objets $delSp  -Attendus $attendus -Etiquette 'deletedItems/servicePrincipal'
    Test-Forme -Objets $delApp -Attendus $attendus -Etiquette 'deletedItems/application'

    $motifLarge = "$Prefix*"
    $motifCasC  = "$Prefix-c-*"
    Add-Observation 'C' 'motif_filtre_large' $motifLarge
    Add-Observation 'C' 'motif_filtre_cas_c' $motifCasC

    $tous = @($delApp + $delSp | Where-Object { $_.displayName -and ($_.displayName -like $motifLarge) })
    Add-Observation 'C' 'objets_supprimes_du_lab' $tous.Count

    # Journaliser la decision du filtre objet par objet. Si un nom qui ne
    # devrait pas correspondre correspond, ce fichier le montrera au lieu de
    # laisser la question ouverte.
    foreach ($o in $tous) {
        Add-Observation 'C' 'objet_supprime' ([pscustomobject]@{
            type            = $o.type
            displayName     = $o.displayName
            id              = $o.id
            appId           = $o.appId
            deletedDateTime = $o.deletedDateTime
            correspond_cas_c = ($o.displayName -like $motifCasC)
        })
    }

    # Appariement par appId : c'est la cle qui lie une inscription a son principal.
    $couples = foreach ($g in ($tous | Where-Object { $_.appId } | Group-Object appId)) {
        $apps = @($g.Group | Where-Object { $_.type -eq 'application' })
        $sps  = @($g.Group | Where-Object { $_.type -eq 'servicePrincipal' })
        [pscustomobject]@{
            appId            = $g.Name
            nom              = @($g.Group.displayName | Select-Object -Unique) -join ' / '
            application      = $apps.Count
            servicePrincipal = $sps.Count
            couple_complet   = ($apps.Count -ge 1 -and $sps.Count -ge 1)
            correspond_cas_c = [bool]@($g.Group | Where-Object { $_.displayName -like $motifCasC }).Count
            supprime_le      = @($g.Group.deletedDateTime | Select-Object -Unique) -join ' / '
        }
    }
    $couples = @($couples)

    foreach ($c in $couples) { Add-Observation 'C' 'couple_deleteditems' $c }
    $couples | Format-Table nom, appId, application, servicePrincipal, couple_complet, correspond_cas_c -AutoSize

    # Conclusion : uniquement sur le couple du cas C.
    $coupleC = @($couples | Where-Object { $_.correspond_cas_c })

    if ($coupleC.Count -eq 0) {
        Add-Observation 'C' 'cascade_application_vers_sp' 'non_conclue'
        Add-Observation 'C' 'cascade_statut_evaluation'   'aucun_couple_cas_c'
        Write-Warning "Aucun couple '$motifCasC' dans deletedItems. Retention depassee, ou prefixe different."
    } elseif ($coupleC.Count -gt 1) {
        Add-Observation 'C' 'cascade_application_vers_sp' 'non_conclue'
        Add-Observation 'C' 'cascade_statut_evaluation'   'plusieurs_couples_cas_c'
        Write-Warning "$($coupleC.Count) couples '$motifCasC'. Plusieurs passages du script de fabrication : impossible de designer lequel sans ambiguite."
    } elseif ($coupleC[0].couple_complet) {
        Add-Observation 'C' 'cascade_application_vers_sp' 'confirmee'
        Add-Observation 'C' 'cascade_statut_evaluation'   'verifie_en_tenant_avec_artefact'
        Add-Observation 'C' 'cascade_couple_probant'      $coupleC[0]
        Write-Host "Cascade confirmee sur le couple appId $($coupleC[0].appId) :"
        Write-Host "l'inscription seule a ete supprimee, le principal est dans deletedItems."
    } else {
        Add-Observation 'C' 'cascade_application_vers_sp' 'infirmee'
        Add-Observation 'C' 'cascade_statut_evaluation'   'evalue'
        Write-Warning "Couple incomplet : application=$($coupleC[0].application) principal=$($coupleC[0].servicePrincipal). La cascade n'a pas eu lieu."
    }

    # Controle negatif explicite. Sa presence ne doit jamais servir a conclure.
    $coupleZ = @($couples | Where-Object { $_.nom -like "*z-jetable*" })
    if ($coupleZ.Count -gt 0) {
        Add-Observation 'C' 'controle_negatif_z' ([pscustomobject]@{
            couple_complet = $coupleZ[0].couple_complet
            note           = 'Les deux objets ont ete supprimes explicitement. Ce couple ne prouve aucune cascade.'
        })
    }
}

# ---------------------------------------------------------------------------
# ETAPE 2 : CANDIDATS
# ---------------------------------------------------------------------------

Write-Host "`nEtape 2 : principaux sans objet application local"

$localAppIds = @(Get-MgApplication -All -Property 'id,appId' | Select-Object -ExpandProperty AppId)
$allSp = @(Get-MgServicePrincipal -All -Property 'id,appId,displayName,servicePrincipalType,appOwnerOrganizationId')

$candidats = @($allSp | Where-Object {
    $_.AppOwnerOrganizationId -and
    $_.AppOwnerOrganizationId -ne $tenantId -and
    $_.AppId -notin $localAppIds -and
    $_.ServicePrincipalType -eq 'Application'
} | ForEach-Object {
    [pscustomobject]@{
        DisplayName            = $_.DisplayName
        SpObjectId             = $_.Id
        AppId                  = $_.AppId
        AppOwnerOrganizationId = $_.AppOwnerOrganizationId
    }
} | Sort-Object DisplayName)

$candPath = Join-Path $OutputDir 'agd-lab-candidats-cas-c.csv'
$candidats | Export-Csv -LiteralPath $candPath -NoTypeInformation -Encoding utf8

Add-Observation 'C' 'candidats_total' $candidats.Count
Write-Host "$($candidats.Count) candidats. Liste complete : $candPath"
Write-Host "La table console tronquait les identifiants, ils sont dans le CSV."

if ($candidats.Count -eq 0) {
    Save-Observations
    throw "Aucun principal sans objet application local. Cas C non instanciable ici."
}

# --- Resolution de la cible -------------------------------------------------

$cible = $null
if ($TargetSpId)  { $cible = @($candidats | Where-Object { $_.SpObjectId -eq $TargetSpId }) }
if ($TargetAppId -and -not $cible) { $cible = @($candidats | Where-Object { $_.AppId -eq $TargetAppId }) }

if (-not $cible) {
    if ($TargetSpId -or $TargetAppId) {
        Save-Observations
        throw "Cible fournie absente de la liste des candidats."
    }
    Write-Host "`nAucune cible fournie. Rien n'a ete ecrit dans le tenant."
    Write-Host "Relancer avec -TargetAppId <appId> pour une cible stable d'un tenant a l'autre,"
    Write-Host "ou -TargetSpId <objectId> pour un objet de ce tenant."
    Save-Observations
    Write-Host "Observations : $obsPath"
    return
}

$cible = @($cible)[0]
Write-Host "`nCible : $($cible.DisplayName)  sp=$($cible.SpObjectId)  app=$($cible.AppId)"

# ---------------------------------------------------------------------------
# ETAPE 3 : RESSOURCE DE LAB ET ATTRIBUTION
#    Le role est expose par une application locale et n'est interprete par
#    aucun service. L'attribution est structurellement identique a une
#    permission applicative, elle ne confere aucun acces.
# ---------------------------------------------------------------------------

Write-Host "`nEtape 3 : ressource de lab"

$resName = "$Prefix-c-ressource"
$resApp = @(Get-MgApplication -Filter "displayName eq '$resName'" -Property 'id,appId,displayName,appRoles')

if ($resApp.Count -eq 0) {
    $roleId = [guid]::NewGuid().ToString()
    $resApp = @(New-MgApplication -BodyParameter @{
        displayName    = $resName
        signInAudience = 'AzureADMyOrg'
        appRoles       = @(@{
            allowedMemberTypes = @('Application')
            description        = 'Role de laboratoire sans effet. Sert a instancier le cas C de app-grant-diff.'
            displayName        = 'AgdLab Probe'
            id                 = $roleId
            isEnabled          = $true
            value              = 'AgdLab.Probe'
        })
    })
    Write-Host "Application ressource creee."
} else {
    $existing = @($resApp[0].AppRoles | Where-Object { $_.Value -eq 'AgdLab.Probe' })
    if ($existing.Count -eq 0) { throw "L'application $resName existe sans le role AgdLab.Probe." }
    $roleId = $existing[0].Id
    Write-Host "Application ressource deja presente."
}

$resAppObj = $resApp[0]
Add-Observation 'C' 'ressource_app_object_id' $resAppObj.Id
Add-Observation 'C' 'ressource_app_id'        $resAppObj.AppId
Add-Observation 'C' 'ressource_role_id'       $roleId

$resSp = @(Get-MgServicePrincipal -Filter "appId eq '$($resAppObj.AppId)'" -Property 'id,appId,appRoles')
if ($resSp.Count -eq 0) {
    $resSp = @(New-MgServicePrincipal -AppId $resAppObj.AppId)
}
$resSpId = $resSp[0].Id
Add-Observation 'C' 'ressource_sp_object_id' $resSpId

# Le role est declare sur l'application, il doit apparaitre sur le principal
# avant d'etre attribuable. La propagation n'est pas instantanee : on attend
# de l'observer plutot que de le supposer.
$sw = [System.Diagnostics.Stopwatch]::StartNew()
$rolePresent = $false
for ($i = 0; $i -lt $MaxRetries; $i++) {
    $sp = Get-MgServicePrincipal -ServicePrincipalId $resSpId -Property 'id,appRoles'
    if (@($sp.AppRoles | Where-Object { $_.Id -eq $roleId }).Count -eq 1) { $rolePresent = $true; break }
    Start-Sleep -Seconds $RetryDelay
}
$sw.Stop()
Add-Observation 'C' 'propagation_role_application_vers_sp_ms' $sw.ElapsedMilliseconds
Add-Observation 'C' 'propagation_role_observee'               $rolePresent

if (-not $rolePresent) {
    Save-Observations
    throw "Le role $roleId n'est pas apparu sur le principal ressource apres $($sw.Elapsed). Rien n'a ete attribue."
}

# Forme recommandee par Learn : appRoleAssignedTo sur le principal RESSOURCE.
$avant = @(Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $cible.SpObjectId -All)
$assign = New-MgServicePrincipalAppRoleAssignedTo -ServicePrincipalId $resSpId -BodyParameter @{
    principalId = $cible.SpObjectId
    resourceId  = $resSpId
    appRoleId   = $roleId
}
Add-Observation 'C' 'assignment_id'      $assign.Id
Add-Observation 'C' 'forme_creation'     'appRoleAssignedTo sur le principal ressource'

$attendu = $avant.Count + 1
$sw = [System.Diagnostics.Stopwatch]::StartNew()
$vus = @()
for ($i = 0; $i -lt $MaxRetries; $i++) {
    $vus = @(Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $cible.SpObjectId -All)
    if ($vus.Count -eq $attendu) { break }
    Start-Sleep -Seconds $RetryDelay
}
$sw.Stop()
Add-Observation 'C' 'grants_avant'       $avant.Count
Add-Observation 'C' 'grants_attendus'    $attendu
Add-Observation 'C' 'grants_observes'    $vus.Count
Add-Observation 'C' 'latence_lecture_ms' $sw.ElapsedMilliseconds

if ($vus.Count -ne $attendu) {
    Write-Warning "Lecture non convergente : attendu $attendu, observe $($vus.Count)."
    Add-Observation 'C' 'statut_evaluation' 'non_convergent'
} else {
    Add-Observation 'C' 'statut_evaluation' 'evalue'
}

# Confirmer l'absence d'objet application local pour la cible, au lieu de la deduire.
$appLocale = @(Get-MgApplication -Filter "appId eq '$($cible.AppId)'" -Property 'id,appId')
Add-Observation 'C' 'objet_application_local_de_la_cible' $appLocale.Count
if ($appLocale.Count -ne 0) {
    Write-Warning "La cible possede un objet application local. Le cas C n'est pas ce qu'il pretend etre."
}

# --- Registre ---------------------------------------------------------------

$csvPath  = Join-Path $OutputDir 'agd-lab-registre.csv'
$jsonPath = Join-Path $OutputDir 'agd-lab-registre.json'

$registre = [System.Collections.Generic.List[object]]::new()
if (Test-Path -LiteralPath $csvPath) {
    foreach ($r in @(Import-Csv -LiteralPath $csvPath)) { if ($r.cas -ne 'C') { $registre.Add($r) } }
}
$registre.Add([pscustomobject]@{
    cas                       = 'C'
    nom                       = $cible.DisplayName
    app_id                    = $cible.AppId
    app_object_id             = '(absent du tenant)'
    sp_object_id              = $cible.SpObjectId
    declare                   = '(inexistant dans le tenant)'
    accorde                   = "AgdLab.Probe ($resName)"
    state_attendu             = 'not_comparable'
    evaluation_status_attendu = 'declared_state_unavailable'
})

$ordonne = $registre | Sort-Object cas
$ordonne | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding utf8
$ordonne | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $jsonPath -Encoding utf8
Save-Observations

$ordonne | Format-Table cas, nom, declare, accorde, state_attendu -AutoSize

Write-Host "`nRevocation :"
Write-Host "Remove-MgServicePrincipalAppRoleAssignedTo -ServicePrincipalId '$resSpId' -AppRoleAssignmentId '$($assign.Id)'"
