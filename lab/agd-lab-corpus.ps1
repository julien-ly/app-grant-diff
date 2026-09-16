#requires -Version 7.0
<#
    agd-lab-corpus.ps1

    Une seule écriture : l'inscription du cas F, si elle n'existe pas déjà.
    Tout le reste est de la lecture et de la génération de fichiers.

    Produit :
      agd-lab-intent.json    manifeste d'intention, appId résolus depuis le tenant
      agd-lab-registre.csv   registre à deux colonnes, sans manifeste et avec

    SÉPARATION QUI COMPTE.
    Les identifiants sont résolus depuis le tenant, parce que ce sont des données
    de recherche. Les attentes, elles, sont écrites en dur dans ce script, parce
    que les dériver de l'état observé rendrait le test auto-confirmant : le
    manifeste vaudrait exactement ce qu'il est censé vérifier, et l'outil
    conclurait NoGapEstablished par construction.

    CAS F.
    Couvre CorrectExclusion, seul état du vocabulaire qu'aucun autre objet ne
    produit. Une inscription qui ne déclare rien, aucun grant, et une entrée de
    manifeste posant Directory.Read.All explicitement non attendu.
    Sans manifeste il ne produit aucune ligne, avec manifeste il en produit une.
#>

[CmdletBinding()]
param(
    [string] $OutputDir,
    [string] $Prefix = 'agd-lab',
    [switch] $SansEcriture
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$GRAPH_APPID = '00000003-0000-0000-c000-000000000000'
$EXO_APPID   = '00000002-0000-0ff1-ce00-000000000000'

if (-not $OutputDir) {
    $base = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
    $OutputDir = Join-Path $base 'agd-lab-out'
}
if (-not (Test-Path -LiteralPath $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }
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

$scopes = if ($SansEcriture) { @('Application.Read.All','Directory.Read.All') }
          else               { @('Application.ReadWrite.All','Directory.Read.All') }
Connect-MgGraph -Scopes $scopes -ContextScope Process -NoWelcome
Write-Host "Tenant  : $((Get-MgContext).TenantId)"

# ── 1. Cas F ────────────────────────────────────────────────────────────────

$nomF = "$Prefix-f-exclusion-attendue"
$appF = @(Get-MgApplication -Filter "displayName eq '$nomF'" -Property 'id,appId,displayName')

if ($appF.Count -eq 0) {
    if ($SansEcriture) {
        Write-Warning "Le cas F n'existe pas et -SansEcriture est actif. Il sera absent du corpus."
    } else {
        Write-Host "`nCreation du cas F"
        $appF = @(New-MgApplication -BodyParameter @{
            displayName            = $nomF
            signInAudience         = 'AzureADMyOrg'
            requiredResourceAccess = @()
        })
        $null = New-MgServicePrincipal -AppId $appF[0].AppId
        Write-Host "  $nomF cree, appId $($appF[0].AppId)"
    }
} elseif ($appF.Count -gt 1) {
    throw "Plusieurs inscriptions nommees '$nomF'. Corpus ambigu, rien n'a ete ecrit."
} else {
    Write-Host "`nCas F deja present, appId $($appF[0].AppId)"
}

# ── 2. Résolution des identifiants ──────────────────────────────────────────

Write-Host "`nResolution des identifiants"

$appsLab = @(Get-MgApplication -All -Property 'id,appId,displayName' |
             Where-Object { $_.DisplayName -and $_.DisplayName.StartsWith($Prefix) } |
             Sort-Object DisplayName)

$doublons = @($appsLab | Group-Object DisplayName | Where-Object { $_.Count -gt 1 })
if ($doublons.Count -gt 0) {
    throw "Inscriptions en double : $(@($doublons.Name) -join ', '). Le registre serait ambigu, rien n'est ecrit."
}

$appIdParNom = @{}
foreach ($a in $appsLab) { $appIdParNom[$a.DisplayName] = $a.AppId }

$exo = @(Get-MgServicePrincipal -Filter "appId eq '$EXO_APPID'" -Property 'id,appId,displayName')
if ($exo.Count -ne 1) { throw "Principal Office 365 Exchange Online introuvable, le cas C n'est pas resoluble." }
$nomC = $exo[0].DisplayName
Write-Host "  $($appsLab.Count) inscription(s) de lab, plus $nomC pour le cas C"

# ── 3. Attentes, écrites à la main ──────────────────────────────────────────
#    Ne jamais dériver ces valeurs du tenant.

$cas = @(
    [pscustomobject]@{ Cas='A'; Nom="$Prefix-a-declare-non-accorde"; Sans='UnderCoverage';   Avec='UnderCoverage';   Partiel='UnderCoverage';    NonResolu='UnderCoverage'
        Commentaire='Inscription locale. Le cote attendu vient de requiredResourceAccess : User.Read.All declare, non consenti.'
        Permissions=$null }
    [pscustomobject]@{ Cas='B'; Nom="$Prefix-b-accorde-non-declare"; Sans='Observed';        Avec='OverCoverage';    Partiel='NotInManifest';    NonResolu='NotInManifest'
        Commentaire='Inscription locale, manifeste Entra vide. complete est ce qui transforme le grant orphelin en OverCoverage : c est la seule difference entre les deux colonnes.'
        Permissions=$null }
    [pscustomobject]@{ Cas='C'; Nom=$nomC;                           Sans='Observed';        Avec='OverCoverage';    Partiel='OverCoverage';     NonResolu='NotInManifest'
        Commentaire='Aucune inscription locale. permissions est ici le seul etat declare qui existe.'
        Permissions=@() }
    [pscustomobject]@{ Cas='D'; Nom="$Prefix-d-aligne";              Sans='CorrectCoverage'; Avec='CorrectCoverage'; Partiel='CorrectCoverage';  NonResolu='CorrectCoverage'
        Commentaire='Declare et accorde identiques.'
        Permissions=$null }
    [pscustomobject]@{ Cas='E'; Nom="$Prefix-e-controle-negatif";    Sans='(aucune ligne)';  Avec='(aucune ligne)';  Partiel='(aucune ligne)';   NonResolu='(aucune ligne)'
        Commentaire='Controle negatif. Son assertion porte sur conclusion.result et assessment.completeness, pas sur un etat de ligne.'
        Permissions=$null }
    [pscustomobject]@{ Cas='F'; Nom=$nomF;                           Sans='(aucune ligne)';  Avec='CorrectExclusion'; Partiel='CorrectExclusion'; NonResolu='CorrectExclusion'
        Commentaire='Seul objet produisant CorrectExclusion. requiredResourceAccess ne peut pas exprimer une attente negative, le manifeste si.'
        Permissions=@(@{ resourceAppId=$GRAPH_APPID; value='Directory.Read.All'; expectedGranted=$false }) }
    [pscustomobject]@{ Cas='G'; Nom="$Prefix-g-troncature";          Sans='Observed';        Avec='OverCoverage';    Partiel='NotInManifest';    NonResolu='NotInManifest'
        Commentaire='25 attributions. Eprouve la limite documentee de 20 elements de $expand. Sans relecture individuelle, cinq lignes manquent sans signal.'
        Permissions=$null }
)

# ── 4. Manifeste ────────────────────────────────────────────────────────────

$principals = [System.Collections.Generic.List[object]]::new()
$absents = [System.Collections.Generic.List[string]]::new()

foreach ($c in $cas) {
    $appId = if ($c.Nom -eq $nomC) { $EXO_APPID }
             elseif ($appIdParNom.ContainsKey($c.Nom)) { $appIdParNom[$c.Nom] }
             else { $null }

    if (-not $appId) { $absents.Add("$($c.Cas) / $($c.Nom)"); continue }

    $entree = [ordered]@{
        appId       = $appId
        displayName = $c.Nom
        complete    = $true
        comment     = $c.Commentaire
    }
    if ($null -ne $c.Permissions) { $entree['permissions'] = @($c.Permissions) }
    $principals.Add([pscustomobject]$entree)
}

# Les objets de lab qui ne sont pas des cas : ressources, supports. Ils doivent
# figurer pour que complete au niveau du manifeste reste vrai.
$nomsCas = @($cas.Nom)
foreach ($a in $appsLab) {
    if ($a.DisplayName -in $nomsCas) { continue }
    $principals.Add([pscustomobject][ordered]@{
        appId       = $a.AppId
        displayName = $a.DisplayName
        complete    = $true
        comment     = 'Objet de support du lab, jamais client. Present pour que le manifeste reste exhaustif.'
        permissions = @()
    })
}

if ($absents.Count -gt 0) {
    Write-Warning "Cas absents du tenant, exclus du manifeste : $(@($absents) -join ' ; ')"
}

$manifeste = [ordered]@{
    description = 'Intention pour le corpus de test app-grant-diff. Attentes ecrites a la main, identifiants resolus depuis le tenant.'
    complete    = $true
    asOf        = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    principals  = $principals.ToArray()
}

$intentPath = Join-Path $OutputDir 'agd-lab-intent.json'
$manifeste | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $intentPath -Encoding utf8

# ── 4bis. Variante partielle ────────────────────────────────────────────────
#    NotInManifest n'est produit par aucun des deux premiers passages, alors que
#    c'est l'etat qu'un tenant reel produira le plus souvent : un grant existe et
#    la reference n'est pas exhaustive ici. Deux chemins distincts y menent dans
#    le moteur, et cette variante les eprouve tous les deux a la fois.
#      G  est retire du manifeste          -> porteur de grants sans entree
#      B  voit son entree passer a complete:false -> entree listee non exhaustive
#    complete au niveau du manifeste passe a false, puisqu'il omet desormais un
#    principal : le laisser a true serait un mensonge, pas un cas de test.

$nomG = "$Prefix-g-troncature"
$nomB = "$Prefix-b-accorde-non-declare"
$partiels = [System.Collections.Generic.List[object]]::new()
foreach ($p in $principals) {
    if ($p.displayName -eq $nomG) { continue }
    if ($p.displayName -eq $nomB) {
        $copie = [ordered]@{}
        foreach ($k in $p.PSObject.Properties.Name) { $copie[$k] = $p.$k }
        $copie['complete'] = $false
        $copie['comment']  = 'Entree listee mais declaree non exhaustive. Son grant orphelin doit sortir en NotInManifest et non en OverCoverage.'
        $partiels.Add([pscustomobject]$copie)
        continue
    }
    $partiels.Add($p)
}

$manifestePartiel = [ordered]@{
    description = 'Variante partielle. Omet volontairement un porteur de grants et declare une entree non exhaustive, pour eprouver NotInManifest par ses deux chemins.'
    complete    = $false
    asOf        = $manifeste.asOf
    principals  = $partiels.ToArray()
}
$intentPartielPath = Join-Path $OutputDir 'agd-lab-intent-partiel.json'
$manifestePartiel | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $intentPartielPath -Encoding utf8

# ── 4ter. Variante non resolue ──────────────────────────────────────────────
#    NotResolved n'est produit par aucune des deux autres. Ni l'une ni l'autre
#    ne peut l'absorber : dans le manifeste complet, une entree non resolue fait
#    tomber les 27 OverCoverage et le corpus perdrait l'etat qui justifie
#    l'outil ; dans le partiel, elle ferait basculer C, qui y reste OverCoverage
#    alors que B et G tombent, et c'est l'assertion la plus fine du harnais.
#
#    Cette variante est donc le manifeste COMPLET, inchange, plus une seule
#    entree designant un principal inexistant. Une seule variable change, donc
#    ce qu'on observe ne peut venir que d'elle : NotResolved passe a un,
#    completeness a Partial, et overDemonstrable tombe pour TOUS les principaux,
#    y compris ceux dont l'entree est complete. C'est ce dernier point qui
#    distingue la degradation par non-resolution de celle par complete:false.

$fantome = [pscustomobject][ordered]@{
    appId       = '00000000-dead-beef-0000-000000000001'
    displayName = "$Prefix-i-non-resolu"
    complete    = $true
    comment     = 'Entree designant un principal inexistant. Doit compter dans NotResolved, degrader completeness a Partial, et empecher l etablissement d OverCoverage sur tout le tenant.'
    permissions = @()
}

$manifesteNonResolu = [ordered]@{
    description = 'Variante non resolue. Le manifeste complet, plus une seule entree designant un principal inexistant.'
    complete    = $true
    asOf        = $manifeste.asOf
    principals  = @(@($principals.ToArray()) + $fantome)
}
$intentNonResoluPath = Join-Path $OutputDir 'agd-lab-intent-nonresolu.json'
$manifesteNonResolu | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $intentNonResoluPath -Encoding utf8

# ── 5. Registre ─────────────────────────────────────────────────────────────

$registre = foreach ($c in $cas) {
    $appId = if ($c.Nom -eq $nomC) { $EXO_APPID }
             elseif ($appIdParNom.ContainsKey($c.Nom)) { $appIdParNom[$c.Nom] } else { '' }
    [pscustomobject]@{
        cas                          = $c.Cas
        nom                          = $c.Nom
        app_id                       = $appId
        state_attendu_sans_manifeste = $c.Sans
        state_attendu_avec_manifeste = $c.Avec
        state_attendu_partiel        = $c.Partiel
        state_attendu_non_resolu     = $c.NonResolu
        commentaire                  = $c.Commentaire
    }
}
$registre = @($registre)

$registrePath = Join-Path $OutputDir 'agd-lab-registre.csv'
$registre | Export-Csv -LiteralPath $registrePath -NoTypeInformation -Encoding utf8

$registre | Format-Table cas, nom, state_attendu_sans_manifeste, state_attendu_avec_manifeste, state_attendu_partiel, state_attendu_non_resolu -AutoSize

Write-Host "Manifeste complet  : $intentPath  ($($principals.Count) principaux)"
Write-Host "Manifeste partiel  : $intentPartielPath  ($($partiels.Count) principaux)"
Write-Host "Manifeste non resolu : $intentNonResoluPath  ($($principals.Count + 1) principaux, dont un fantome)"
Write-Host "Registre           : $registrePath"
Write-Host ""
Write-Host "Quatre passages. Ce sont les differences entre eux qui prouvent quelque chose,"
Write-Host "pas la conformite de l'un d'eux pris isolement."
Write-Host "  .\agd-lab-analyse.ps1"
Write-Host "  .\agd-lab-analyse.ps1 -IntentPath '$intentPath'"
Write-Host "  .\agd-lab-analyse.ps1 -IntentPath '$intentPartielPath' -Colonne state_attendu_partiel"
Write-Host "  .\agd-lab-analyse.ps1 -IntentPath '$intentNonResoluPath' -Colonne state_attendu_non_resolu"
