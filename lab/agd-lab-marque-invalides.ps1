#requires -Version 7.0
<#
    agd-lab-marque-invalides.ps1

    LECTURE-ECRITURE sur un seul fichier local. Ne touche a aucun tenant.

    Marque les enregistrements de observations.json produits par le passage de
    fabrication du 8 septembre qui ont ete fausses par le retour a virgule : sur
    un tableau imbrique, .Count vaut 1 quel que soit le contenu reel, vide
    compris.

    MARQUER PLUTOT QUE SUPPRIMER.
    Effacer ferait disparaitre la trace qu'une mesure a ete fausse, et cette
    trace est elle-meme une information sur la methode. Un fichier de preuve qui
    ne garde que ses bonnes mesures ne prouve plus rien sur sa propre fiabilite.
    Les entrees restent, avec invalide=true et la raison.

    Les valeurs justes existent par ailleurs, produites par agd-lab-verifie puis
    par agd-lab-analyse. Ce script ne les touche pas.
#>

[CmdletBinding()]
param(
    [string] $Path,
    [switch] $WhatIfOnly
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

if (-not $Path) {
    $base = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
    $Path = Join-Path $base 'agd-lab-out\agd-lab-observations.json'
}
if (-not (Test-Path -LiteralPath $Path)) {
    throw "Fichier introuvable. Fourni : $Path. Resolu en : $([System.IO.Path]::GetFullPath($Path, (Get-Location).Path))."
}
Write-Host "Fichier : $Path"

# Les entrees visees : produites par agd-lab-fabrique, donc sans champ source.
# La liste a ete etablie en remontant a l'origine de CHAQUE valeur dans le
# script, pas en retenant les faits qui semblaient concernes. La difference
# n'est pas theorique : grants_apres_retrait figurait dans ma premiere liste et
# n'y a pas sa place. Il vient d'un appel direct au cmdlet, ligne 340, sans
# passer par une fonction a retour virgule. Sa valeur est juste, et la marquer
# aurait corrompu l'artefact dans l'autre sens.
#
# Les raisons different aussi. Trois valeurs sont directement issues d'un
# tableau imbrique. La quatrieme est derivee : son premier terme etait juste,
# c'est le second qui l'a empoisonnee. Une raison unique effacerait cette
# distinction, qui est ce qu'on apprend en relisant le fichier.

$nesting = 'Mesure invalide. Comptage effectue sur un tableau imbrique, ou .Count vaut 1 quelle que soit la collection reelle, vide comprise. Valeur juste etablie ensuite par agd-lab-verifie puis agd-lab-analyse.'

$faitsVises = @{
    'B' = @{
        'declaration_apres_retrait' = $nesting
        'grant_survit_au_retrait'   = 'Conclusion invalide, mais par derivation et non par mesure directe. Elle valait ($bAfter.Count -eq 1 -and $bDeclaredAfter.Count -eq 0). Le premier terme etait juste, le second sortait d un tableau imbrique et valait 1 au lieu de 0, ce qui a rendu la conjonction fausse. Le fait inverse a ete etabli ensuite : agd-lab-verifie rend le cas B en granted_not_declared, donc le grant survit bien au retrait de la declaration.'
    }
    'E' = @{
        'declare' = $nesting
        'accorde' = $nesting
    }
}

$entrees = @(Get-Content -LiteralPath $Path -Raw -Encoding utf8 | ConvertFrom-Json)
Write-Host "$($entrees.Count) enregistrement(s)"

$sortie = [System.Collections.Generic.List[object]]::new()
$marques = [System.Collections.Generic.List[object]]::new()

foreach ($e in $entrees) {
    $aSource = [bool]$e.PSObject.Properties['source']
    $vise = $faitsVises.ContainsKey($e.cas) -and $faitsVises[$e.cas].ContainsKey($e.fait) -and -not $aSource

    if ($vise -and -not $e.PSObject.Properties['invalide']) {
        $copie = [ordered]@{}
        foreach ($k in $e.PSObject.Properties.Name) { $copie[$k] = $e.$k }
        $copie['invalide'] = $true
        $copie['raison']   = $faitsVises[$e.cas][$e.fait]
        $obj = [pscustomobject]$copie
        $sortie.Add($obj)
        $marques.Add($obj)
    } else {
        $sortie.Add($e)
    }
}

if ($marques.Count -eq 0) {
    Write-Host "Aucun enregistrement a marquer. Deja fait, ou le fichier ne contient pas ce passage."
    return
}

$marques | Format-Table horodatage, cas, fait, valeur -AutoSize
Write-Host "grants_apres_retrait n'est PAS marque : appel direct au cmdlet, sa valeur est juste."

if ($WhatIfOnly) {
    Write-Host "$($marques.Count) enregistrement(s) seraient marque(s). Rien n'a ete ecrit."
    return
}

# Sauvegarde avant reecriture. Le fichier est un artefact : on ne le remplace
# pas sans garder ce qu'il etait.
$sauvegarde = "$Path.$(Get-Date -Format 'yyyyMMdd-HHmmss').bak"
Copy-Item -LiteralPath $Path -Destination $sauvegarde
$sortie | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $Path -Encoding utf8

# Relire pour verifier, plutot que supposer que l'ecriture a fait ce qu'on croit.
$relu = @(Get-Content -LiteralPath $Path -Raw -Encoding utf8 | ConvertFrom-Json)
$relusMarques = @($relu | Where-Object { $_.PSObject.Properties['invalide'] -and $_.invalide })

Write-Host "$($marques.Count) enregistrement(s) marque(s), $($relusMarques.Count) relu(s) marque(s)"
if ($relu.Count -ne $entrees.Count) {
    Write-Warning "Le fichier relu porte $($relu.Count) enregistrements contre $($entrees.Count) avant. Sauvegarde : $sauvegarde"
} elseif ($relusMarques.Count -ne $marques.Count) {
    Write-Warning "Ecart entre ce qui a ete marque et ce qui se relit. Sauvegarde : $sauvegarde"
} else {
    Write-Host "Sauvegarde : $sauvegarde"
}
