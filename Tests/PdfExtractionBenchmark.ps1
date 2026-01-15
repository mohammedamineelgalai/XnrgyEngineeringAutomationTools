# =============================================================================
# PDF TABLE EXTRACTION BENCHMARK - XNRGY DXF Verifier
# =============================================================================
# Script PowerShell pour tester differentes methodes d'extraction de tableaux PDF
# Reference: CSV = verite terrain (181 tags)
# PDF: 02-Machines.pdf
# =============================================================================

$ErrorActionPreference = "Continue"

# Fichiers de reference
$PDF_PATH = "C:\Vault\Engineering\Projects\10381\REF13\M02\6-Shop Drawing PDF\Production\BatchPrint\02-Machines.pdf"
$CSV_PATH = "C:\Vault\Engineering\Projects\10381\REF13\M02\5_Exportation\Sheet_Metal_Nesting\Punch\10381-13-M02.csv"

# Pattern pour les tags XNRGY (2-4 lettres + 3-4 chiffres + tiret + 3-4 chiffres)
$TagPattern = "([A-Z]{2,4}\d{3,4}[-_]\d{3,4})"

Write-Host "=" * 80 -ForegroundColor Cyan
Write-Host "PDF TABLE EXTRACTION BENCHMARK - XNRGY DXF Verifier" -ForegroundColor Cyan
Write-Host "=" * 80 -ForegroundColor Cyan
Write-Host ""

# =============================================================================
# 1. CHARGER LA REFERENCE CSV
# =============================================================================
Write-Host "[>] Chargement de la reference CSV..." -ForegroundColor Yellow

if (-not (Test-Path $CSV_PATH)) {
    Write-Host "[-] ERREUR: CSV non trouve: $CSV_PATH" -ForegroundColor Red
    exit 1
}

$csvLines = Get-Content $CSV_PATH | Select-Object -Skip 1
$referenceTagsCSV = @{}

foreach ($line in $csvLines) {
    $parts = $line -split '[;,]'
    if ($parts.Count -gt 0) {
        $tag = $parts[0].Trim().ToUpper().Replace("_", "-")
        if ($tag -match $TagPattern) {
            $referenceTagsCSV[$tag] = $true
        }
    }
}

$totalRefTags = $referenceTagsCSV.Keys.Count
Write-Host "[+] Reference CSV: $totalRefTags tags uniques" -ForegroundColor Green
Write-Host ""

# =============================================================================
# 2. CHARGER PDFPIG (via le projet compile)
# =============================================================================
Write-Host "[>] Chargement de PdfPig..." -ForegroundColor Yellow

$pdfPigPath = "C:\Users\mohammedamine.elgala\source\repos\XnrgyEngineeringAutomationTools\bin\Release\UglyToad.PdfPig.dll"

if (-not (Test-Path $pdfPigPath)) {
    Write-Host "[-] PdfPig non trouve. Build du projet requis." -ForegroundColor Red
    Write-Host "    Chemin attendu: $pdfPigPath" -ForegroundColor Gray
    exit 1
}

Add-Type -Path $pdfPigPath
Write-Host "[+] PdfPig charge avec succes" -ForegroundColor Green
Write-Host ""

# =============================================================================
# 3. OUVRIR LE PDF
# =============================================================================
Write-Host "[>] Ouverture du PDF: $PDF_PATH" -ForegroundColor Yellow

if (-not (Test-Path $PDF_PATH)) {
    Write-Host "[-] PDF non trouve!" -ForegroundColor Red
    exit 1
}

try {
    $document = [UglyToad.PdfPig.PdfDocument]::Open($PDF_PATH)
    $pageCount = $document.NumberOfPages
    Write-Host "[+] PDF ouvert: $pageCount pages" -ForegroundColor Green
    Write-Host ""
}
catch {
    Write-Host "[-] Erreur ouverture PDF: $_" -ForegroundColor Red
    exit 1
}

# =============================================================================
# FONCTION: Evaluer les resultats
# =============================================================================
function Evaluate-Result {
    param(
        [string]$MethodName,
        [hashtable]$ExtractedTags
    )
    
    $totalExtracted = $ExtractedTags.Keys.Count
    
    # True Positives (tags corrects)
    $truePositives = 0
    $correctTags = @()
    foreach ($tag in $ExtractedTags.Keys) {
        if ($referenceTagsCSV.ContainsKey($tag)) {
            $truePositives++
            $correctTags += $tag
        }
    }
    
    # False Negatives (tags manquants)
    $falseNegatives = 0
    $missingTags = @()
    foreach ($tag in $referenceTagsCSV.Keys) {
        if (-not $ExtractedTags.ContainsKey($tag)) {
            $falseNegatives++
            $missingTags += $tag
        }
    }
    
    # False Positives (tags en trop)
    $falsePositives = 0
    $extraTags = @()
    foreach ($tag in $ExtractedTags.Keys) {
        if (-not $referenceTagsCSV.ContainsKey($tag)) {
            $falsePositives++
            $extraTags += $tag
        }
    }
    
    # Metriques
    $precision = if ($totalExtracted -gt 0) { [math]::Round(($truePositives / $totalExtracted) * 100, 1) } else { 0 }
    $recall = if ($totalRefTags -gt 0) { [math]::Round(($truePositives / $totalRefTags) * 100, 1) } else { 0 }
    $f1 = if (($precision + $recall) -gt 0) { [math]::Round((2 * $precision * $recall / ($precision + $recall)), 1) } else { 0 }
    
    return @{
        MethodName = $MethodName
        TotalExtracted = $totalExtracted
        TruePositives = $truePositives
        FalseNegatives = $falseNegatives
        FalsePositives = $falsePositives
        Precision = $precision
        Recall = $recall
        F1Score = $f1
        MissingTags = $missingTags
        ExtraTags = $extraTags
    }
}

# =============================================================================
# METHODE 1: Texte brut simple
# =============================================================================
Write-Host "[>] Test: Methode 1 - Texte brut simple" -ForegroundColor Cyan

$method1Tags = @{}
for ($i = 1; $i -le $pageCount; $i++) {
    $page = $document.GetPage($i)
    $text = $page.Text
    
    $matches = [regex]::Matches($text, $TagPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    foreach ($match in $matches) {
        $tag = $match.Groups[1].Value.ToUpper().Replace("_", "-")
        $method1Tags[$tag] = $true
    }
}

$result1 = Evaluate-Result -MethodName "Methode 1: Texte brut" -ExtractedTags $method1Tags
Write-Host "    Extraits: $($result1.TotalExtracted) | Corrects: $($result1.TruePositives) | Manquants: $($result1.FalseNegatives) | En trop: $($result1.FalsePositives)" -ForegroundColor White
Write-Host "    Precision: $($result1.Precision)% | Recall: $($result1.Recall)% | F1: $($result1.F1Score)%" -ForegroundColor Gray
Write-Host ""

# =============================================================================
# METHODE 2: GetWords()
# =============================================================================
Write-Host "[>] Test: Methode 2 - GetWords()" -ForegroundColor Cyan

$method2Tags = @{}
for ($i = 1; $i -le $pageCount; $i++) {
    $page = $document.GetPage($i)
    $words = $page.GetWords()
    
    foreach ($word in $words) {
        if ($word.Text -match $TagPattern) {
            $tag = $matches[1].ToUpper().Replace("_", "-")
            $method2Tags[$tag] = $true
        }
    }
}

$result2 = Evaluate-Result -MethodName "Methode 2: GetWords" -ExtractedTags $method2Tags
Write-Host "    Extraits: $($result2.TotalExtracted) | Corrects: $($result2.TruePositives) | Manquants: $($result2.FalseNegatives) | En trop: $($result2.FalsePositives)" -ForegroundColor White
Write-Host "    Precision: $($result2.Precision)% | Recall: $($result2.Recall)% | F1: $($result2.F1Score)%" -ForegroundColor Gray
Write-Host ""

# =============================================================================
# METHODE 3: Lettres groupees par Y
# =============================================================================
Write-Host "[>] Test: Methode 3 - Lettres groupees par Y" -ForegroundColor Cyan

$method3Tags = @{}
for ($i = 1; $i -le $pageCount; $i++) {
    $page = $document.GetPage($i)
    $letters = $page.Letters
    
    # Grouper par Y (arrondi)
    $groups = @{}
    foreach ($letter in $letters) {
        $y = [math]::Round($letter.GlyphRectangle.Bottom, 0)
        if (-not $groups.ContainsKey($y)) {
            $groups[$y] = @()
        }
        $groups[$y] += $letter
    }
    
    # Reconstruire les lignes
    foreach ($y in $groups.Keys) {
        $sortedLetters = $groups[$y] | Sort-Object { $_.GlyphRectangle.Left }
        $lineText = ($sortedLetters | ForEach-Object { $_.Value }) -join ""
        
        $matches = [regex]::Matches($lineText, $TagPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        foreach ($match in $matches) {
            $tag = $match.Groups[1].Value.ToUpper().Replace("_", "-")
            $method3Tags[$tag] = $true
        }
    }
}

$result3 = Evaluate-Result -MethodName "Methode 3: Y-Grouping" -ExtractedTags $method3Tags
Write-Host "    Extraits: $($result3.TotalExtracted) | Corrects: $($result3.TruePositives) | Manquants: $($result3.FalseNegatives) | En trop: $($result3.FalsePositives)" -ForegroundColor White
Write-Host "    Precision: $($result3.Precision)% | Recall: $($result3.Recall)% | F1: $($result3.F1Score)%" -ForegroundColor Gray
Write-Host ""

# =============================================================================
# METHODE 4: Mots groupes par Y avec tolerance
# =============================================================================
Write-Host "[>] Test: Methode 4 - Mots groupes par Y (tolerance 5)" -ForegroundColor Cyan

$method4Tags = @{}
for ($i = 1; $i -le $pageCount; $i++) {
    $page = $document.GetPage($i)
    $words = $page.GetWords()
    
    # Grouper par Y avec tolerance
    $groups = @{}
    foreach ($word in $words) {
        $y = [math]::Round($word.BoundingBox.Bottom / 5) * 5
        if (-not $groups.ContainsKey($y)) {
            $groups[$y] = @()
        }
        $groups[$y] += $word
    }
    
    # Reconstruire les lignes
    foreach ($y in $groups.Keys) {
        $sortedWords = $groups[$y] | Sort-Object { $_.BoundingBox.Left }
        $lineText = ($sortedWords | ForEach-Object { $_.Text }) -join " "
        
        $matches = [regex]::Matches($lineText, $TagPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        foreach ($match in $matches) {
            $tag = $match.Groups[1].Value.ToUpper().Replace("_", "-")
            $method4Tags[$tag] = $true
        }
    }
}

$result4 = Evaluate-Result -MethodName "Methode 4: Y-Tolerance" -ExtractedTags $method4Tags
Write-Host "    Extraits: $($result4.TotalExtracted) | Corrects: $($result4.TruePositives) | Manquants: $($result4.FalseNegatives) | En trop: $($result4.FalsePositives)" -ForegroundColor White
Write-Host "    Precision: $($result4.Precision)% | Recall: $($result4.Recall)% | F1: $($result4.F1Score)%" -ForegroundColor Gray
Write-Host ""

# =============================================================================
# METHODE 5: Combinee (toutes les approches)
# =============================================================================
Write-Host "[>] Test: Methode 5 - Combinee" -ForegroundColor Cyan

$method5Tags = @{}

# Ajouter tous les tags des methodes precedentes
foreach ($tag in $method1Tags.Keys) { $method5Tags[$tag] = $true }
foreach ($tag in $method2Tags.Keys) { $method5Tags[$tag] = $true }
foreach ($tag in $method3Tags.Keys) { $method5Tags[$tag] = $true }
foreach ($tag in $method4Tags.Keys) { $method5Tags[$tag] = $true }

$result5 = Evaluate-Result -MethodName "Methode 5: Combinee" -ExtractedTags $method5Tags
Write-Host "    Extraits: $($result5.TotalExtracted) | Corrects: $($result5.TruePositives) | Manquants: $($result5.FalseNegatives) | En trop: $($result5.FalsePositives)" -ForegroundColor White
Write-Host "    Precision: $($result5.Precision)% | Recall: $($result5.Recall)% | F1: $($result5.F1Score)%" -ForegroundColor Gray
Write-Host ""

# =============================================================================
# FERMER LE DOCUMENT
# =============================================================================
$document.Dispose()

# =============================================================================
# RESUME COMPARATIF
# =============================================================================
Write-Host "=" * 80 -ForegroundColor Cyan
Write-Host "RESUME COMPARATIF" -ForegroundColor Cyan
Write-Host "=" * 80 -ForegroundColor Cyan
Write-Host ""
Write-Host "Reference CSV: $totalRefTags tags uniques" -ForegroundColor White
Write-Host ""

$allResults = @($result1, $result2, $result3, $result4, $result5) | Sort-Object { $_.F1Score } -Descending

Write-Host ("{0,-5} {1,-30} {2,-10} {3,-10} {4,-10} {5,-10}" -f "Rang", "Methode", "Recall", "Precision", "F1", "Manquants") -ForegroundColor Yellow
Write-Host ("-" * 85) -ForegroundColor Gray

$rank = 1
foreach ($result in $allResults) {
    $color = if ($rank -eq 1) { "Green" } else { "White" }
    Write-Host ("{0,-5} {1,-30} {2,7}% {3,9}% {4,7}% {5,8}" -f $rank, $result.MethodName, $result.Recall, $result.Precision, $result.F1Score, $result.FalseNegatives) -ForegroundColor $color
    $rank++
}

Write-Host ""
$best = $allResults[0]
Write-Host "[+] MEILLEURE METHODE: $($best.MethodName)" -ForegroundColor Green
Write-Host "    Recall: $($best.Recall)% ($($best.TruePositives)/$totalRefTags tags trouves)" -ForegroundColor Green
Write-Host "    Precision: $($best.Precision)%" -ForegroundColor Green
Write-Host "    Tags manquants: $($best.FalseNegatives)" -ForegroundColor Yellow
Write-Host ""

# =============================================================================
# ANALYSE DES TAGS MANQUANTS
# =============================================================================
if ($best.MissingTags.Count -gt 0) {
    Write-Host "=" * 80 -ForegroundColor Cyan
    Write-Host "TAGS MANQUANTS ($($best.MissingTags.Count))" -ForegroundColor Cyan
    Write-Host "=" * 80 -ForegroundColor Cyan
    Write-Host ""
    
    foreach ($tag in ($best.MissingTags | Sort-Object)) {
        Write-Host "  - $tag" -ForegroundColor Red
    }
    Write-Host ""
}

# Verifier si les tags manquants existent dans le PDF
if ($best.MissingTags.Count -gt 0) {
    Write-Host "Verification dans le PDF brut:" -ForegroundColor Yellow
    
    $document2 = [UglyToad.PdfPig.PdfDocument]::Open($PDF_PATH)
    $fullText = ""
    for ($i = 1; $i -le $document2.NumberOfPages; $i++) {
        $fullText += $document2.GetPage($i).Text
    }
    $document2.Dispose()
    
    $foundInRaw = 0
    foreach ($tag in $best.MissingTags) {
        if ($fullText -match [regex]::Escape($tag)) {
            $foundInRaw++
            Write-Host "  [+] '$tag' existe dans le PDF brut" -ForegroundColor Green
        }
        else {
            Write-Host "  [-] '$tag' N'EXISTE PAS dans le PDF brut" -ForegroundColor Red
        }
    }
    
    Write-Host ""
    Write-Host "Resultat: $foundInRaw/$($best.MissingTags.Count) tags manquants sont dans le PDF brut" -ForegroundColor Yellow
    $vraimentAbsents = $best.MissingTags.Count - $foundInRaw
    Write-Host "=> $vraimentAbsents tags sont VRAIMENT ABSENTS du PDF (pas dans le fichier source)" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "=" * 80 -ForegroundColor Cyan
Write-Host "FIN DU BENCHMARK" -ForegroundColor Cyan
Write-Host "=" * 80 -ForegroundColor Cyan
