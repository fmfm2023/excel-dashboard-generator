# ============================================================
# test-api.ps1 — Script de test de l'API Dashboard Generator
# Usage : .\test-api.ps1
# Prérequis : PowerShell 5+ et l'API en cours d'exécution
# ============================================================

param(
    [string]$ApiUrl   = "http://localhost:5000",
    [string]$TestFile = "$PSScriptRoot\ventes_janvier.xlsx",
    [string]$Email    = "test@example.com"
)

$ErrorActionPreference = "Stop"

# ── Couleurs pour l'affichage ──────────────────────────────────────────────
function Write-Step($msg)    { Write-Host "`n🔵 $msg" -ForegroundColor Cyan }
function Write-Success($msg) { Write-Host "✅ $msg" -ForegroundColor Green }
function Write-Warn($msg)    { Write-Host "⚠️  $msg" -ForegroundColor Yellow }
function Write-Fail($msg)    { Write-Host "❌ $msg" -ForegroundColor Red }

Write-Host "`n========================================" -ForegroundColor Magenta
Write-Host "  📊 Excel Dashboard Generator — TEST" -ForegroundColor Magenta
Write-Host "========================================`n" -ForegroundColor Magenta

# ── Étape 1 : Health check ────────────────────────────────────────────────
Write-Step "Étape 1/4 — Vérification de l'API ($ApiUrl/health)"
try {
    $health = Invoke-RestMethod -Uri "$ApiUrl/health" -Method GET -TimeoutSec 10
    if ($health.status -eq "ok") {
        Write-Success "API opérationnelle (version $($health.version))"
    } else {
        Write-Fail "API répond mais statut inattendu : $($health.status)"
        exit 1
    }
} catch {
    Write-Fail "Impossible de joindre l'API : $_"
    Write-Warn "Lancez d'abord : python generate-dashboard.py"
    exit 1
}

# ── Étape 2 : Vérifier le fichier de test ────────────────────────────────
Write-Step "Étape 2/4 — Vérification du fichier de test ($TestFile)"
if (-not (Test-Path $TestFile)) {
    Write-Fail "Fichier introuvable : $TestFile"
    Write-Warn "Modifiez le paramètre -TestFile ou copiez un fichier CSV/XLSX dans le dossier."
    exit 1
}
$fileInfo = Get-Item $TestFile
$sizeMb = [math]::Round($fileInfo.Length / 1MB, 2)
Write-Success "Fichier trouvé : $($fileInfo.Name) ($sizeMb Mo)"

# ── Étape 3 : Encoder le fichier en base64 ───────────────────────────────
Write-Step "Étape 3/4 — Encodage du fichier en base64..."
$fileBytes = [System.IO.File]::ReadAllBytes($TestFile)
$base64    = [System.Convert]::ToBase64String($fileBytes)
Write-Success "Encodage terminé ($([math]::Round($base64.Length / 1KB, 1)) Ko base64)"

# ── Étape 4 : Appel API et vérification ──────────────────────────────────
Write-Step "Étape 4/4 — Appel de l'API /generate-dashboard..."

$payload = @{
    filename  = $fileInfo.Name
    file_data = $base64
    email     = $Email
    file_type = $fileInfo.Extension.TrimStart(".")
} | ConvertTo-Json -Depth 3

$start = Get-Date

try {
    $response = Invoke-RestMethod `
        -Uri         "$ApiUrl/generate-dashboard" `
        -Method      POST `
        -Body        $payload `
        -ContentType "application/json" `
        -TimeoutSec  180

    $elapsed = [math]::Round(((Get-Date) - $start).TotalSeconds, 1)

    if ($response.status -eq "success") {
        Write-Success "Dashboard généré en ${elapsed}s !"

        # Afficher les KPIs
        Write-Host "`n📊 KPIs détectés :" -ForegroundColor White
        Write-Host "   • Enregistrements   : $($response.kpis.total_rows)" -ForegroundColor Gray
        Write-Host "   • Colonnes totales  : $($response.kpis.total_columns)" -ForegroundColor Gray
        Write-Host "   • Colonnes num.     : $($response.kpis.numeric_columns)" -ForegroundColor Gray
        Write-Host "   • Colonnes catég.   : $($response.kpis.categorical_columns)" -ForegroundColor Gray

        # Sauvegarder le fichier Excel généré
        $outName  = "Dashboard_$($fileInfo.BaseName).xlsx"
        $outPath  = Join-Path $PSScriptRoot $outName
        $excelBytes = [System.Convert]::FromBase64String($response.excel_base64)
        [System.IO.File]::WriteAllBytes($outPath, $excelBytes)
        Write-Success "Fichier Excel sauvegardé : $outPath"

        # Ouvrir le fichier automatiquement
        Write-Host "`n🚀 Ouverture du fichier Excel..." -ForegroundColor Cyan
        Start-Process $outPath

    } elseif ($response.status -eq "error") {
        Write-Fail "L'API a retourné une erreur :"
        Write-Host "   $($response.error_message)" -ForegroundColor Red
        exit 1
    } else {
        Write-Warn "Statut inattendu : $($response.status)"
    }

} catch {
    $elapsed = [math]::Round(((Get-Date) - $start).TotalSeconds, 1)
    Write-Fail "Erreur après ${elapsed}s : $_"

    if ($_.Exception.Response) {
        $statusCode = $_.Exception.Response.StatusCode.value__
        Write-Host "   Code HTTP : $statusCode" -ForegroundColor Red
        try {
            $errBody = $_.ErrorDetails.Message | ConvertFrom-Json
            Write-Host "   Message   : $($errBody.error_message)" -ForegroundColor Red
        } catch {}
    }
    exit 1
}

Write-Host "`n========================================" -ForegroundColor Magenta
Write-Host "  ✅ TEST COMPLET — Tout fonctionne !" -ForegroundColor Green
Write-Host "========================================`n" -ForegroundColor Magenta
