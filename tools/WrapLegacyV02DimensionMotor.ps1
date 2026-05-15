# Copia motor de acotacion desde Planos_Automaticos_v02 (solo lectura) y envuelve en Namespace LegacyV02Dimensioning.
# No modifica el arbol v02.
$ErrorActionPreference = "Stop"
$srcRoot = "C:\Users\Tecnica4\Planos_Automaticos_v02\Services\Dimensioning"
$repoRoot = Split-Path $PSScriptRoot -Parent
$dstRoot = Join-Path $repoRoot "Services\Dimensioning\LegacyV02IsolatedMotor"
$ns = "Extraer_dft_dxf_flatdxf.LegacyV02Dimensioning"

$files = @(
    "DimensioningNormConfig.vb",
    "ProtectedZone2D.vb",
    "ReferenceDrawingDimensioningService.vb",
    "DimensionCandidate.vb",
    "DimensionGeometryReader.vb",
    "DimensionPlacementService.vb",
    "DimensionValidationService.vb",
    "Iso129DimensioningService.vb",
    "DimensionLogger.vb",
    "DimensionExtremePoint.vb",
    "DimensionInsertionConfig.vb",
    "ViewGeometryReader.vb",
    "SolidEdgeContext.vb",
    "DraftViewCollector.vb",
    "DrawingViewGeometryReader.vb",
    "DrawingViewDimensionPlanner.vb",
    "DrawingViewDimensionCreator.vb",
    "DimensionStyleResolver.vb",
    "UniqueDvAutoDimensioningEngine.vb",
    "Une129ArrangeExistingDimensions.vb",
    "DimensionPlacementEngine.vb",
    "DimensionReposition.vb",
    "VerticalExteriorAnchors.vb",
    "RealExtremePointsResolver.vb",
    "DimensionCoordinateDiagnostics.vb",
    "StableExteriorReferences.vb",
    "DimensionManualCompare.vb",
    "ViewPlacementFrame.vb"
)

New-Item -ItemType Directory -Force -Path $dstRoot | Out-Null

foreach ($rel in $files) {
    $srcPath = Join-Path $srcRoot $rel
    if (-not (Test-Path $srcPath)) { throw "Falta archivo fuente: $srcPath" }
    $dstPath = Join-Path $dstRoot $rel
    $raw = Get-Content -LiteralPath $srcPath -Raw -Encoding UTF8
    $lines = $raw -split "`r?`n", -1
    $head = New-Object System.Collections.Generic.List[string]
    $i = 0
    while ($i -lt $lines.Length) {
        $t = $lines[$i].TrimStart()
        if ($t.Length -eq 0) { $i++; continue }
        if ($t.StartsWith("Option ")) { $head.Add($lines[$i]); $i++; continue }
        if ($t.StartsWith("Imports ")) { $head.Add($lines[$i]); $i++; continue }
        break
    }
    $body = @()
    if ($i -lt $lines.Length) { $body = $lines[$i..($lines.Length-1)] }
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($h in $head) { $out.Add($h) }
    $out.Add("Namespace $ns")
    $out.Add("' --- Copia aislada desde Planos_Automaticos_v02 (no editar v02); motor independiente del arbol Services\Dimensioning vigente. ---")
    foreach ($b in $body) { $out.Add($b) }
    $out.Add("End Namespace")
    $text = ($out -join "`r`n") + "`r`n"
    Set-Content -LiteralPath $dstPath -Value $text -Encoding UTF8
    Write-Host "OK $rel"
}

Write-Host "Destino: $dstRoot"
