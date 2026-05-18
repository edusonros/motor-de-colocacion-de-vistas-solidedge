# 1. Configuración de rutas (Ahora detecta la carpeta donde esté el script)
$CarpetaActual = $PSScriptRoot
if([string]::IsNullOrEmpty($CarpetaActual)) { $CarpetaActual = Get-Location }

$archivoOriginal = Join-Path $CarpetaActual "SESDK_COMPLETO_UNIFICADO.txt"
$prefijoSalida = Join-Path $CarpetaActual "SESDK_COMPLETO_UNIFICADO_parte_"

# Verificar si el archivo existe antes de empezar
if (-not (Test-Path $archivoOriginal)) {
    Write-Host "ERROR: No encuentro el archivo '$archivoOriginal'" -ForegroundColor Red
    Write-Host "Asegúrate de que el .txt esté en la misma carpeta que este script." -ForegroundColor Yellow
    pause
    exit
}

# 2. Contar líneas totales para calcular el tamaño de cada trozo
Write-Host "Contando líneas... esto puede tardar un poco." -ForegroundColor Cyan
$lineasTotales = (Get-Content $archivoOriginal | Measure-Object -Line).Lines

# 3. Calcular cuántas líneas van por cada una de las 10 partes
$lineasPorArchivo = [Math]::Ceiling($lineasTotales / 10)

# 4. Dividir el archivo
Write-Host "Iniciando división en 10 partes ($lineasPorArchivo líneas cada una)..." -ForegroundColor Green

$lector = [System.IO.File]::OpenText($archivoOriginal)
try {
    for ($i = 1; $i -le 10; $i++) {
        $archivoDestino = "$prefijoSalida$i.txt"
        $contadorCerrado = 0
        
        # Crear cada archivo de salida
        $escritor = [System.IO.File]::CreateText($archivoDestino)
        try {
            while ($contadorCerrado -lt $lineasPorArchivo -and ($linea = $lector.ReadLine()) -ne $null) {
                $escritor.WriteLine($linea)
                $contadorCerrado++
            }
        } finally {
            $escritor.Close()
        }
        Write-Host "Finalizado: $archivoDestino"
    }
} finally {
    $lector.Close()
}

Write-Host "¡Proceso completado!" -ForegroundColor Gold