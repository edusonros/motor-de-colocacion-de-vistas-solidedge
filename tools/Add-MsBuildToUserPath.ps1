# Añade MSBuild de Visual Studio al PATH del usuario (persistente).
$msb = 'C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin'
$cur = [Environment]::GetEnvironmentVariable('Path', 'User')
if ($null -eq $cur) { $cur = '' }

$norm = { param($p) if ($null -eq $p) { '' } else { $p.TrimEnd('\') } }
$already = $false
foreach ($part in ($cur -split ';')) {
    if ((& $norm $part) -ieq (& $norm $msb)) {
        $already = $true
        break
    }
}

if ($already) {
    Write-Host "Ya estaba en PATH de usuario: $msb"
    exit 0
}

$newPath = if ([string]::IsNullOrWhiteSpace($cur)) { $msb } else { $cur.TrimEnd(';') + ';' + $msb }
[Environment]::SetEnvironmentVariable('Path', $newPath, 'User')
Write-Host "Añadido al PATH de usuario: $msb"
Write-Host "Cierra y vuelve a abrir terminales (y Cursor) para que reconozcan 'msbuild'."
