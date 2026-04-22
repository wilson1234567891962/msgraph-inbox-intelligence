# 1. Conexión (usamos la sesión activa)
$contexto = Get-MgContext
if ($null -eq $contexto) { Connect-MgGraph -Scopes "Mail.Read" }

Write-Host "[-] Iniciando escaneo defensivo de 7,339 correos..." -ForegroundColor Cyan
$uri = "https://graph.microsoft.com/v1.0/me/messages?`$select=from&`$top=1000"
$estadisticas = @{} 
$totalProcesados = 0
$totalSaltados = 0

try {
    while ($uri) {
        $data = Invoke-MgRestMethod -Method Get -Uri $uri
        
        # Validamos que el bloque contenga datos
        if ($null -ne $data.value) {
            foreach ($msg in $data.value) {
                # VALIDACIÓN DEFENSIVA: Verificamos toda la jerarquía antes de procesar
                if ($null -ne $msg.from -and $null -ne $msg.from.emailAddress -and $null -ne $msg.from.emailAddress.address) {
                    $email = $msg.from.emailAddress.address.ToLower()
                    $estadisticas[$email] = [int]$estadisticas[$email] + 1
                    $totalProcesados++
                } else {
                    $totalSaltados++
                }
            }
        }

        Write-Host "    -> [$totalProcesados] procesados... ($totalSaltados borradores/nulos saltados)" -ForegroundColor Gray
        
        # Paginación
        $uri = $data.'@odata.nextLink'
    }

    # 2. Construcción del reporte para Excel
    Write-Host "[-] Generando tabla de frecuencias..." -ForegroundColor Cyan
    $reporteFinal = $estadisticas.GetEnumerator() | ForEach-Object {
        [PSCustomObject]@{
            Remitente      = $_.Key
            Dominio        = $_.Key.Split('@')[-1]
            Cantidad       = $_.Value
            Prioridad      = if ($_.Value -gt 100) { "CRÍTICO (Borrar)" } elseif ($_.Value -gt 50) { "Newsletter" } else { "Normal" }
        }
    } | Sort-Object Cantidad -Descending

    # 3. Exportación a CSV
    $rutaCSV = "$env:USERPROFILE\Downloads\Auditoria_Final_Wilson.csv"
    $reporteFinal | Export-Csv -Path $rutaCSV -NoTypeInformation -Encoding utf8 -Delimiter ","

    Write-Host "`n" + "="*60
    Write-Host "¡PROCESO FINALIZADO EXITOSAMENTE!" -ForegroundColor Green
    Write-Host "Mensajes analizados: $totalProcesados"
    Write-Host "Mensajes sin remitente (saltados): $totalSaltados"
    Write-Host "Reporte listo en: $rutaCSV" -ForegroundColor Yellow
    Write-Host "="*60 + "`n"

    # Muestra de los 15 más pesados
    $reporteFinal | select -First 15 | ft -AutoSize

} catch {
    Write-Host "[!] Error técnico en el flujo: $($_.Exception.Message)" -ForegroundColor Red
}
