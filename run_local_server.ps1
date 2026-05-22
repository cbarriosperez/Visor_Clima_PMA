# Script de PowerShell para levantar un servidor web local y evadir restricciones de CORS.
# Ejecuta este script desde una consola de PowerShell en la raíz de tu proyecto.

$port = 8080
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$port/")

try {
    $listener.Start()
    Write-Host "==========================================================" -ForegroundColor Green
    Write-Host "  SERVIDOR LOCAL DE DESARROLLO - VISOR DE CLIMA" -ForegroundColor Cyan
    Write-Host "==========================================================" -ForegroundColor Green
    Write-Host "Servidor activo en: http://localhost:$port/" -ForegroundColor Yellow
    Write-Host "Para detener el servidor: Presiona Ctrl + C en esta consola" -ForegroundColor DarkGray
    Write-Host "==========================================================" -ForegroundColor Green
    
    # Abre el navegador por defecto automáticamente
    Start-Process "http://localhost:$port/"
    
    $currentDir = (Get-Location).Path
    
    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $request = $context.Request
        $response = $context.Response
        
        # Obtener ruta de archivo relativa y limpiar
        $urlPath = $request.Url.LocalPath.TrimStart('/')
        if ($urlPath -eq "") { $urlPath = "index.html" }
        
        # Sanitizar y verificar que la ruta se mantenga en el directorio actual
        $filePath = [System.IO.Path]::GetFullPath((Join-Path $currentDir $urlPath))
        if (-not $filePath.StartsWith($currentDir)) {
            $response.StatusCode = 403
            $response.Close()
            continue
        }
        
        if (Test-Path $filePath -PathType Leaf) {
            $bytes = [System.IO.File]::ReadAllBytes($filePath)
            
            # Mapear Content-Type
            $ext = [System.IO.Path]::GetExtension($filePath).ToLower()
            $contentType = switch ($ext) {
                ".html" { "text/html; charset=utf-8" }
                ".css"  { "text/css; charset=utf-8" }
                ".js"   { "application/javascript; charset=utf-8" }
                ".csv"  { "text/csv; charset=utf-8" }
                ".xlsx" { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }
                ".png"  { "image/png" }
                ".jpg"  { "image/jpeg" }
                default { "application/octet-stream" }
            }
            
            $response.ContentType = $contentType
            $response.ContentLength64 = $bytes.Length
            $response.OutputStream.Write($bytes, 0, $bytes.Length)
            Write-Host "[200] Servido: /$urlPath ($contentType)" -ForegroundColor Gray
        } else {
            $response.StatusCode = 404
            Write-Host "[404] No encontrado: /$urlPath" -ForegroundColor Red
        }
        $response.Close()
    }
} catch {
    Write-Host "Error al iniciar el servidor: $($_.Exception.Message)" -ForegroundColor Red
} finally {
    $listener.Stop()
    Write-Host "Servidor detenido." -ForegroundColor Yellow
}
