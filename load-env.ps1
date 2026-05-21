# Script para cargar variables de entorno desde .env
# Uso: . .\load-env.ps1

$envFile = ".env"

if (Test-Path $envFile) {
    Write-Host "Cargando variables de entorno desde $envFile..."
    Get-Content $envFile | ForEach-Object {
        if ($_ -match "^\s*#" -or $_ -match "^\s*$") {
            return  # Skip comments and empty lines
        }
        $name, $value = $_.Split("=", 2)
        $name = $name.Trim()
        $value = $value.Trim()
        [System.Environment]::SetEnvironmentVariable($name, $value, "Process")
        Write-Host "  $name = ✓" -ForegroundColor Green
    }
    Write-Host "Variables de entorno cargadas." -ForegroundColor Green
} else {
    Write-Host "Archivo .env no encontrado. Crea uno basado en .env.example" -ForegroundColor Yellow
}
