# Configuración de Voest Paquets

## Variable de Entorno: OPENROUTER_API_KEY

La aplicación requiere una API key de OpenRouter para funcionar. Esta se puede cargar de dos formas:

### Opción 1: Variable de Entorno (Recomendado)

#### En PowerShell (una sola sesión):
```powershell
$env:OPENROUTER_API_KEY = "tu-api-key-aqui"
python app.py
```

#### Con script automático (.env):
1. El archivo `.env` contiene la API key (ya configurado)
2. Antes de ejecutar la app, carga las variables:
```powershell
. .\load-env.ps1
python app.py
```

#### Permanentemente en el sistema (Windows):
```powershell
[System.Environment]::SetEnvironmentVariable("OPENROUTER_API_KEY", "tu-api-key-aqui", "User")
```

### Opción 2: Interfaz Gráfica

1. Ejecutar la app: `python app.py`
2. En la barra lateral, ingresar la API key en el campo "API Key OpenRouter"
3. Hacer clic en "Guardar API Key"

La clave se guardará en el archivo `.env` para futuras sesiones.

## Archivos de Configuración

- `.env` - API key local (excluido de git, no se debe commitear)
- `.env.example` - Plantilla de ejemplo
- `load-env.ps1` - Script PowerShell para cargar variables desde .env
- `.gitignore` - Archivos ignorados por git

## Ejecución

```bash
# Cargar variables de entorno y ejecutar
. .\load-env.ps1
python app.py
```

La app se abrirá en http://localhost:7860
