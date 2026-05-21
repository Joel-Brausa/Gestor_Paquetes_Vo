# Modelos OpenRouter para Voest Paquets

## Problema Actual

La API key en `.env` ha alcanzado su límite de cuota (Error 403). Necesitas:
1. Recargar créditos en OpenRouter, O
2. Usar una API key válida con créditos disponibles

## Modelos Recomendados

### Con Visión (Para procesar PDFs):

**Gratuitos (sin créditos):**
- `meta-llama/llama-2-7b-chat:free` - LLaMA 2 7B
- `meta-llama/llama-3-8b:free` - LLaMA 3 8B (recomendado)

**Con créditos (mejores resultados):**
- `openai/gpt-4o-mini` - GPT-4o Mini (rápido, económico) ⭐ RECOMENDADO
- `openai/gpt-4o` - GPT-4o Vision (más preciso, más caro)
- `anthropic/claude-3-5-sonnet` - Claude 3.5 Sonnet (excelente, moderado)
- `google/gemini-2.0-flash` - Gemini 2.0 Flash (rápido, económico)

### Configuración

1. **Cambiar modelo predeterminado:**
   - Edita `config.py`:
   ```python
   DEFAULT_MODEL = "meta-llama/llama-3-8b:free"  # O tu modelo elegido
   ```

2. **Cambiar en la interfaz:**
   - Abre la app
   - En el sidebar, usa "Nuevo modelo" para agregar uno diferente
   - Selecciona en el dropdown "Modelo activo"

## Resolución del Problema

### Opción A: Recargar créditos (Rápido)
1. Ve a https://openrouter.ai/account/billing/overview
2. Recarga créditos con tu método de pago
3. Usa el mismo API key

### Opción B: Nueva API key
1. Crea una nueva clave en https://openrouter.ai/settings/keys
2. Actualiza `.env`:
   ```
   OPENROUTER_API_KEY=sk-or-v1-nueva-clave-aqui
   ```
3. Reinicia la app

### Opción C: Usar modelo gratuito
```
Cambiar DEFAULT_MODEL a "meta-llama/llama-3-8b:free"
(Sin créditos necesarios, pero más lento)
```

## Testear Conectividad

Ejecuta el script de prueba:
```powershell
. .\load-env.ps1
python test_api.py
```

Debe mostrar: `[OK] API call successful`
