# pdfPlumber Extraction Method Design

> **Goal:** Add pdfPlumber as an alternative extraction method alongside Nemotron-Parse, allowing users to extract text-based PDFs more efficiently without OCR overhead.

---

## Architecture

### Current Pipeline
```
PDF → JPEG (base64) → Nemotron-Parse (vision) → Markdown
                                                     ↓
                                            Text Model (NVIDIA/OpenRouter) → JSON
```

### New Option
```
PDF → pdfPlumber → Text (raw) → Text Model (NVIDIA/OpenRouter) → JSON
```

### Fallback Logic
If pdfPlumber extracts 0 significant lines (PDF is scanned/has no text layer), automatically fallback to Nemotron-Parse. Result includes `extraction_method` field indicating which method was actually used.

---

## User Interface

### Sidebar Changes

Add a new radio button at the top of the sidebar (after NVIDIA API key status):

```
🔹 MÉTODO DE EXTRACCIÓN
   ◉ Nemotron-Parse + Text Model
   ○ pdfPlumber
```

**Conditional Display Logic:**

- **If "Nemotron-Parse + Text Model" selected:**
  - Show: "Modelo OpenRouter" section (selectbox + ➕/➖)
  - Show: "Modelo NVIDIA" section (selectbox + ➕/➖)
  - Show: "Proveedor paso 2" radio (choose NVIDIA or OpenRouter for step 2)

- **If "pdfPlumber" selected:**
  - Hide: "Modelo OpenRouter" section
  - Hide: "Modelo NVIDIA" section
  - Hide: "Proveedor paso 2" radio
  - Show: Info label "Método: pdfPlumber"
  - Note: Step 2 will use a sensible default or stored preference

### Session State

Store in `st.session_state`:
- `extraction_method`: "nemotron" | "pdfplumber"

---

## Data Flow

### New Function: `_extract_data_pdfplumber()`

Location: `extractor.py`

**Signature:**
```python
def _extract_data_pdfplumber(
    pdf_bytes: bytes,
    step2_provider: str,          # "nvidia" | "openrouter"
    step2_model: str,             # model ID for step 2
    openrouter_api_key: str = "",
    nvidia_api_key: str = "",
    on_status=None,
) -> dict:
```

**Logic:**
1. Extract text using pdfplumber
2. If text is empty or minimal, fallback to `_extract_data_nvidia()`
3. Otherwise, send extracted text to step2_model (text model)
4. Return JSON with `extraction_method` field

### Modified Function: `extract_data()`

Add new parameter:
```python
def extract_data(
    pdf_bytes: bytes,
    extraction_method: str = "nemotron",  # "nemotron" | "pdfplumber"
    openrouter_model: str,
    openrouter_api_key: str,
    nvidia_api_key: str = "",
    step2_provider: str = "nvidia",
    step2_nvidia_model: str = "",
    on_status=None,
) -> dict:
```

Routes to either `_extract_data_nvidia()` or `_extract_data_pdfplumber()` based on `extraction_method`.

### Response Structure

JSON response includes new field:
```json
{
  "of_number": "OF26004215",
  "extraction_method": "pdfplumber",
  "n_pedido": "2026000508 / 1",
  "articulo": "...",
  ...
}
```

Or on fallback:
```json
{
  "of_number": "OF26004215",
  "extraction_method": "nemotron-parse (fallback)",
  ...
}
```

---

## Error Handling & Fallback

### When to Fallback

Fallback to Nemotron-Parse if:
1. pdfPlumber raises an exception (corrupted PDF, unsupported format)
2. Extracted text is empty or contains 0 significant lines
3. Text extraction produces less than a threshold (e.g., 10 characters per page)

### Threshold Definition

```python
MIN_TEXT_CHARS_PER_PAGE = 50  # If avg < 50 chars/page, fallback
```

### Fallback Behavior

Automatic and silent to the extraction logic, but `extraction_method` field indicates it happened.

---

## UI Integration (app.py)

### Import Display

After PDF extraction completes, display method used:

```python
if data:
    method_used = data.get("extraction_method", "unknown")
    if "fallback" in method_used:
        st.warning(f"⚠️ Método: {method_used}")
    else:
        st.success(f"✓ Método: {method_used}")
```

### Default Step 2 for pdfPlumber

When pdfPlumber is selected, step 2 will default to:
- NVIDIA text model (since no OpenRouter config shown)
- Or retrieve stored preference from session state

---

## Testing

### Unit Tests (test_extractor.py)

1. Test pdfPlumber extraction with text-based PDF → returns correct JSON with `extraction_method: "pdfplumber"`
2. Test pdfPlumber with scanned PDF → fallback to Nemotron-Parse
3. Test corrupted PDF → fallback to Nemotron-Parse with appropriate error in `extraction_method`
4. Test `extract_data()` routing based on `extraction_method` parameter

### Integration Tests (app.py)

1. Select "pdfPlumber" → verify UI hides model selectors
2. Upload text-based PDF → verify extraction_method shows "pdfplumber"
3. Upload scanned PDF → verify automatic fallback and warning display
4. Switch back to Nemotron-Parse → verify UI shows model selectors

---

## Dependencies

- **pdfplumber** (add to requirements.txt if not present)
- Existing: openai, requests, pdf2image, Pillow

---

## Rollout

1. Add pdfPlumber extraction method alongside existing Nemotron-Parse
2. Default to Nemotron-Parse (backward compatible)
3. Users can opt-in to pdfPlumber via UI radio button
4. Automatic fallback ensures robustness

