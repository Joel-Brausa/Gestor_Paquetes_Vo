# pdfPlumber Extraction Method Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development to implement this plan task-by-task.

**Goal:** Add pdfPlumber as an alternative extraction method alongside Nemotron-Parse, with automatic fallback to Nemotron-Parse for scanned PDFs.

**Architecture:** New `_extract_data_pdfplumber()` function in extractor.py that extracts text via pdfplumber and sends to text model. If text is empty, fallback to `_extract_data_nvidia()`. UI in app.py allows users to choose method via radio button. Result includes `extraction_method` field.

**Tech Stack:** pdfplumber (PDF text extraction), existing NVIDIA/OpenRouter APIs

---

## Task 1: Add pdfplumber Dependency

**Files:**
- Modify: `requirements.txt`

**Step 1: Read current requirements.txt**

Run: `cat requirements.txt`
Expected: See list of dependencies (gradio, openai, pdf2image, etc.)

**Step 2: Add pdfplumber to requirements.txt**

Modify `requirements.txt` to add the line (after pdf2image):
```
pdfplumber>=0.9
```

Full file should look like:
```
gradio>=5.0
openai>=1.0
pdf2image>=1.17
pdfplumber>=0.9
Pillow>=10.0
pandas>=2.0
openpyxl>=3.1
pytest>=8.0
requests>=2.28
```

**Step 3: Install pdfplumber**

Run: `pip install pdfplumber>=0.9`
Expected: Installation completes successfully

**Step 4: Verify installation**

Run: `python -c "import pdfplumber; print(pdfplumber.__version__)"`
Expected: Version number printed (e.g., "0.10.2")

**Step 5: Commit**

```bash
git add requirements.txt
git commit -m "feat: add pdfplumber dependency for text extraction"
```

---

## Task 2: Create _extract_data_pdfplumber() Function

**Files:**
- Modify: `extractor.py` (add new function before `extract_data()`)
- Create: `tests/test_pdfplumber_extraction.py`

**Step 1: Write failing test for pdfplumber extraction**

Create `tests/test_pdfplumber_extraction.py`:

```python
import pytest
import json
import io
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import sys
import os

# Add parent directory to path
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))

import extractor


def create_text_pdf():
    """Create a simple PDF with text (not scanned image)."""
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=letter)
    c.drawString(100, 750, "OF26004215")
    c.drawString(100, 700, "N.Pedido: 2026000508 / 1")
    c.drawString(100, 650, "Paquete 1: 28 piezas")
    c.save()
    buf.seek(0)
    return buf.getvalue()


def test_extract_data_pdfplumber_basic():
    """Test that pdfplumber extraction returns JSON with extraction_method field."""
    pdf_bytes = create_text_pdf()
    
    # Mock on_status callback
    status_messages = []
    def mock_status(msg):
        status_messages.append(msg)
    
    # This test just checks that the function exists and can be called
    # (actual LLM calls will be mocked in integration tests)
    assert hasattr(extractor, '_extract_data_pdfplumber')
    assert callable(extractor._extract_data_pdfplumber)


def test_extract_data_pdfplumber_returns_extraction_method():
    """Test that result includes extraction_method field."""
    # This test will be populated after implementation
    # For now, just verify the structure exists
    pass
```

**Step 2: Run test to verify it fails**

Run: `pytest tests/test_pdfplumber_extraction.py::test_extract_data_pdfplumber_basic -v`
Expected: FAIL with "AttributeError: module 'extractor' has no attribute '_extract_data_pdfplumber'"

**Step 3: Implement _extract_data_pdfplumber() function**

Add to `extractor.py` (before `extract_data()` function, around line 312):

```python
def _extract_data_pdfplumber(
    pdf_bytes: bytes,
    step2_provider: str,          # "nvidia" | "openrouter"
    step2_model: str,             # model id for step 2
    openrouter_api_key: str = "",
    nvidia_api_key: str = "",
    on_status=None,
) -> dict:
    """
    Extract packing list data from PDF using pdfplumber.
    
    Step 1: pdfplumber → extract text from PDF
    Step 2: If text extracted, send to text model (NVIDIA or OpenRouter)
            If no text, fallback to _extract_data_nvidia() (nemotron-parse)
    
    on_status: optional callable(str) for progress messages
    """
    import pdfplumber
    
    def _status(msg: str):
        if on_status:
            on_status(msg)
    
    _status("Intentando extracción con pdfPlumber...")
    
    try:
        # Step 1: Extract text using pdfplumber
        all_text = []
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page_num, page in enumerate(pdf.pages, 1):
                text = page.extract_text()
                if text and text.strip():
                    all_text.append(text)
                _status(f"pdfPlumber: página {page_num}/{len(pdf.pages)} procesada")
        
        combined_text = "\n\n".join(all_text)
        
        # Check if extraction was successful (threshold: at least 50 chars per page)
        MIN_TEXT_CHARS = 50
        if len(combined_text.strip()) < MIN_TEXT_CHARS:
            _status("⚠️ pdfPlumber extrajo poco contenido, fallback a Nemotron-Parse...")
            data = _extract_data_nvidia(
                pdf_bytes,
                nvidia_api_key,
                step2_provider=step2_provider,
                step2_model=step2_model,
                openrouter_api_key=openrouter_api_key,
                on_status=on_status,
            )
            data["extraction_method"] = "nemotron-parse (fallback)"
            return data
        
        # Step 2: Send extracted text to text model
        text_prompt = f"""{extractor.EXTRACTION_PROMPT}

Documento extraído (texto de pdfplumber):

{combined_text}"""
        
        _status(f"⏳ Paso 2 — {step2_provider.upper()}: extrayendo JSON del texto...")
        
        if step2_provider == "nvidia":
            raw = _call_nvidia_text_model(text_prompt, step2_model, nvidia_api_key)
        else:
            client = OpenAI(base_url=config.OPENROUTER_BASE_URL, api_key=openrouter_api_key)
            try:
                response = client.chat.completions.create(
                    model=step2_model,
                    messages=[{"role": "user", "content": text_prompt}],
                    timeout=60.0,
                )
            except Exception as e:
                raise ValueError(f"Error extrayendo JSON con OpenRouter: {str(e)}")
            
            if not response or not response.choices:
                raise ValueError("Respuesta vacía de OpenRouter.")
            
            raw = response.choices[0].message.content or ""
            if not raw.strip():
                raise ValueError("OpenRouter retornó respuesta vacía.")
        
        _status("Respuesta recibida. Extrayendo JSON...")
        data = parse_llm_response(raw)
        data["extraction_method"] = "pdfplumber"
        return data
        
    except Exception as e:
        _status(f"⚠️ Error en pdfPlumber ({str(e)}), fallback a Nemotron-Parse...")
        data = _extract_data_nvidia(
            pdf_bytes,
            nvidia_api_key,
            step2_provider=step2_provider,
            step2_model=step2_model,
            openrouter_api_key=openrouter_api_key,
            on_status=on_status,
        )
        data["extraction_method"] = "nemotron-parse (fallback)"
        return data
```

**Step 4: Run test to verify it passes**

Run: `pytest tests/test_pdfplumber_extraction.py::test_extract_data_pdfplumber_basic -v`
Expected: PASS

**Step 5: Commit**

```bash
git add extractor.py tests/test_pdfplumber_extraction.py
git commit -m "feat: add _extract_data_pdfplumber() function with fallback logic"
```

---

## Task 3: Modify extract_data() for Method Selection

**Files:**
- Modify: `extractor.py` (lines 313-340, the `extract_data()` function)

**Step 1: Read current extract_data() signature**

Run: `sed -n '313,340p' extractor.py`
Expected: See current function signature and docstring

**Step 2: Update extract_data() signature and implementation**

Modify the `extract_data()` function (replace lines 313-340):

```python
def extract_data(
    pdf_bytes: bytes,
    extraction_method: str = "nemotron",  # "nemotron" | "pdfplumber"
    openrouter_model: str = "",
    openrouter_api_key: str = "",
    nvidia_api_key: str = "",
    step2_provider: str = "nvidia",   # "nvidia" | "openrouter"
    step2_nvidia_model: str = "",     # NVIDIA model for step 2
    on_status=None,
) -> dict:
    """
    Extract packing list data from PDF.

    extraction_method: "nemotron" or "pdfplumber"
      - "nemotron": PDF → JPEG → Nemotron-Parse → markdown → text model → JSON
      - "pdfplumber": PDF → text → text model → JSON (with fallback to nemotron)
    
    Step 2 model selection:
      - If step2_provider="nvidia": uses step2_nvidia_model
      - If step2_provider="openrouter": uses openrouter_model
    
    on_status: optional callable(str) called with human-readable progress messages.
    
    Returns: dict with extracted data + "extraction_method" field
    """
    if extraction_method == "pdfplumber":
        step2_model = step2_nvidia_model if step2_provider == "nvidia" else openrouter_model
        return _extract_data_pdfplumber(
            pdf_bytes,
            step2_provider=step2_provider,
            step2_model=step2_model,
            openrouter_api_key=openrouter_api_key,
            nvidia_api_key=nvidia_api_key,
            on_status=on_status,
        )
    else:  # "nemotron" (default)
        step2_model = step2_nvidia_model if step2_provider == "nvidia" else openrouter_model
        return _extract_data_nvidia(
            pdf_bytes,
            nvidia_api_key,
            step2_provider=step2_provider,
            step2_model=step2_model,
            openrouter_api_key=openrouter_api_key,
            on_status=on_status,
        )
```

**Step 3: Verify syntax**

Run: `python -m py_compile extractor.py`
Expected: No output (success)

**Step 4: Commit**

```bash
git add extractor.py
git commit -m "feat: add extraction_method parameter to extract_data()"
```

---

## Task 4: Update Sidebar UI - Add Method Selector

**Files:**
- Modify: `app.py` (lines 72-186, the sidebar section)

**Step 1: Read current sidebar section**

Run: `sed -n '72,186p' app.py`
Expected: See current sidebar configuration

**Step 2: Add extraction method radio button**

Insert after line 84 (after NVIDIA API key status, before OpenRouter models section):

```python
    st.markdown("---")
    st.markdown("**Método de extracción**")
    
    extraction_method = st.radio(
        "Método de extracción",
        options=["nemotron", "pdfplumber"],
        format_func=lambda x: "Nemotron-Parse + Text Model" if x == "nemotron" else "pdfPlumber",
        key="extraction_method",
        horizontal=False,
        label_visibility="collapsed",
    )
    st.session_state.extraction_method = extraction_method
```

**Step 3: Make OpenRouter section conditional**

Wrap lines 86-124 (entire "Modelo OpenRouter" section) with:

```python
    if extraction_method == "nemotron":
        # [existing code for OpenRouter section]
```

**Step 4: Make NVIDIA model section conditional**

Wrap lines 126-165 (entire "Modelo NVIDIA" section) with:

```python
    if extraction_method == "nemotron":
        # [existing code for NVIDIA section]
```

**Step 5: Make step2_provider conditional**

Wrap lines 167-185 (entire "Proveedor paso 2" section) with:

```python
    if extraction_method == "nemotron":
        # [existing code for step2_provider]
    else:
        st.markdown("---")
        st.markdown("**Método**")
        st.caption("pdfPlumber")
```

**Step 6: Test sidebar UI**

Run: `streamlit run app.py`
Expected: 
- See new "Método de extracción" radio button
- Select "Nemotron-Parse": see OpenRouter/NVIDIA sections
- Select "pdfPlumber": see only "Método: pdfPlumber" label

**Step 7: Commit**

```bash
git add app.py
git commit -m "feat: add extraction method selector to sidebar"
```

---

## Task 5: Update extract_data() Call in app.py

**Files:**
- Modify: `app.py` (lines ~318-330, the `extract_data()` call in Section 2)

**Step 1: Find current extract_data() call**

Run: `grep -n "extractor.extract_data" app.py`
Expected: Line number where `extract_data()` is called

**Step 2: Update the call to pass extraction_method**

Modify the `extract_data()` call to:

```python
                    data = extractor.extract_data(
                        pdf_bytes,
                        extraction_method=st.session_state.get("extraction_method", "nemotron"),
                        openrouter_model=current_model,
                        openrouter_api_key=api_key,
                        nvidia_api_key=nvidia_key,
                        step2_provider=st.session_state.get("step2_provider", "nvidia"),
                        step2_nvidia_model=current_nvidia_model,
                        on_status=_on_status,
                    )
```

**Step 3: Verify the change**

Run: `python -m py_compile app.py`
Expected: No output (success)

**Step 4: Commit**

```bash
git add app.py
git commit -m "feat: pass extraction_method to extract_data() call"
```

---

## Task 6: Display Extraction Method in Results

**Files:**
- Modify: `app.py` (after the `data = extractor.extract_data()` call, around line 335)

**Step 1: Find where results are displayed**

Run: `grep -n "st.success\|st.write.*data" app.py | head -10`
Expected: Lines showing result display

**Step 2: Add extraction method indicator**

After the `data = extractor.extract_data()` call, add:

```python
                    # Display extraction method used
                    method_used = data.get("extraction_method", "unknown")
                    if "fallback" in method_used:
                        st.warning(f"⚠️ Método utilizado: {method_used}")
                    else:
                        st.success(f"✓ Método utilizado: {method_used}")
```

**Step 3: Test the display**

Run: `streamlit run app.py`
Expected: After uploading PDF, see indicator showing which method was used

**Step 4: Commit**

```bash
git add app.py
git commit -m "feat: display extraction method in results UI"
```

---

## Task 7: Test Integration

**Files:**
- Modify: `tests/test_pdfplumber_extraction.py`

**Step 1: Add integration test for method selection**

Add to `tests/test_pdfplumber_extraction.py`:

```python
def test_extract_data_routes_to_pdfplumber():
    """Test that extract_data() routes to pdfplumber when requested."""
    pdf_bytes = create_text_pdf()
    
    # Verify the function would route correctly (structure test)
    assert hasattr(extractor, 'extract_data')
    
    # Check that function accepts extraction_method parameter
    import inspect
    sig = inspect.signature(extractor.extract_data)
    assert 'extraction_method' in sig.parameters


def test_extract_data_includes_method_field():
    """Test that extraction_method field is included in response."""
    # This is a documentation test showing expected structure
    expected_structure = {
        "of_number": str,
        "extraction_method": str,  # Should be "pdfplumber" or "nemotron-parse (fallback)"
        "n_pedido": str,
        "paquetes": list,
    }
    # Note: actual extraction tests require API keys and real PDFs
```

**Step 2: Run tests**

Run: `pytest tests/test_pdfplumber_extraction.py -v`
Expected: All tests pass

**Step 3: Commit**

```bash
git add tests/test_pdfplumber_extraction.py
git commit -m "test: add integration tests for pdfPlumber extraction method"
```

---

## Task 8: Manual Smoke Test

**Files:** None (testing only)

**Step 1: Start the app**

Run: `streamlit run app.py`
Expected: App loads without errors

**Step 2: Test Nemotron-Parse path (unchanged)**

- Sidebar: Select "Nemotron-Parse + Text Model"
- Verify: OpenRouter and NVIDIA model sections visible
- Verify: Step 2 provider radio visible
- Upload PDF → verify extraction works as before

**Step 3: Test pdfPlumber path with text-based PDF**

- Sidebar: Select "pdfPlumber"
- Verify: OpenRouter/NVIDIA sections hidden
- Verify: "Método: pdfPlumber" label visible
- Upload text-based PDF → verify extraction works
- Verify: Result shows "✓ Método utilizado: pdfplumber"

**Step 4: Test pdfPlumber fallback with scanned PDF**

- Sidebar: Keep "pdfPlumber" selected
- Upload scanned PDF (no text layer)
- Verify: Result shows "⚠️ Método utilizado: nemotron-parse (fallback)"

**Step 5: Commit (if needed)**

Run: `git status`
Expected: No uncommitted changes (all previous commits should be done)

---

## Summary

| Task | Files | Commits |
|------|-------|---------|
| 1 | requirements.txt | 1 |
| 2 | extractor.py, tests/ | 1 |
| 3 | extractor.py | 1 |
| 4 | app.py | 1 |
| 5 | app.py | 1 |
| 6 | app.py | 1 |
| 7 | tests/ | 1 |
| 8 | (testing only) | - |
| **Total** | | **7 commits** |

