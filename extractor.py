import base64
import io
import json
import re
from typing import Optional

import requests
from openai import OpenAI
from PIL import Image

import config

EXTRACTION_PROMPT = """
Extract ALL data from this Brausa packing list. The document may span multiple pages.
Return ONLY a valid JSON object with this exact structure (no markdown, no explanation):

{
  "of_number": "OF26004215",
  "n_pedido": "2026000508 / 1",
  "articulo": "P517780300GL1951 PERFIL C 160x67,5x67,5x19x19 Esp. 3 mm",
  "ref_pedido": "4500096097 CUBE WAREHOUSE T. D11",
  "nota": "-63000",
  "total_piezas": 28,
  "kilos_teoricos": 2284.0,
  "paquetes": [
    {
      "paquete_num": 1,
      "kilos_paquete": 1285.0,
      "lineas": [
        {"linea": 1, "piezas": 3, "longitud": 12.713, "marca": "11889"}
      ]
    }
  ]
}

Rules:
- of_number: the OF number (e.g. "OF26004215")
- kilos_teoricos and kilos_paquete: decimal numbers (use dot as decimal separator, ignore thousand separators)
- longitud: decimal number in meters
- marca: the Marca column value as string
- Include ALL paquetes and ALL lineas from ALL pages — do not skip any
- Pages are separated by "---" in the markdown; treat them as one continuous document
- Do NOT include: cliente, desarrollo, espesor, calidad, fecha_doc
- Return ONLY the JSON, no other text
"""


def pdf_to_base64_images(pdf_bytes: bytes, dpi: int = 150) -> list[str]:
    from pdf2image import convert_from_bytes
    images = convert_from_bytes(pdf_bytes, dpi=dpi, fmt="PNG")
    result = []
    for img in images:
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        result.append(base64.b64encode(buf.getvalue()).decode("utf-8"))
    return result


def _pdf_to_base64_images_for_parse(pdf_bytes: bytes) -> list[str]:
    """
    Convert PDF pages to compressed JPEG base64 images suitable for
    NVIDIA nemotron-parse. Uses lower DPI and JPEG compression to stay
    within the API's ~1 MB per-image limit.
    Auto-reduces quality further if the image is still too large.
    """
    from pdf2image import convert_from_bytes

    DPI = 120
    MAX_BYTES = 900_000   # ~900 KB safety margin below the 1 MB limit
    QUALITY_START = 85

    images = convert_from_bytes(pdf_bytes, dpi=DPI, fmt="PNG")
    result = []
    for img in images:
        quality = QUALITY_START
        while True:
            buf = io.BytesIO()
            img.save(buf, format="JPEG", quality=quality)
            size = buf.tell()
            if size <= MAX_BYTES or quality <= 40:
                break
            # Reduce quality proportionally and try again
            quality = max(40, int(quality * MAX_BYTES / size) - 5)
        result.append(base64.b64encode(buf.getvalue()).decode("utf-8"))
    return result


def _extract_data_openrouter(pdf_bytes: bytes, model: str, api_key: str,
                             on_status=None) -> dict:
    """Extract packing list data using OpenRouter (vision LLM, single call)."""
    def _status(msg: str):
        if on_status:
            on_status(msg)

    _status("Convirtiendo PDF a imágenes...")
    b64_images = pdf_to_base64_images(pdf_bytes)
    if not b64_images:
        raise ValueError("No se pudo convertir el PDF a imágenes.")

    n_pages = len(b64_images)
    _status(f"PDF convertido ({n_pages} página{'s' if n_pages != 1 else ''}). Preparando petición...")

    if not api_key or not api_key.strip():
        raise ValueError("API key de OpenRouter no configurada. Ingresa la clave en el menú lateral.")

    client = OpenAI(base_url=config.OPENROUTER_BASE_URL, api_key=api_key)

    content = [
        {
            "type": "image_url",
            "image_url": {"url": f"data:image/png;base64,{b}"},
        }
        for b in b64_images
    ]
    content.append({"type": "text", "text": EXTRACTION_PROMPT})

    _status(f"⏳ Enviando a OpenRouter [{model}]... (puede tardar 30-90 s)")

    try:
        response = client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": content}],
            timeout=120.0,
        )
    except Exception as e:
        error_msg = str(e)
        if "timeout" in error_msg.lower() or "timed out" in error_msg.lower():
            raise ValueError("Error: La solicitud tardó demasiado (>120s). El modelo o la API está muy lenta. Intenta con un modelo más rápido: openai/gpt-4o-mini o meta-llama/llama-3-8b:free")
        elif "403" in error_msg or "limit exceeded" in error_msg.lower():
            raise ValueError("Error de OpenRouter: Límite de cuota excedido. Recarga créditos en https://openrouter.ai o usa una API key válida.")
        elif "401" in error_msg or "unauthorized" in error_msg.lower():
            raise ValueError("Error de OpenRouter: API key inválida. Verifica la clave en el menú lateral.")
        elif "404" in error_msg or "not found" in error_msg.lower():
            raise ValueError(f"Error de OpenRouter: Modelo '{model}' no encontrado. Selecciona un modelo válido en el menú lateral.")
        else:
            raise ValueError(f"Error al conectar con OpenRouter: {error_msg}")

    if not response or not response.choices:
        raise ValueError("Respuesta vacía de OpenRouter. Verifica la API key y el modelo.")

    raw = response.choices[0].message.content or ""
    if not raw or not raw.strip():
        raise ValueError("OpenRouter retornó una respuesta vacía. Esto podría indicar: API key inválida, modelo no encontrado, o límite de cuota excedido.")

    _status("Respuesta recibida. Extrayendo JSON...")
    return parse_llm_response(raw)


def _call_nemotron_parse(image_b64: str, nvidia_api_key: str) -> str:
    """
    Call nvidia/nemotron-parse for one image and return extracted markdown text.
    Uses markdown_no_bbox tool (text without bounding boxes).
    """
    tool_name = "markdown_no_bbox"
    media_tag = f'<img src="data:image/jpeg;base64,{image_b64}" />'

    headers = {
        "Authorization": f"Bearer {nvidia_api_key}",
        "Accept": "application/json",
        "Content-Type": "application/json",
    }
    payload = {
        "model": config.NVIDIA_PARSE_MODEL,
        "messages": [{"role": "user", "content": media_tag}],
        "tools": [{"type": "function", "function": {"name": tool_name}}],
        "tool_choice": {"type": "function", "function": {"name": tool_name}},
        "max_tokens": 8192,
    }

    response = requests.post(config.NVIDIA_PARSE_URL, headers=headers, json=payload, timeout=120)
    if not response.ok:
        try:
            detail = response.json()
        except Exception:
            detail = response.text[:500]
        raise requests.HTTPError(
            f"{response.status_code} {response.reason} — {detail}",
            response=response,
        )

    choice = response.json()["choices"][0]["message"]

    if not choice.get("tool_calls"):
        return choice.get("content", "")

    args_str = choice["tool_calls"][0]["function"]["arguments"]
    try:
        args = json.loads(args_str)
        # API returns a JSON array; first element is the result
        if isinstance(args, list) and args:
            result = args[0]
        else:
            result = args
        if isinstance(result, dict):
            return result.get("text", result.get("content", str(result)))
        return str(result)
    except json.JSONDecodeError:
        return args_str


def _call_nvidia_text_model(text_prompt: str, model: str, nvidia_api_key: str) -> str:
    """
    Call a NVIDIA text model via the OpenAI-compatible endpoint for markdown→JSON step.
    No image involved — pure text in, text out.
    """
    client = OpenAI(base_url=config.NVIDIA_BASE_URL, api_key=nvidia_api_key)
    try:
        response = client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": text_prompt}],
            timeout=120.0,
        )
    except Exception as e:
        raise ValueError(f"Error al llamar al modelo NVIDIA [{model}]: {str(e)}")

    if not response or not response.choices:
        raise ValueError(f"Respuesta vacía del modelo NVIDIA [{model}].")

    raw = response.choices[0].message.content or ""
    if not raw.strip():
        raise ValueError(f"El modelo NVIDIA [{model}] retornó una respuesta vacía.")
    return raw


def _extract_data_nvidia(
    pdf_bytes: bytes,
    nvidia_api_key: str,
    step2_provider: str,          # "nvidia" | "openrouter"
    step2_model: str,             # model id for step 2
    openrouter_api_key: str = "", # only needed when step2_provider == "openrouter"
    on_status=None,
) -> dict:
    """
    Two-step extraction:
      1. nvidia/nemotron-parse  — image → markdown  (always via NVIDIA API)
      2. step2_provider model   — markdown → JSON    (NVIDIA or OpenRouter)
    """
    def _status(msg: str):
        if on_status:
            on_status(msg)

    _status("Convirtiendo PDF a imágenes...")
    b64_images = _pdf_to_base64_images_for_parse(pdf_bytes)
    if not b64_images:
        raise ValueError("No se pudo convertir el PDF a imágenes.")

    n_pages = len(b64_images)
    _status(f"PDF convertido ({n_pages} página{'s' if n_pages != 1 else ''}).")

    if not nvidia_api_key or not nvidia_api_key.strip():
        raise ValueError("NVIDIA API key no encontrada. Revisa el archivo .env (NVIDIA_API_KEY / NVIDIA_PAI_KEY).")

    if step2_provider == "openrouter" and not (openrouter_api_key or "").strip():
        raise ValueError("API key de OpenRouter requerida para el paso 2 (extracción JSON).")

    # ── Step 1: image → markdown via nemotron-parse ────────────────────────
    all_markdown = []
    for i, b64 in enumerate(b64_images):
        _status(f"⏳ Paso 1/2 — NVIDIA nemotron-parse: página {i + 1}/{n_pages}...")
        try:
            md = _call_nemotron_parse(b64, nvidia_api_key)
            if md:
                all_markdown.append(md)
            _status(f"✓ Paso 1/2 — NVIDIA nemotron-parse: página {i + 1}/{n_pages} procesada")
        except requests.HTTPError as e:
            http_status = e.response.status_code if e.response is not None else "?"
            if http_status == 401:
                raise ValueError("NVIDIA API key inválida. Revisa el archivo .env.")
            elif http_status == 429:
                raise ValueError("NVIDIA API: límite de solicitudes excedido. Espera y vuelve a intentar.")
            raise ValueError(f"Error NVIDIA nemotron-parse (página {i + 1}, HTTP {http_status}): {e}")
        except Exception as e:
            raise ValueError(f"Error en nemotron-parse (página {i + 1}): {str(e)}")

    if not all_markdown:
        raise ValueError("nemotron-parse no devolvió contenido para ninguna página.")

    combined_markdown = "\n\n---\n\n".join(all_markdown)
    page_note = (
        f"The document has {n_pages} page{'s' if n_pages != 1 else ''}, "
        f"each separated by '---' below. Extract ALL paquetes and lineas from ALL pages.\n\n"
        if n_pages > 1 else ""
    )
    text_prompt = f"{EXTRACTION_PROMPT}\n\n{page_note}Documento extraído (markdown):\n\n{combined_markdown}"

    # ── Step 2: markdown → JSON ───────────────────────────────────────────
    if step2_provider == "nvidia":
        _status(f"⏳ Paso 2/2 — NVIDIA [{step2_model}]: extrayendo JSON del markdown...")
        raw = _call_nvidia_text_model(text_prompt, step2_model, nvidia_api_key)
    else:
        _status(f"⏳ Paso 2/2 — OpenRouter [{step2_model}]: extrayendo JSON del markdown...")
        client = OpenAI(base_url=config.OPENROUTER_BASE_URL, api_key=openrouter_api_key)
        try:
            response = client.chat.completions.create(
                model=step2_model,
                messages=[{"role": "user", "content": text_prompt}],
                timeout=60.0,
            )
        except Exception as e:
            raise ValueError(f"Error en extracción JSON con OpenRouter (paso 2): {str(e)}")

        if not response or not response.choices:
            raise ValueError("Respuesta vacía de OpenRouter en paso 2.")

        raw = response.choices[0].message.content or ""
        if not raw.strip():
            raise ValueError("OpenRouter retornó respuesta vacía en paso 2.")

    _status("Respuesta recibida. Extrayendo JSON...")
    return parse_llm_response(raw)


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
        text_prompt = f"""{EXTRACTION_PROMPT}

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


def parse_llm_response(text: str) -> dict:
    """
    Extract and validate JSON from LLM response text.
    Raises ValueError if JSON is missing or invalid.
    """
    if not text or not text.strip():
        raise ValueError("Respuesta vacía del LLM.")

    clean = re.sub(r"<think>.*?</think>", "", text, flags=re.DOTALL).strip()
    if not clean:
        raise ValueError("Respuesta del LLM contiene solo etiquetas vacías.")

    code_block = re.search(r"```(?:json)?\s*\n(.*?)```", clean, re.DOTALL)
    json_str = code_block.group(1).strip() if code_block else clean

    if not json_str:
        raise ValueError("No se encontró JSON en la respuesta del LLM.")

    try:
        data = json.loads(json_str)
    except json.JSONDecodeError as e:
        raise ValueError(f"Respuesta del LLM no es JSON válido: {e}\n\nRespuesta recibida:\n{text[:500]}")

    if not isinstance(data, dict):
        raise ValueError(f"LLM retornó tipo {type(data).__name__}, se esperaba dict (objeto JSON).")

    required = ["of_number", "paquetes"]
    for field in required:
        if field not in data:
            raise ValueError(f"Campo requerido '{field}' no encontrado en la respuesta del LLM.")

    if not isinstance(data.get("paquetes"), list):
        raise ValueError(f"Campo 'paquetes' debe ser una lista, se recibió {type(data.get('paquetes')).__name__}.")

    return data
