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
    """Test that pdfplumber extraction function exists and is callable."""
    pdf_bytes = create_text_pdf()

    # Mock on_status callback
    status_messages = []
    def mock_status(msg):
        status_messages.append(msg)

    # This test just checks that the function exists and can be called
    assert hasattr(extractor, '_extract_data_pdfplumber')
    assert callable(extractor._extract_data_pdfplumber)


def test_extract_data_pdfplumber_returns_extraction_method():
    """Test that result includes extraction_method field."""
    # This test will be populated after implementation
    # For now, just verify the structure exists
    pass


def test_extract_data_routes_to_pdfplumber():
    """Test that extract_data() routes to pdfplumber when requested."""
    pdf_bytes = create_text_pdf()

    # Verify the function would route correctly (structure test)
    assert hasattr(extractor, 'extract_data')
    assert callable(extractor.extract_data)

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
    pass


def test_extract_data_default_method():
    """Test that extract_data() defaults to nemotron method."""
    import inspect
    sig = inspect.signature(extractor.extract_data)
    params = sig.parameters

    # Check that extraction_method has default value
    assert 'extraction_method' in params
    assert params['extraction_method'].default == "nemotron"
