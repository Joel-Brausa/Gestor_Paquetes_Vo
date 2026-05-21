import json
import os

BASE_DIR = os.path.dirname(__file__)

OPENROUTER_BASE_URL = "https://openrouter.ai/api/v1"
DEFAULT_MODEL = "nvidia/nemotron-3-nano-omni-30b-a3b-reasoning:free"

NVIDIA_BASE_URL = "https://integrate.api.nvidia.com/v1"
NVIDIA_PARSE_URL = "https://integrate.api.nvidia.com/v1/chat/completions"
NVIDIA_PARSE_MODEL = "nvidia/nemotron-parse"
DEFAULT_NVIDIA_TEXT_MODEL = "nvidia/llama-3.3-nemotron-super-49b-v1.5"

MODELS_FILE = os.path.join(BASE_DIR, "models", "models.json")
NVIDIA_TEXT_MODELS_FILE = os.path.join(BASE_DIR, "models", "nvidia_models.json")


def _get_secret(key: str) -> str:
    """Read a secret from env var (local dev) or st.secrets (Streamlit Cloud)."""
    value = os.environ.get(key, "")
    if value:
        return value.strip()
    try:
        import streamlit as st
        value = st.secrets.get(key, "")
        if value:
            return value.strip()
    except Exception:
        pass
    return ""


def load_models() -> list[str]:
    if os.path.exists(MODELS_FILE):
        try:
            with open(MODELS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return [DEFAULT_MODEL]
    return [DEFAULT_MODEL]


def save_models(models: list[str]) -> None:
    os.makedirs(os.path.dirname(MODELS_FILE), exist_ok=True)
    with open(MODELS_FILE, "w", encoding="utf-8") as f:
        json.dump(models, f, ensure_ascii=False, indent=2)


def load_nvidia_text_models() -> list[str]:
    if os.path.exists(NVIDIA_TEXT_MODELS_FILE):
        try:
            with open(NVIDIA_TEXT_MODELS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return [DEFAULT_NVIDIA_TEXT_MODEL]
    return [DEFAULT_NVIDIA_TEXT_MODEL]


def save_nvidia_text_models(models: list[str]) -> None:
    os.makedirs(os.path.dirname(NVIDIA_TEXT_MODELS_FILE), exist_ok=True)
    with open(NVIDIA_TEXT_MODELS_FILE, "w", encoding="utf-8") as f:
        json.dump(models, f, ensure_ascii=False, indent=2)


def load_api_key() -> str:
    return _get_secret("OPENROUTER_API_KEY")


def load_nvidia_api_key() -> str:
    return _get_secret("NVIDIA_API_KEY") or _get_secret("NVIDIA_PAI_KEY")
