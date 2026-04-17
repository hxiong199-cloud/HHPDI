"""
DocFlow 全局配置管理
支持在线 API 模型 和 本地模型（Ollama 等）
"""

import json
import os
from pathlib import Path

CONFIG_DIR = Path.home() / ".docflow"
CONFIG_FILE = CONFIG_DIR / "config.json"

DEFAULT_CONFIG = {
    "model_mode": "online",
    "online": {
        "provider": "siliconflow",
        "api_key": "",
        "base_url": "https://api.siliconflow.cn/v1",
        "model": "Qwen/Qwen2.5-VL-72B-Instruct",
    },
    "vlm_fallback": {
        "enabled": False,
        "base_url": "",
        "api_key": "",
        "model": "",
    },
    "llm_fallback": {
        "enabled": False,
        "base_url": "",
        "api_key": "",
        "model": "",
    },
    "local": {
        "base_url": "http://localhost:11434/v1",
        "model": "llava:latest",
        "api_key": "ollama",
    },
    "providers": {
        "siliconflow": {
            "name": "硅基流动",
            "base_url": "https://api.siliconflow.cn/v1",
            "models": [
                "Qwen/Qwen2.5-VL-72B-Instruct",
                "Qwen/Qwen2-VL-7B-Instruct",
                "Pro/Qwen/Qwen2-VL-7B-Instruct",
            ],
        },
        "openai": {
            "name": "OpenAI",
            "base_url": "https://api.openai.com/v1",
            "models": ["gpt-4o", "gpt-4o-mini"],
        },
        "azure": {
            "name": "Azure OpenAI",
            "base_url": "",
            "models": ["gpt-4o"],
        },
        "google": {
            "name": "Google AI Studio",
            "base_url": "https://generativelanguage.googleapis.com/v1beta/openai",
            "models": ["gemini-2.0-flash", "gemini-1.5-pro"],
        },
        "custom": {
            "name": "自定义",
            "base_url": "",
            "models": [],
        },
    },
    "parse_options": {
        "render_dpi": 150,
        "extract_tables_as_md": True,
        "extract_formulas": True,
        "add_bbox_comments": True,
    },
    "theme": "light",
    "output_dir": str(Path.home() / "DocFlow_Output"),
}


def load_config() -> dict:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
            return _deep_merge(DEFAULT_CONFIG, saved)
        except Exception:
            pass
    return dict(DEFAULT_CONFIG)


def save_config(cfg: dict):
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def _deep_merge(base: dict, override: dict) -> dict:
    result = dict(base)
    for k, v in override.items():
        if k in result and isinstance(result[k], dict) and isinstance(v, dict):
            result[k] = _deep_merge(result[k], v)
        else:
            result[k] = v
    return result


_config = None


def get_config() -> dict:
    global _config
    if _config is None:
        _config = load_config()
    return _config


def update_config(updates: dict):
    global _config
    cfg = get_config()
    _config = _deep_merge(cfg, updates)
    save_config(_config)
