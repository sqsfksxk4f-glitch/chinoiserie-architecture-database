"""API 配置管理：集中管理所有外部服务（LLM + 专用服务）的地址、模型与密钥。

密钥存入本地 config.local.json（不随代码公开），首次运行生成空配置，
地址与模型名使用各通道的默认值。
"""

import json
import os

CONFIG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.local.json")

# ---- 通道定义（元数据）----

# LLM 类通道：走 OpenAI 兼容协议，用户可自定义 base_url + model + api_key
LLM_CHANNELS = [
    {
        "id": "search_parse",
        "name": "搜索筛选 / 参考文献解析",
        "description": "一键搜索与 DS 解析结果，默认模型 deepseek-chat",
        "default_base_url": "https://api.deepseek.com/v1",
        "default_model": "deepseek-chat",
    },
    {
        "id": "verify_assist",
        "name": "核验工作流 / 聊天助手",
        "description": "DS 核验与聊天助手，默认模型 deepseek-chat",
        "default_base_url": "https://api.deepseek.com/v1",
        "default_model": "deepseek-chat",
    },
    {
        "id": "field_completion",
        "name": "字段补全 / 自动完成",
        "description": "字段自动补全与 G 核验，默认模型 gemini-2.5-flash",
        "default_base_url": "https://yunwu.ai/v1",
        "default_model": "gemini-2.5-flash",
    },
    {
        "id": "keyword_translate",
        "name": "关键词翻译 / 生成",
        "description": "关键词翻译与生成，默认模型 gpt-4o",
        "default_base_url": "https://yunwu.ai/v1",
        "default_model": "gpt-4o",
    },
]

# 专用服务类：固定服务，只有密钥，无法更换模型
SPECIAL_SERVICES = [
    {"id": "google_geocode", "name": "Google 地理编码（坐标解析）"},
    {"id": "serper", "name": "Serper.dev（Google 搜索）"},
    {"id": "baidu", "name": "百度千帆（百度搜索）"},
]


def default_config():
    """构造默认配置：LLM 通道用默认地址/模型 + 空密钥；专用服务空密钥。"""
    cfg = {}
    for ch in LLM_CHANNELS:
        cfg[ch["id"]] = {
            "base_url": ch["default_base_url"],
            "model": ch["default_model"],
            "api_key": "",
        }
    for svc in SPECIAL_SERVICES:
        cfg[svc["id"]] = {"api_key": ""}
    return cfg


def load_config():
    """读取本地配置；文件不存在或损坏时回落到默认配置。"""
    cfg = default_config()
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                saved = json.load(f)
            for key, val in saved.items():
                if key in cfg and isinstance(val, dict):
                    cfg[key].update(val)
        except (json.JSONDecodeError, OSError):
            pass
    return cfg


def save_config(cfg):
    """保存配置到本地文件（覆盖写）。"""
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def get_llm(channel_id):
    """返回某个 LLM 通道的 (base_url, model, api_key)。"""
    cfg = load_config()
    c = cfg.get(channel_id, {})
    return c.get("base_url", ""), c.get("model", ""), c.get("api_key", "")


def get_service_key(service_id):
    """返回某个专用服务的 api_key。"""
    return load_config().get(service_id, {}).get("api_key", "")


def missing_keys():
    """返回尚未配置密钥的通道/服务显示名列表，供缺 key 提示使用。"""
    cfg = load_config()
    missing = []
    for ch in LLM_CHANNELS:
        if not cfg.get(ch["id"], {}).get("api_key", "").strip():
            missing.append(ch["name"])
    for svc in SPECIAL_SERVICES:
        if not cfg.get(svc["id"], {}).get("api_key", "").strip():
            missing.append(svc["name"])
    return missing
