import json
import os
import requests
from pathlib import Path

# ── 配置 ──────────────────────────────────────────────────────────────
# API_KEY 从项目根目录的 .env 文件读取（格式：API_KEY=xxx），不再硬编码。
def _load_env() -> None:
    """从 .env 读取 KEY=VALUE（不覆盖已存在的环境变量）。"""
    env_file = Path(__file__).parent / ".env"
    if not env_file.exists():
        return
    for line in env_file.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        os.environ.setdefault(key.strip(), value.strip())


_load_env()

API_KEY       = os.environ.get("API_KEY", "")
MODEL         = "deepseek-v4-pro"   # 启动时可从可用列表中重新选择
BASE_URL      = "https://api.deepseek.com"
ENDPOINT      = f"{BASE_URL}/chat/completions"
MODELS_URL    = f"{BASE_URL}/models"

HEADERS = {
    "Authorization": f"Bearer {API_KEY}",
    "Content-Type":  "application/json",
}

_MODEL_CACHE = []


def _extract_error_message(response: requests.Response) -> str:
    """Read error.message from JSON response when possible."""
    try:
        return (response.json().get("error") or {}).get("message", "")
    except (ValueError, AttributeError):
        return ""


def _refresh_model_cache() -> list:
    """Refresh and return available model IDs from API."""
    global _MODEL_CACHE
    _MODEL_CACHE = list_models()
    return _MODEL_CACHE


def _ensure_supported_model() -> None:
    """Ensure configured MODEL is available; fallback to first available if needed."""
    global MODEL
    try:
        models = _refresh_model_cache()
    except requests.exceptions.RequestException:
        # Keep configured MODEL when model discovery endpoint is temporarily unavailable.
        return
    if models and MODEL not in models:
        MODEL = models[0]


# ── 查询可用模型 ───────────────────────────────────────────────────────
def list_models() -> list:
    """Return list of available model IDs from the API."""
    resp = requests.get(MODELS_URL, headers=HEADERS, timeout=30)
    resp.raise_for_status()
    data = resp.json()
    # OpenAI-compatible: {"data": [{"id": "...", ...}, ...]}
    if "data" in data:
        return [m["id"] for m in data["data"]]
    # Fallback: flat list or other structure
    return data if isinstance(data, list) else [str(data)]


def pick_model() -> str:
    """Print available models and let user pick one (or type manually)."""
    global MODEL
    print("正在查询可用模型列表…")
    models = []
    try:
        models = list_models()
        print("\n可用模型:")
        for i, m in enumerate(models, 1):
            marker = " ◄ 当前默认" if m == MODEL else ""
            print(f"  {i}. {m}{marker}")
        prompt = f"\n请输入序号选择模型，或直接输入模型名（回车使用默认 '{MODEL}'）: "
    except Exception as e:
        print(f"[模型列表查询失败] {e}")
        print("请手动输入模型名称。")
        prompt = f"模型名称（回车使用默认 '{MODEL}'）: "

    choice = input(prompt).strip()
    if not choice:
        pass  # keep default
    elif choice.isdigit() and models and 1 <= int(choice) <= len(models):
        MODEL = models[int(choice) - 1]
    else:
        MODEL = choice  # manual input

    print(f"使用模型: {MODEL}\n")
    return MODEL


# ── API 调用 ───────────────────────────────────────────────────────────
def call_api(messages: list) -> dict:
    _ensure_supported_model()
    payload = {
        "model":       MODEL,
        "messages":    messages,
        "temperature": 0.1,
        "stream":      False,
    }
    response = requests.post(
        ENDPOINT,
        headers=HEADERS,
        json=payload,
        timeout=60,
    )

    if response.status_code == 400:
        message = _extract_error_message(response).lower()
        if "model" in message and "not available" in message:
            available = _refresh_model_cache()
            fallback = next((m for m in available if m != MODEL), None)
            if fallback:
                # Retry once with a model that is visible to this subscription.
                payload["model"] = fallback
                globals()["MODEL"] = fallback
                response = requests.post(
                    ENDPOINT,
                    headers=HEADERS,
                    json=payload,
                    timeout=60,
                )

    response.raise_for_status()
    return response.json()


# ── 主循环 ────────────────────────────────────────────────────────────
def chat():
    #pick_model()   # 查询并选择模型
    history = []
    print(f"输入 'exit' 退出。\n")

    while True:
        try:
            user_input = input("你: ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\n再见！")
            break

        if not user_input:
            continue
        if user_input.lower() in ("exit", "quit", "退出"):
            print("再见！")
            break

        history.append({"role": "user", "content": user_input})

        try:
            data = call_api(history)

            if "choices" in data:
                reply = data["choices"][0]["message"]["content"]
            elif "data" in data and "messages" in data["data"]:
                msgs = data["data"]["messages"]
                reply = msgs[-1]["content"] if msgs else "(无回复)"
            else:
                reply = json.dumps(data, ensure_ascii=False, indent=2)

            history.append({"role": "assistant", "content": reply})
            print(f"\nGemini: {reply}\n")

        except requests.exceptions.HTTPError as e:
            body = e.response.text if e.response is not None else "(no body)"
            print(f"\n[HTTP错误] {e.response.status_code}: {body}\n")
        except requests.exceptions.RequestException as e:
            print(f"\n[请求错误] {e}\n")
        except (KeyError, IndexError, json.JSONDecodeError) as e:
            print(f"\n[解析错误] {e}\n")



if __name__ == "__main__":
    chat()