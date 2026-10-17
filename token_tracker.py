"""
token_tracker.py – Lightweight token usage logging for LLM calls.

Every call to log_usage() appends one record to a local JSONL file so usage
survives across Streamlit reruns and app restarts. app.py reads this file to
render the token usage dashboard.
"""
import json
import threading
from datetime import datetime
from pathlib import Path

import pandas as pd

LOG_PATH = Path(__file__).parent / "token_usage_log.jsonl"
_LOCK = threading.Lock()

COLUMNS = ["timestamp", "context", "model", "input_tokens", "output_tokens", "total_tokens"]


def log_usage(context: str, model: str, usage: dict) -> None:
    """Append one usage record.

    `usage` is the OpenAI-style usage dict (prompt_tokens / completion_tokens /
    total_tokens) returned alongside the LLM response; missing keys default to 0.
    """
    usage = usage or {}
    record = {
        "timestamp":     datetime.now().isoformat(timespec="seconds"),
        "context":       context,
        "model":         model,
        "input_tokens":  usage.get("prompt_tokens", 0) or 0,
        "output_tokens": usage.get("completion_tokens", 0) or 0,
        "total_tokens":  usage.get("total_tokens", 0) or 0,
    }
    with _LOCK, open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(json.dumps(record, ensure_ascii=False) + "\n")


def load_usage_log() -> pd.DataFrame:
    """Read the full usage log as a DataFrame (empty DataFrame if no log yet)."""
    if not LOG_PATH.exists():
        return pd.DataFrame(columns=COLUMNS)

    records = []
    with open(LOG_PATH, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                records.append(json.loads(line))
            except json.JSONDecodeError:
                continue

    if not records:
        return pd.DataFrame(columns=COLUMNS)

    df = pd.DataFrame(records)
    df["timestamp"] = pd.to_datetime(df["timestamp"])
    return df


def clear_usage_log() -> None:
    """Delete the usage log file (used by the dashboard's reset button)."""
    with _LOCK:
        if LOG_PATH.exists():
            LOG_PATH.unlink()