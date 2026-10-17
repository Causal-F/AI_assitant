"""
agent.py  –  Test Report Comparison Agent

Flow:
  system prompt  = full text from test_report_prompt_v15.docx
  user message   = TC content + TR content + user focus instructions
  LLM call       = chat.call_api()
  output         = structured result dict
"""
import io
import json
import re
from pathlib import Path

import requests

import chat
import token_tracker
from chat import call_api


def _describe_http_error(e: Exception) -> str:
    """Return a user-friendly message for HTTPError (parsing the API's JSON body), else str(e)."""
    response = getattr(e, "response", None)
    if not isinstance(e, requests.exceptions.HTTPError) or response is None:
        return str(e)

    body = response.text.strip()
    message = ""
    try:
        message = (response.json().get("error") or {}).get("message", "")
    except (ValueError, AttributeError):
        pass

    if "sensitive word" in message.lower():
        return (
            "内容安全过滤拦截 / Content blocked by safety filter：LLM 服务检测到输入内容中含有敏感词，"
            "已拒绝处理。请检查 Test Case / Test Report 中是否混入了异常文字、编号或日期等敏感信息，"
            "修改后重试。"
            f"\n(API message: {message})"
        )
    if "exceeded" in message.lower() and "token" in message.lower():
        return (
            "本月 token 配额已用尽 / Monthly token quota exceeded：当前订阅套餐的月度用量已超限，"
            "LLM 服务拒绝了本次请求。请联系管理员升级套餐，或等待下个计费周期配额重置后重试。"
            f"\n(API message: {message})"
        )
    if message:
        return f"{response.status_code} 错误: {message}\n(原始返回: {body[:500]})"
    return f"{e} | 响应内容: {body[:1000]}" if body else str(e)


# ── Prompt Loader ──────────────────────────────────────────────────────────────

def load_prompt(prompt_path: str) -> str:
    """Load comparison prompt from .docx – extracts paragraphs AND table cells."""
    try:
        from docx import Document
        doc = Document(prompt_path)
        parts = []
        for p in doc.paragraphs:
            if p.text.strip():
                parts.append(p.text.strip())
        for table in doc.tables:
            for row in table.rows:
                row_text = "  |  ".join(c.text.strip() for c in row.cells if c.text.strip())
                if row_text:
                    parts.append(row_text)
        return "\n".join(parts)
    except Exception as e:
        return f"[Could not load prompt from {prompt_path}: {e}]"


# ── File Text Extraction ───────────────────────────────────────────────────────

def extract_text(file_bytes: bytes, filename: str) -> str:
    """Extract readable text from uploaded file bytes."""
    suffix = Path(filename).suffix.lower()
    try:
        if suffix in (".xlsx", ".xls"):
            import openpyxl
            wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
            lines = []
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                lines.append(f"[Sheet: {sheet_name}]")
                for row in ws.iter_rows(values_only=True):
                    cells = "\t".join("" if c is None else str(c) for c in row).strip()
                    if cells:
                        lines.append(cells)
            return "\n".join(lines)

        elif suffix == ".docx":
            from docx import Document
            doc = Document(io.BytesIO(file_bytes))
            parts = []
            for p in doc.paragraphs:
                if p.text.strip():
                    parts.append(p.text.strip())
            for table in doc.tables:
                for row in table.rows:
                    row_text = "\t".join(c.text.strip() for c in row.cells)
                    if row_text.strip():
                        parts.append(row_text)
            return "\n".join(parts)

        elif suffix == ".pdf":
            from pypdf import PdfReader
            reader = PdfReader(io.BytesIO(file_bytes))
            pages = []
            for i, page in enumerate(reader.pages, 1):
                text = page.extract_text() or ""
                if text.strip():
                    pages.append(f"[Page {i}]\n{text.strip()}")
            return "\n\n".join(pages) if pages else "[PDF contains no extractable text]"

        elif suffix in (".txt", ".csv"):
            return file_bytes.decode("utf-8", errors="replace")

        else:
            return f"[Unsupported file type: {suffix}]"

    except Exception as e:
        return f"[Error reading {filename}: {e}]"


# ── Message Builder ──────────────────────────────────────────────────────────

def _build_messages(
    prompt_text: str,
    tc_text: str, tc_name: str,
    tr_text: str, tr_name: str,
    chat_history: list = None,
) -> list:
    system_content = (
        prompt_text
        + "\n\n"
        + "请严格按照上述「7 审查结果输出格式」等章节的要求，以如下 JSON 格式返回结果，不要输出 JSON 以外的任何内容。\n"
        + "description、requirement_text、report_text、recommendations 字段必须同时包含中文和英文，格式为 \"English / 中文\"。\n"
        + "category 必须从下列九类中选择其一（原文照抄，不要翻译）：版本一致性问题 | 基础信息问题 | 仪器设备校验问题 | "
        + "测试规程不一致问题 | 测试判定不一致问题 | 测试内容执行与测试规程不一致 | 测试原始数据问题 | 结论不符合测试内容问题 | 笔误 | 图片与描述不符\n"
        + "severity 必须从下列四级中选择其一（原文照抄，不要翻译）：严重 | 一般 | 轻微 | 需人工复核\n"
        + "overall 必须从下列三种结论中选择其一（原文照抄，不要翻译）：符合 | 不符合 | 需人工复核后确认\n"
        + '{\n'
        + '  "overall": "符合 | 不符合 | 需人工复核后确认",\n'
        + '  "summary": "One-line English conclusion / 一句话中文总结",\n'
        + '  "issues": [\n'
        + '    {\n'
        + '      "no": 1,\n'
        + '      "description": "Problem description incl. its location in the report, in English / 问题描述及所在报告位置",\n'
        + '      "requirement_text": "Quoted original wording from our test requirement / 我方测试要求原文引用",\n'
        + '      "report_text": "Quoted original wording from the supplier test report / 供应商测试报告原文引用",\n'
        + '      "category": "one of the nine category names above",\n'
        + '      "severity": "严重 | 一般 | 轻微 | 需人工复核"\n'
        + '    }\n'
        + '  ],\n'
        + '  "manual_review_images": [\n'
        + '    {\n'
        + '      "location": "Where in the report / 报告中的位置",\n'
        + '      "requirement_text": "Corresponding requirement or judgment criterion / 对应的规程条款或判定要求",\n'
        + '      "report_text": "Text description of the image in the report, if any / 报告中对该图片的文字描述",\n'
        + '      "reason": "Why the image content cannot be reliably judged / 无法识别的具体原因"\n'
        + '    }\n'
        + '  ],\n'
        + '  "recommendations": "Suggested remediation for non-conforming items, or manual confirmation needed / 针对不符合项的整改建议"\n'
        + '}'
    )
    # 将完整对话历史（含助手追问与工程师回复）注入分析上下文
    focus = ""
    if chat_history:
        turns = [
            m for m in chat_history
            if m.get("role") in ("user", "assistant") and m.get("content", "").strip()
        ]
        if turns:
            role_map = {"user": "工程师", "assistant": "助手"}
            lines = [f"[{role_map.get(m['role'], m['role'])}] {m['content']}" for m in turns]
            focus = "\n\n【分析前对话上下文】\n" + "\n".join(lines)
    user_content = (
        f"【Test Case 文件名】{tc_name}\n"
        f"【Test Case 内容】\n{tc_text}\n\n"
        f"{'─' * 60}\n\n"
        f"【Test Report 文件名】{tr_name}\n"
        f"【Test Report 内容】\n{tr_text}"
        f"{focus}"
    )
    return [
        {"role": "system", "content": system_content},
        {"role": "user",   "content": user_content},
    ]


# ── Response Parser ────────────────────────────────────────────────────────────

def _parse_response(raw_text: str) -> dict:
    """Extract structured dict from LLM response text."""
    match = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", raw_text, re.DOTALL)
    candidate = match.group(1) if match else raw_text
    for text in (candidate, raw_text):
        brace = re.search(r"\{.*\}", text, re.DOTALL)
        if brace:
            try:
                data = json.loads(brace.group())
                return {
                    "overall":              data.get("overall", "N/A"),
                    "summary":              data.get("summary", ""),
                    "issues":               data.get("issues", []),
                    "manual_review_images": data.get("manual_review_images", []),
                    "recommendations":      data.get("recommendations", ""),
                    # backward compat: keep details if LLM still returns it
                    "details":              data.get("details", ""),
                }
            except json.JSONDecodeError:
                continue
    return {
        "overall": "需人工复核后确认", "summary": raw_text[:200],
        "issues": [], "manual_review_images": [], "recommendations": "", "details": raw_text,
    }


# ── Core Comparison Function ───────────────────────────────────────────────────

def compare(
    tc_bytes: bytes,
    tc_name: str,
    tr_bytes: bytes,
    tr_name: str,
    prompt_text: str,
    chat_history: list = None,
) -> dict:
    """
    Compare a Test Case against a Test Report via the LLM.
    Returns dict: overall, summary, issues, manual_review_images, recommendations, details
    """
    tc_text = extract_text(tc_bytes, tc_name)
    tr_text = extract_text(tr_bytes, tr_name)
    messages = _build_messages(
        prompt_text, tc_text, tc_name, tr_text, tr_name, chat_history
    )
    try:
        response = call_api(messages)
        token_tracker.log_usage(f"Compare: {tr_name}", chat.MODEL, response.get("usage"))
        if "choices" in response:
            raw_text = response["choices"][0]["message"]["content"]
        elif "data" in response and "messages" in response["data"]:
            msgs = response["data"]["messages"]
            raw_text = msgs[-1]["content"] if msgs else ""
        else:
            raw_text = json.dumps(response, ensure_ascii=False)
        return _parse_response(raw_text)
    except Exception as e:
        detail = _describe_http_error(e)
        return {
            "overall": "Error", "summary": f"LLM 调用失败: {detail}",
            "issues": [], "manual_review_images": [], "recommendations": "", "details": detail,
        }


# ── Pre-analysis Chat ──────────────────────────────────────────────────────────

CHAT_SYSTEM_PROMPT = (
    "你是一个专业的测试报告分析助手。"
    "用户即将上传测试用例（Test Case）和测试报告（Test Report）进行对比分析。"
    "你的任务是通过简短对话帮助用户明确分析方向和关注重点。\n\n"
    "指导原则：\n"
    "- 对于信息明确的指令，简短确认并告知可以点击 ▶ Run Comparison 开始分析\n"
    "- 对于模糊或宽泛的描述，追问 1~2 个最关键的细节，不要一次提多个问题\n"
    "- 保持对话简洁高效，避免冗余解释\n"
    "- 当信息足够时，明确告知用户：信息已足够，请点击 ▶ Run Comparison 开始分析\n\n"
    "请始终用中文回复，保持专业简洁。"
)


def chat_reply(history: list) -> str:
    """
    Generate an LLM reply for the pre-analysis conversational phase.
    history: list of {"role": "user"|"assistant", "content": "..."}
    """
    messages = [{"role": "system", "content": CHAT_SYSTEM_PROMPT}] + history
    try:
        response = call_api(messages)
        token_tracker.log_usage("Pre-analysis chat", chat.MODEL, response.get("usage"))
        if "choices" in response:
            return response["choices"][0]["message"]["content"]
        elif "data" in response and "messages" in response["data"]:
            msgs = response["data"]["messages"]
            return msgs[-1]["content"] if msgs else "（无回复）"
        return "（无法解析回复）"
    except Exception as e:
        return f"（对话服务暂时不可用: {_describe_http_error(e)}）"