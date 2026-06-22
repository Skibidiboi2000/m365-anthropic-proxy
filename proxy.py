"""Anthropic /v1/messages <-> OpenAI Chat Completions proxy

Translates the full Anthropic Messages API (text, tool_use, tool_result, images,
streaming, system as list) to OpenAI Chat Completions format and sends to 9router.

Required for M365 add-in to work through a gateway.
"""

import json
import os
import uuid
from typing import Any, Dict, List, Optional, Tuple

import httpx
from fastapi import FastAPI, Request, Response
from fastapi.responses import JSONResponse, StreamingResponse, PlainTextResponse
from starlette.middleware.cors import CORSMiddleware

VERSION = "1.0.0"
app = FastAPI()

# CORS: the M365 add-in runs inside pivot.claude.ai
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
    expose_headers=["*"],
)


@app.middleware("http")
async def catch_empty_body(request: Request, call_next):
    """Handle empty body requests (curl --data '' etc) gracefully."""
    if request.method == "POST":
        content_length = request.headers.get("content-length")
        if content_length == "0" or (content_length and int(content_length) < 2):
            return PlainTextResponse("Bad Request: empty body", status_code=400)
    return await call_next(request)

ROUTER_URL = os.environ.get("ROUTER_URL", "http://host.docker.internal:20128/v1")
DEFAULT_MODEL = os.environ.get("DEFAULT_MODEL", "claude-sonnet-4-6")

SUPPORTED_MODELS = [
    "add your model here",
    "add your model here",
    "add your model here",
]

CLIENT = httpx.AsyncClient(timeout=600)


@app.on_event("shutdown")
async def shutdown():
    await CLIENT.aclose()


# ---------------------------------------------------------------
# Anthropic -> OpenAI message conversion
# ---------------------------------------------------------------

def _anth_content_to_oai_text(content: Any) -> str:
    """Flatten Anthropic content blocks -> flat text string for OpenAI."""
    if isinstance(content, str):
        return content
    if isinstance(content, list):
        parts = []
        for block in content:
            if not isinstance(block, dict):
                continue
            t = block.get("type", "")
            if t == "text":
                parts.append(block.get("text", ""))
            elif t == "image":
                # Images: OpenAI can handle as base64 <-> image_url if needed
                # For 9router/deepseek, skip image blocks
                parts.append("[image]")
            elif t == "tool_result":
                tc = block.get("content", "")
                tool_use_id = block.get("tool_use_id", "")
                if isinstance(tc, list):
                    texts = [s.get("text", "") for s in tc if isinstance(s, dict) and s.get("type") == "text"]
                    parts.append(f'[Tool result ({tool_use_id}): {" ".join(texts)}]')
                elif tc:
                    parts.append(f"[Tool result ({tool_use_id}): {tc}]")
            elif t == "tool_use":
                parts.append(f"[Calling tool: {block.get('name', '')}]")
        return "\n".join(parts)
    return str(content)


def _build_oai_messages(body: dict) -> Tuple[List[dict], Optional[List[dict]]]:
    """Build OpenAI messages + tools from an Anthropic request body."""
    messages = body.get("messages", [])
    system = body.get("system", None)
    tools = body.get("tools", None)

    openai_messages: List[dict] = []

    # System prompt: Anthropic accepts string or list of text blocks
    system_text = ""
    if isinstance(system, str):
        system_text = system
    elif isinstance(system, list):
        system_text = "".join(b.get("text", "") for b in system if isinstance(b, dict) and b.get("type") == "text")
    if system_text:
        openai_messages.append({"role": "system", "content": system_text})

    for msg in messages:
        role = msg.get("role", "user")
        content = msg.get("content", "")

        if role == "assistant":
            if isinstance(content, list):
                text_parts = []
                for block in content:
                    if isinstance(block, dict) and block.get("type") == "text":
                        text_parts.append(block.get("text", ""))
                openai_messages.append({"role": "assistant", "content": "\n".join(text_parts) if text_parts else ""})
            else:
                openai_messages.append({"role": "assistant", "content": content or ""})
        elif role == "user":
            openai_messages.append({"role": "user", "content": _anth_content_to_oai_text(content)})
        else:
            openai_messages.append({"role": role, "content": str(content) if content else ""})

    # Server-side Anthropic tools that must be stripped for non-Anthropic backends
    STRIP_TOOLS = {"web_search", "web_fetch", "code_execution", "computer_use", "text_editor"}

    oai_tools = None
    if tools:
        oai_tools = []
        for t in tools:
            name = t.get("name", "")
            if name in STRIP_TOOLS:
                continue
            params = dict(t.get("input_schema", {}))
            if params.get("type") is None:
                params["type"] = "object"
            if "additionalProperties" not in params:
                params["additionalProperties"] = False
            oai_tools.append({
                "type": "function",
                "function": {
                    "name": name,
                    "description": t.get("description", ""),
                    "parameters": params,
                },
            })

    return openai_messages, oai_tools


# ---------------------------------------------------------------
# OpenAI -> Anthropic response conversion
# ---------------------------------------------------------------

def _oai_response_to_anth(result: dict, model: str) -> dict:
    """Convert OpenAI Chat Completions / Responses API to Anthropic format."""
    has_choices = "choices" in result
    has_output = isinstance(result.get("response"), dict) and "output" in result["response"]

    content = ""
    reasoning = None
    tool_calls = None
    finish_reason = "end_turn"
    model_used = model
    usage_data: dict = {}
    resp_id = "msg_unknown"

    if has_output:
        # OpenAI Responses API (what 9router combo returns)
        r = result["response"]
        model_used = r.get("model", model)
        for item in r.get("output", []):
            if item.get("type") == "message":
                for cb in item.get("content", []):
                    ct = cb.get("type", "")
                    if ct == "output_text":
                        content += cb.get("text", "")
                    elif ct == "reasoning":
                        reasoning = (reasoning or "") + cb.get("text", "")
            elif item.get("type") == "function_call":
                if tool_calls is None:
                    tool_calls = []
                args_raw = item.get("arguments", "{}")
                try:
                    args = json.loads(args_raw) if isinstance(args_raw, str) else args_raw
                except json.JSONDecodeError:
                    args = {}
                tool_calls.append({
                    "type": "tool_use",
                    "id": item.get("call_id", f"tu_{uuid.uuid4().hex[:16]}"),
                    "name": item.get("name", ""),
                    "input": args,
                })
        usage_data = r.get("usage", {})
        finish_reason = "end_turn" if r.get("status") == "completed" else "max_tokens"
        resp_id = r.get("id", resp_id)

    elif has_choices:
        choice = result["choices"][0]
        msg = choice.get("message", {})
        content = msg.get("content", "") or ""
        reasoning = msg.get("reasoning_content", None)

        oai_tc = msg.get("tool_calls", None)
        if oai_tc:
            tool_calls = []
            for tc in oai_tc:
                fn = tc.get("function", {})
                args_raw = fn.get("arguments", "{}")
                try:
                    args = json.loads(args_raw) if isinstance(args_raw, str) else args_raw
                except json.JSONDecodeError:
                    args = {}
                tool_calls.append({
                    "type": "tool_use",
                    "id": tc.get("id", f"tu_{uuid.uuid4().hex[:16]}"),
                    "name": fn.get("name", ""),
                    "input": args,
                })

        finish_reason = {
            "stop": "end_turn", "length": "max_tokens", "max_tokens": "max_tokens",
            "tool_calls": "tool_use",
        }.get(choice.get("finish_reason"), "end_turn")
        model_used = result.get("model", model)
        usage_data = result.get("usage", {})

    # Build Anthropic content blocks (reasoning/thinking is intentionally excluded)
    blocks: list = []
    if content:
        blocks.append({"type": "text", "text": content})
    if tool_calls:
        blocks.extend(tool_calls)
    if not blocks:
        blocks.append({"type": "text", "text": ""})

    # Usage mapping
    anth_usage = {
        "input_tokens": usage_data.get("prompt_tokens", 0) or 0,
        "output_tokens": usage_data.get("completion_tokens", 0) or 0,
    }
    p = usage_data.get("prompt_tokens_details") or {}
    if p.get("cached_tokens", 0):
        anth_usage["cache_read_input_tokens"] = p["cached_tokens"]

    if not resp_id.startswith("msg_"):
        resp_id = f"msg_{resp_id[:24]}"

    return {
        "id": resp_id,
        "type": "message",
        "role": "assistant",
        "content": blocks,
        "model": model_used,
        "stop_reason": finish_reason,
        "stop_sequence": None,
        "usage": anth_usage,
    }


# ---------------------------------------------------------------
# /v1/messages — main endpoint
# ---------------------------------------------------------------

@app.post("/v1/messages")
async def messages(request: Request):
    body = await request.json()
    model = body.get("model", DEFAULT_MODEL)
    stream = body.get("stream", False)
    max_tokens = body.get("max_tokens", 4096)
    temperature = body.get("temperature")

    openai_messages, oai_tools = _build_oai_messages(body)

    payload: dict = {
        "model": model,
        "messages": openai_messages,
        "max_tokens": min(max_tokens, 128000),
        "stream": False,
    }
    if temperature is not None:
        payload["temperature"] = temperature
    if oai_tools:
        payload["tools"] = oai_tools
        payload["tool_choice"] = body.get("tool_choice", "auto")

    try:
        resp = await CLIENT.post(
            f"{ROUTER_URL}/chat/completions",
            headers={"Content-Type": "application/json"},
            content=json.dumps(payload),
        )
        raw = resp.text
    except Exception as e:
        return JSONResponse(status_code=502, content={
            "error": {"type": "api_error", "message": f"9router error: {e}"}
        })

    result = _parse_9router_response(raw)
    if result is None:
        return JSONResponse(status_code=502, content={
            "error": {"type": "api_error", "message": f"Bad response from upstream: {raw[:300]}"}
        })

    if "error" in result:
        err = result["error"]
        msg = err.get("message", str(err)) if isinstance(err, dict) else str(err)
        return JSONResponse(status_code=400, content={
            "error": {"type": "api_error", "message": msg}
        })

    anth_response = _oai_response_to_anth(result, model)

    if stream:
        return _make_streaming_response(anth_response)

    return JSONResponse(content=anth_response)


# ---------------------------------------------------------------
# /v1/messages/count_tokens — token counting
# ---------------------------------------------------------------

@app.post("/v1/messages/count_tokens")
async def count_tokens(request: Request):
    """Return a rough token count estimate (the add-in calls this)."""
    body = await request.json()
    # Simple estimation: ~4 chars per token
    total_input = 0
    for msg in body.get("messages", []):
        c = msg.get("content", "")
        if isinstance(c, str):
            total_input += len(c)
        elif isinstance(c, list):
            for block in c:
                if isinstance(block, dict) and block.get("type") == "text":
                    total_input += len(block.get("text", ""))
    system = body.get("system", "") or ""
    if isinstance(system, list):
        system = "".join(b.get("text", "") for b in system if isinstance(b, dict) and b.get("type") == "text")
    total_input += len(system)

    return JSONResponse(content={
        "input_tokens": max(1, total_input // 4),
    })


# ---------------------------------------------------------------
# /v1/models
# ---------------------------------------------------------------

@app.get("/v1/models")
async def list_models():
    return {
        "data": [
            {"id": m, "object": "model", "created": 1677610602, "owned_by": "anthropic"}
            for m in SUPPORTED_MODELS
        ]
    }


# ---------------------------------------------------------------
# Health check for Docker / cloudflared
# ---------------------------------------------------------------

@app.get("/health")
@app.get("/healthz")
async def health():
    return {"status": "ok", "version": VERSION}


# ---------------------------------------------------------------
# CORS preflight
# ---------------------------------------------------------------

@app.options("/v1/messages")
@app.options("/v1/messages/count_tokens")
@app.options("/v1/models")
@app.options("/health")
@app.options("/healthz")
async def options_handler():
    return Response(status_code=200, headers={
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
        "Access-Control-Allow-Headers": "x-api-key, authorization, content-type, anthropic-version, anthropic-beta",
        "Access-Control-Max-Age": "86400",
    })


# ---------------------------------------------------------------
# SSE streaming (non-streaming wrapped as Anthropic SSE events)
# ---------------------------------------------------------------

def _make_streaming_response(anth_resp: dict):
    msg_id = anth_resp["id"]
    model = anth_resp["model"]
    blocks = anth_resp["content"]
    stop_reason = anth_resp.get("stop_reason", "end_turn")
    usage = anth_resp.get("usage", {})

    lines = []

    # message_start
    lines.append(f"event: message_start\ndata: {json.dumps({'type': 'message_start', 'message': {'id': msg_id, 'type': 'message', 'role': 'assistant', 'content': [], 'model': model, 'stop_reason': None, 'stop_sequence': None, 'usage': usage}})}\n\n")

    for idx, block in enumerate(blocks):
        btype = block.get("type", "text")

        if btype == "text":
            text = block.get("text", "")
            lines.append(f"event: content_block_start\ndata: {json.dumps({'type': 'content_block_start', 'index': idx, 'content_block': {'type': 'text', 'text': ''}})}\n\n")
            if text:
                lines.append(f"event: content_block_delta\ndata: {json.dumps({'type': 'content_block_delta', 'index': idx, 'delta': {'type': 'text_delta', 'text': text}})}\n\n")
            lines.append(f"event: content_block_stop\ndata: {json.dumps({'type': 'content_block_stop', 'index': idx})}\n\n")

        elif btype == "tool_use":
            lines.append(f"event: content_block_start\ndata: {json.dumps({'type': 'content_block_start', 'index': idx, 'content_block': {'type': 'tool_use', 'id': block.get('id', ''), 'name': block.get('name', ''), 'input': block.get('input', {})}})}\n\n")
            lines.append(f"event: content_block_stop\ndata: {json.dumps({'type': 'content_block_stop', 'index': idx})}\n\n")

    # message_delta
    lines.append(f"event: message_delta\ndata: {json.dumps({'type': 'message_delta', 'delta': {'stop_reason': stop_reason, 'stop_sequence': None}, 'usage': {'output_tokens': usage.get('output_tokens', 0)}})}\n\n")

    # message_stop
    lines.append(f"event: message_stop\ndata: {json.dumps({'type': 'message_stop'})}\n\n")

    async def gen():
        for line in lines:
            yield line

    return StreamingResponse(gen(), media_type="text/event-stream", headers={
        "Access-Control-Allow-Origin": "*",
        "x-request-id": msg_id,
    })


# ---------------------------------------------------------------
# Parsing 9router response (JSON or SSE)
# ---------------------------------------------------------------

def _parse_9router_response(raw: str) -> Optional[dict]:
    s = raw.strip()
    if not s:
        return None

    try:
        return json.loads(s)
    except json.JSONDecodeError:
        pass

    for line in s.split("\n"):
        line = line.strip()
        if line.startswith("data:") and "[DONE]" not in line:
            try:
                return json.loads(line[5:].strip())
            except json.JSONDecodeError:
                continue

    for line in s.split("\n"):
        line = line.strip()
        if line.startswith("{") or line.startswith("data:{"):
            try:
                js = line[5:].strip() if line.startswith("data:") else line
                return json.loads(js)
            except json.JSONDecodeError:
                continue

    return None
