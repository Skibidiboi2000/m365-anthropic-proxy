# anthropic-proxy

Translate Anthropic Messages API → OpenAI Chat Completions format.

Allows Microsoft 365 add-ins (Claude for Excel, Word, PowerPoint, Outlook) to work with any OpenAI-compatible backend.

## Why

M365 add-ins use the Anthropic Messages API natively. If you point them at a non-Anthropic gateway, the format mismatch causes errors. This proxy sits in between and translates both request and response formats.

```
M365 add-in → anthropic-proxy → any OpenAI-compatible backend
     (Anthropic format)          (OpenAI Chat Completions)
```

## Quick Start

```bash
docker build -t anthropic-proxy .
docker run -d --name anthropic-proxy --restart unless-stopped \
  -p 4000:8787 \
  -e ROUTER_URL=http://localhost:4000/v1 \
  -e DEFAULT_MODEL=claude-sonnet-4-6 \
  anthropic-proxy:latest
```

## Configuration

| Env var | Default | Description |
|---|---|---|
| `ROUTER_URL` | `http://localhost:4000/v1` | Upstream OpenAI-compatible endpoint |
| `DEFAULT_MODEL` | `claude-sonnet-4-6` | Fallback model name |

## Features

- Translates text, tool_use, tool_result, image blocks
- Strips Anthropic server-side tools (web_search, code_execution) for non-Anthropic backends
- Fixes input_schema type: "object" for OpenAI compatibility
- Handles streaming (SSE → Anthropic event format)
- Token counting endpoint
- CORS enabled

## Endpoints

- `POST /v1/messages` — main chat endpoint (Anthropic format)
- `POST /v1/messages/count_tokens` — token estimation
- `GET /v1/models` — list available models
- `GET /health` — health check

## License

MIT
