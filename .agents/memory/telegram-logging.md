---
name: Telegram logging safety
description: Telegram bot transport logging must not expose token-bearing request URLs.
---

Telegram client libraries can log request URLs containing the bot token at INFO level.

**Why:** Workflow output is durable diagnostic data and should never contain credentials.

**How to apply:** Keep `httpx` and `httpcore` transport loggers at WARNING or lower verbosity in Telegram bot entrypoints.