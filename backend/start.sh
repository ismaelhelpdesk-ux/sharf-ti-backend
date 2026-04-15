#!/bin/sh
# Entrypoint script — resolves PORT env var before starting uvicorn.
# Railway injects $PORT at runtime; fall back to 8000 for local/dev use.

PORT=${PORT:-8000}

exec uvicorn main:app --host 0.0.0.0 --port "$PORT" --workers 2
