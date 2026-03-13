#!/bin/bash
export ANTHROPIC_API_KEY="your-key-here"
uvicorn proxy:app --host 0.0.0.0 --port 8000
