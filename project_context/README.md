# Hackathon Evaluation Agents — OpenAI Port

Drop-in replacement that keeps your architecture and features.
- Uses `langchain_openai.ChatOpenAI` with multimodal `gpt-4o-mini` by default.
- Same prompts, JSON parsing, workflow, rate limiting, retries, and timeouts.

## Env
```
OPENAI_API_KEY=sk-...
OPENAI_MODEL=gpt-4o-mini        # optional override
OPENAI_SEED=42                  # optional deterministic seed
LLM_TIMEOUT_S=90
LLM_MAX_RETRIES=2
RATE_LIMIT_RPM=12
BURST_TOKENS=12
TEAM_GLOB=/path/to/ppts/*.pdf   # optional
USE_COMBINED=1                  # use CombinedAgent
MAX_CONCURRENCY=2
```

## Install
```
pip install -U langchain langchain-openai pydantic python-dotenv pypdf python-pptx pillow
```

## Run
```
python orchestrator.py
```