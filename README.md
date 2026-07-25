# reviewer3

`reviewer3` is a pip-installable command line reviewer for scientific `.docx` texts.

It reads the full document, asks an LLM for a structured JSON list of feedback entries, and saves a reviewed copy with:

- inline tracked edits (insertions/deletions, author set to `reviewer3`) for local text modifications,
- Word comments for regional clarification questions, and
- a prepended review preface containing:
	- what the document is about,
	- why it is important for science / the scientific community,
	- global feedback (open questions and cross-section coherence),
	- a short summary of local and regional feedback.

Only local modifications are emitted as tracked suggestions. Regional questions are comments. Global feedback appears in the preface.

The original file is never overwritten. Output is written as `*_rv3.docx`.

## Installation

Install the CLI from PyPI:

```bash
pip install reviewer3
```

## Local LLM Setup (Ollama)

`reviewer3` is configured by default to call an OpenAI-compatible server at `http://localhost:11434/v1`, which matches Ollama's local API.

1. Install Ollama from https://ollama.com.
2. Pull the default model used by `reviewer3`:

```bash
ollama pull gpt-oss:20b
```

Hardware note: Running `gpt-oss:20b` effectively requires a dedicated 8 GB GPU.

## Environment Variables

`reviewer3` reads configuration from these environment variables:

- `REVIEWER3_BASE_URL`: OpenAI-compatible API base URL (default: `http://localhost:11434/v1`)
- `REVIEWER3_API_KEY`: API key/token for the server (default: `ollama`)
- `REVIEWER3_MODEL`: Model name to use (default: `gpt-oss:20b`)
- `REVIEWER3_TEMPERATURE`: Sampling temperature (default: `0.2`)

Example (PowerShell):

```powershell
$env:REVIEWER3_BASE_URL = "http://localhost:11434/v1"
$env:REVIEWER3_API_KEY = "ollama"
$env:REVIEWER3_MODEL = "gpt-oss:20b"
$env:REVIEWER3_TEMPERATURE = "0.2"
```

Example (bash/zsh):

```bash
export REVIEWER3_BASE_URL="http://localhost:11434/v1"
export REVIEWER3_API_KEY="ollama"
export REVIEWER3_MODEL="gpt-oss:20b"
export REVIEWER3_TEMPERATURE="0.2"
```

## Connect To An LLM Server

`reviewer3` uses the OpenAI Python client and can connect to any OpenAI-compatible server:

1. Set `REVIEWER3_BASE_URL` to your server's base URL.
2. Set `REVIEWER3_API_KEY` to the token/key expected by that server.
3. Set `REVIEWER3_MODEL` to a model your server exposes.
4. Run `reviewer3` as usual.

Example for a remote server:

```bash
export REVIEWER3_BASE_URL="https://your-llm-server.example.com/v1"
export REVIEWER3_API_KEY="your-token"
export REVIEWER3_MODEL="your-model-name"
```