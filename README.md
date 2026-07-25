# reviewer3

`reviewer3` is a pip-installable command line reviewer for scientific `.docx` manuscripts. 
It is intended for personal use, if you need feedback about a document you wrote yourself.
Per default it uses a locally installed, privacy-preserving large language model (LLM) via [ollama](https://ollama.com). Optionally, you can configure it to use LLMs on remote servers.
In case you use it for reviewing manuscript of others, make sure to follow best-practices for preseving confidentiality! Do not submit manuscripts to remote servers if you`re not sure what service providers do with it. Also check institutional guidelines on the use of generative artificial intelligence for reviewing, e.g. [guidelines provided by the German Research Foundation](https://www.dfg.de/en/news/news-topics/announcements-proposals/2026/ifr-26-18). In general, the human reviewer is responsible for submitting a reviewed document, no matter what [artificial intelligence] tools they used for reviewing. Keep in mind that LLMs do mistakes and so does `reviewer3`.

![](docs/screenshot.png)

`reviewer3` reads the document, potentially splits it in paragraphs, and asks an LLM for a structured feedback:
- inline tracked edits (insertions/deletions) for text suggestions,
- comments for clarification questions, and
- a prepended review preface containing:
	- what the document is about,
	- why it is important for science / the scientific community,
	- global feedback (open questions and cross-section coherence),
	- a short summary of the feedback.

The original file is never overwritten. Output is written as `*_rv3.docx`.

## What `reviewer3` is not

`reviewer3` is not intended for judging manuscripts. Many language models are not capable of drawing educated decisions about manuscript quality, scientific novelty or if manuscripts should be accepted or not for publishing. Instead, `reviewer3` is intended for providing constructive feedback, allowing to improve a manuscript, not matter how innovative or relevant it might be for the scientific society.

`reviewer3` is not a replacement for human reviewing. It can assist spotting obvious mistakes, or open questions. However, as it splits long documents in multiple blocks for reviewing, it may miss the big picture and may provide misleading suggestions. Review its suggestions and comments carfully, in particular if you use `reviewer3` for preparing feedback for people who asked _you_ to review their work.

`reviewer3` cannot give feedback about figures yet. It only processes the text in manuscript files.

## Installation

Install the CLI from PyPI:

```bash
pip install reviewer3
```

## Local LLM Setup (Ollama)

`reviewer3` is configured by default to call an OpenAI-compatible server at `http://localhost:11434/v1`, which matches Ollama's local API.

1. Install Ollama from https://ollama.com.
2. Pull the default model [gpt-oss:20b]() used by `reviewer3`:

```bash
ollama pull gpt-oss:20b
```

Hardware note: Running `gpt-oss:20b` effectively requires a GPU with at least 8 GB of memory.

## Environment Variables

`reviewer3` optionally reads configuration from environment variables. You can set those if you would like to use `reviewer3` with a remote server. Keep confidentiality in mind when configuring remote LLM servers. 

- `REVIEWER3_BASE_URL`: OpenAI-compatible API base URL (default: `http://localhost:11434/v1`)
- `REVIEWER3_API_KEY`: API key/token for the server (default: `ollama`)
- `REVIEWER3_MODEL`: Model name to use (default: `gpt-oss:20b`)
- `REVIEWER3_TEMPERATURE`: Sampling temperature (default: `0.2`)
- `REVIEWER3_CHUNK_SIZE`: Maximum characters sent per review request (default: `12000`)

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

Example for a remote server:

```bash
export REVIEWER3_BASE_URL="https://your-llm-server.example.com/v1"
export REVIEWER3_API_KEY="your-token"
export REVIEWER3_MODEL="your-model-name"
```

## Contributing

Contributions are welcome! Consider opening an issue first to discuss potential improvements, before preparing a pull-request. Note that `reviewer3` was largely vibe-coded using Github Copilot and OpenAI Codex integrations in Visual Studio Code. Consider using similar tools for modifying the code.


