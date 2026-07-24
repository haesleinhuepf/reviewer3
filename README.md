# reviewer3

`reviewer3` is a pip-installable command line reviewer for scientific `.docx` texts.

It reads a document paragraph-by-paragraph, asks an LLM for constructive feedback, and saves a reviewed copy with:

- Word comments (for reviewer questions), and
- tracked-insertion suggestions (author set to `reviewer3`).

The original file is never overwritten. Output is written as `*_rv3.docx`.