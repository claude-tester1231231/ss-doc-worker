# ss-doc-worker

Small document extract/rebuild engine used by [SkrivSikkert](https://skrivsikkert.dk).

It does exactly two things per format and nothing else:

- **extract** — document in → ordered text blocks (`{id, text}`) out.
  PDF blocks come from the document's own structure (PyMuPDF `get_text("dict")`)
  with adjacent same-style blocks merged into paragraphs. DOCX paragraphs carry
  inline-formatting markers (`[[B]]…[[/B]]`, `[[I]]…[[/I]]`, `[[BI]]…[[/BI]]`).
- **build** — original document + `{id: replacement-text}` in → rebuilt document
  out. PDF: redact + insert in the original rect (bold/color preserved,
  shrink-to-fit floor 8pt, fail-don't-clip). DOCX: runs rebuilt from markers with
  a hard marker-count contract and a text-never-lost fail-safe.

There is deliberately **no translation logic, no prompts, no API keys and no
business logic** here — the caller owns all of that. This service is the entire
PyMuPDF-touching surface, published under AGPL-3.0 (PyMuPDF's license).

## API

    GET  /healthz
    POST /v1/pdf/extract    multipart: file            -> {blocks:[{id,text}], pages}
    POST /v1/pdf/build      multipart: file, translations={id:text} -> translated.pdf
    POST /v1/docx/extract   multipart: file            -> {blocks:[{id,text}]}
    POST /v1/docx/build     multipart: file, translations={id:text} -> translated.docx

All endpoints require `Authorization: Bearer $WORKER_TOKEN`.
Build enforces a hard 1:1 contract: every extracted block id must get a
non-empty replacement, otherwise HTTP 422 — a document can never lose text
silently.

## Run

    python -m venv .venv && .venv/bin/pip install -r requirements.txt
    WORKER_TOKEN=... .venv/bin/gunicorn -w 2 -b 127.0.0.1:8093 app:app

## License

AGPL-3.0 — see LICENSE.
