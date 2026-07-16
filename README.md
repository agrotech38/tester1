# PSS + COA Generator (merged)

Single Streamlit app that generates **PSS** and **COA** documents from one shared
set of inputs, then merges them into a single downloadable PDF (PSS pages first,
COA pages after).

## Repo layout

```
app.py
MOD PSS.docx
FAR PSS.docx
PH LIPL MOD COA.docx
PH LIPL FAR COA.docx
requirements.txt
packages.txt
```

All four `.docx` templates **must** be committed to the repo root (exact
filenames above) — the app looks them up by name relative to `app.py`.

## How the merge works

`docxcompose` (`Composer`) is used to append the filled COA document onto the
end of the filled PSS document, with a page break in between — this preserves
each document's own formatting, styles, and tables far more reliably than
manually copying XML. The merged `.docx` is the primary output.

## Why `packages.txt`?

The merged `.docx` download works with no extra setup. `packages.txt`
(`libreoffice`) is only needed for the **optional** "Convert merged document
to PDF" button, which shells out to headless LibreOffice
(`soffice --headless --convert-to pdf`) on the already-merged file. Streamlit
Cloud reads `packages.txt` and installs `libreoffice` via `apt-get`
automatically on deploy. If you never need the PDF button, you can delete
`packages.txt` and remove `docx_bytes_to_pdf_bytes` / the PDF button from
`app.py`. If you deploy elsewhere and want the PDF button, make sure
`libreoffice` (or at least `libreoffice-writer` + `libreoffice-core`) is
installed and `soffice` is on `PATH`.

## What changed vs. the two original apps

- One "Choose Type" selector (`MOD` / `FAR`) drives both templates — no more
  separate 001/002 code and MOD/FAR selectors.
- Batch number is entered **once per batch** (tabs 1–4) and reused for both
  PSS's `{{B1}}..{{B4}}` tokens and COA's `BATCH_1..BATCH_4` tokens.
- One shared **Date** field (previously COA only had a date in the Batch 1 tab).
- Generate button fills both templates and merges them into a single `.docx`
  (PSS pages, then a page break, then COA pages) via `docxcompose`. An
  optional button converts that merged file to PDF. Individual `.docx` files
  per document are still available under an expander, in case you need them.

## Local testing

```bash
pip install -r requirements.txt
# install LibreOffice locally, e.g. on Debian/Ubuntu:
sudo apt-get install libreoffice
streamlit run app.py
```
