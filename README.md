# CUSTOMATE
## Customs Document Automation

A Python workflow for preparing **Surat Keterangan Barang** supporting customs documentation.

By **Azaria Syahla Fitan Adibah**.

**Upload Word → Review descriptions → Paste/drop images → Download Word**

## Why this project

Repeated copying of part descriptions and photographs makes document preparation tedious. CUSTOMATE extracts item records from Word, provides editable drafts and consolidates image insertion and document export in a local browser interface.

## Features

- Reads changing part numbers/descriptions; item lists are not hardcoded.
- Seven material categories and 20 Indonesian sentence-opening patterns.
- Image copy/paste, drag-and-drop, upload and public source-link import.
- Per-item Google Images links using full PN + Description.
- Save/restore progress using JSON.
- Preserves template font/size for descriptions and Detail Picture labels; known English terms are italicized.
- Centered images, spacing before subsequent items, left-aligned signature and Jakarta export date.
- English navy interface; Titillium Web loads when Google Fonts is reachable.

## Run locally

Requires Python 3.11+.

On Windows, extract the project, run **SETUP_SEKALI.bat** once, then **JALANKAN.bat**. Keep the console open while working. No VS Code is needed for everyday use.

Alternatively:

```bash
python -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt
.venv\Scripts\python app.py
```

On Linux/macOS use `.venv/bin/python`. The app binds only to `127.0.0.1` on an available port.

## Synthetic demonstration

```bash
python demo.py
```

This generates `demo_output/demo_template.docx` with three dummy part numbers and a clearly labeled image placeholder. Upload the template and paste/drop the image to exercise the workflow. No company documents, private identifiers or supplier photographs are bundled.

## Template format

| Column | Purpose |
| --- | --- |
| Item No | Row identifier |
| Part Number | Item reference |
| Description | Part name |
| Function | Existing text retained; blank cells receive drafts |

Each item row must be followed by a `Detail Picture` row. Other structures require adapting the parser. Use a blank Function column to generate fresh wording.

## Kaggle

Import `notebooks/CUSTOMATE_Demo.ipynb`. The notebook is self-contained: it embeds the document functions, creates synthetic records, adds placeholder images, exports Word and validates its structure. Enable Internet only if dependency installation is needed. The local browser app and internet-image search are outside the notebook.

## Tests

```bash
python -m unittest discover -s tests -v
```

Checks cover changing item records, template fonts, English italics, paragraph variants and centered picture export. Structural tests do not guarantee visual pagination. Review exported pages in Word.

## Windows executable

Run **BUAT_EXE.bat** on Windows and share the whole `dist/CUSTOMATE` folder. A prebuilt executable is not included; the Windows build must be tested on Windows.

## Architecture

| File | Responsibility |
| --- | --- |
| `app.py` | Local interface, uploads and public image import |
| `customate_core.py` | Word parsing, drafting, formatting and export |
| `demo.py` | Synthetic demonstration generation |
| `notebooks/CUSTOMATE_Demo.ipynb` | Standalone Kaggle demonstration |
| `tests/test_core.py` | Regression checks |

## Limitations

- This is **rule-based automation, not a trained AI model**. It does not infer material composition from images.
- Materials, functions and resistance claims are draft suggestions requiring human verification. Unknown terms can receive a provisional `metal alloy` category and generic wording. Assemblies/repair kits need particular care.
- English italics use a glossary, not comprehensive language detection.
- Existing Function text is preserved, including its original language.
- Date/signature rules recognize common Jakarta letter patterns; signature text defaults to 11 pt. Other layouts require customization.
- Source sites may return 403 or omit image metadata. Use paste/drop if import fails.
- Images are resized; export requests are limited to 40 MB.
- Automatic Google image retrieval remains experimental and is excluded from this manual portfolio release. CAPTCHA, source restrictions and inconsistent image relevance are unresolved concerns in that development track.

## Data handling

Word files and progress remain in the local browser/server session until downloaded. Closing the app discards unsaved work; a new upload replaces the active document. Use Save progress before closing.

Search links send PN + Description to Google. Source-link import contacts the supplied public URL. Google Fonts may be requested for the interface. No database, cloud document storage or analytics is used.

## Next steps

Improve the part taxonomy, expose uncertain mappings more clearly, make signature/date layout configurable, and evaluate candidate-image retrieval with explicit human review.

See `PUBLISHING_GUIDE.md` for GitHub/Kaggle publishing instructions. No remote repository or notebook has been published by preparing this package. No source-code license has been selected on the author's behalf.
