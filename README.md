# PDF to Word Pro

A browser-based PDF-to-DOCX converter. Files are processed locally; no document upload endpoint is used. PDF.js and docx are loaded from third-party CDNs, so first load requires internet access.

## Run

Enable GitHub Pages in Settings → Pages → Deploy from a branch → main → / (root). After deployment, open the Pages URL shown by GitHub. Alternatively serve this directory with `python -m http.server 8000` and open http://localhost:8000.

## Modes

- Editable text: extracts text into paragraphs with simplified formatting. It does not reconstruct tables, columns, figures, equations, or original typography.
- Visual copy: renders each PDF page as an image inside Word. Appearance is retained approximately; text is not editable or searchable in Word. Output can be large.

Scanned PDFs have no editable text without OCR. OCR and a Python backend are not included. Neither mode guarantees perfect conversion. Maximum input: 50 MB and 150 pages. Password-protected documents may require removing protection in an authorized PDF editor first.

## Privacy and deployment

Do not commit your PDFs or converted documents to this repository. This source repository is public. Dependency hosts receive ordinary network requests to load libraries, but app code does not send PDF bytes to them. A production deployment should vendor and audit its dependencies and add automated tests.

## Verification

This initial implementation has not been browser-tested. Before relying on it, test an ordinary text PDF, a scan, a multi-column paper, a rotated page, and a PDF containing tables. Compare the DOCX in Microsoft Word against the original. GitHub Pages must be enabled separately; committing these files does not itself publish a site.
