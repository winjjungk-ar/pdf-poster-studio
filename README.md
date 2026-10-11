# PDF Poster Studio

**PDF Poster Studio** is a Windows desktop app for creating print-ready PDFs from PDF and image files.

It helps you print large maps, posters, game boards, and engineering drawings across normal printer paper such as **A4** or **US Letter**.

## Features

- Open PDFs and images
  - PDF
  - PNG
  - JPG / JPEG
  - BMP
  - WEBP
- Output Area
  - Whole Page
  - Selected Area
- Create Poster PDF
- Create Single Page PDF
- Output Files
  - One PDF
  - Separate PDF per selected page
- Source PDFs
- Print Order
- Auto Select Current Page
- Auto Select All Pages
- Auto Select Page Range
- Go to Page
- Find Text
- A4 / US Letter support
- Poster presets
  - Poster Small / Medium / Large / Huge
  - A3 / A2 / A1 / A0 from A4
  - ANSI B / C / D / E
  - Arch B / C / D / E
- Tiling Preview
- Page rotation
- Project save/load
- Multilingual User Guide
  - English
  - Korean
  - Japanese
  - Simplified Chinese

## Use Cases

- Print game maps across multiple pages
- Print TRPG / tabletop maps
- Create wall posters
- Print engineering drawings
- Make field work print packages
- Combine selected pages or areas from multiple PDFs/images

## Download

Download the latest Windows build from the **Releases** page.

## Quick Start

1. Run `PDF Poster Studio.exe`
2. Click **Open PDFs / Images**
3. Open a PDF or image file
4. Choose **Output Area**
   - Whole Page
   - Selected Area
5. Click **Build Print Order**
6. Choose **Output Files**
   - One PDF
   - Separate PDF per selected page
7. Click **Poster PDF** or **Single Page PDF**

## Notes

- This is a beta release.
- Large PDFs may take time to process.
- Use **Auto Select Page Range** for very large PDFs.
- Image files are converted internally to temporary PDF pages.
- Original PDF/image files are not modified.
- Find Text works only with PDFs that already contain text. Scanned PDFs may require OCR in another tool.

## Build from Source

```bash
pip install -r requirements.txt
python main.py

## Version 0.9.8-dev

This development version focuses on workflow cleanup and Office conversion features.

### Highlights

- Task-first welcome screen
- PDF viewer workflow
- Poster creation workflow
- PDF merge workflow
- PDF split/extract workflow
- PowerPoint, Word, Excel conversion
- Improved PDF viewer rendering quality
- Unified text search with OCR fallback

See `docs/RELEASE_NOTES_0.9.8.md` for details.

