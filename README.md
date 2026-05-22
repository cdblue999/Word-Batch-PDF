# Word Batch PDF Converter

Microsoft Word 2013 VBA macro that converts all Word files in a folder to PDF in one batch.

## Features

- Select a folder, convert every Word file inside to PDF
- Supported formats:
  - **Word documents**: .doc, .docx, .docm
  - **Word templates**: .dot, .dotx, .dotm
  - **Word-compatible**: .odt, .rtf
- Files opened **hidden** and **read-only** — no visual flickering, no accidental edits
- PDFs saved in the **same folder** as originals
- PDF bookmarks created from Word headings
- Shows **summary** with success/failure count
- **Opens output folder** in Explorer when done

## Requirements

- **Microsoft Word 2013** (or later)
- Windows

## Installation

1. Open **Word 2013**
2. Press **Alt+F11** to open the VBA editor
3. **File → Import File** → select `WordBatchPDF.bas`
4. Close the VBA editor

## Usage

1. **Alt+F8** to open the Macros dialog
2. Select **WordBatchPDF.BatchWordToPDF**
3. Click **Run**
4. Select a folder containing Word files
5. Wait — progress is shown in the status bar
6. A summary dialog shows results
7. The output folder opens automatically in Explorer

## Sample

```
Before:  C:\Projekty\Sprawozdanie.docx
         C:\Projekty\Umowa.doc
         C:\Projekty\Zalacznik.odt

After:   C:\Projekty\Sprawozdanie.pdf
         C:\Projekty\Umowa.pdf
         C:\Projekty\Zalacznik.pdf
```

## Troubleshooting

| Problem | Solution |
|---------|----------|
| "No Word files found" | Folder is empty or contains unsupported file types |
| File fails to convert | File may be corrupted, password-protected, or from an incompatible format |
| Word security warning | Enable macros in File → Options → Trust Center → Macro Settings |
| Can't find the macro | Check module name: `WordBatchPDF`, macro name: `BatchWordToPDF` |

## Files

| File | What |
|------|------|
| `WordBatchPDF.bas` | VBA source — import into Word via Alt+F11 → File → Import File |
| `README.md` | This file |

## License

MIT License
