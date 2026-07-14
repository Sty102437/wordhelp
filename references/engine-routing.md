# Engine Routing Strategy

wordhelp uses a dual-engine architecture. The engine is selected automatically based on workload estimation.

## Engines

| Engine | Technology | Best For |
|--------|-----------|----------|
| python-docx | Python | Light tasks: cell fills, text replacement, simple edits |
| minimax-docx | OpenXML SDK (.NET 8.0) | Heavy tasks: batch generation, complex layouts, professional formatting |

## Workload Estimation

### python-docx (Light Engine)
- Replacing or filling text in a few cells or paragraphs
- Inserting a small number of images or tables
- Simple text replacements without structural changes

### minimax-docx (Heavy Engine)
- Generating large amounts of body text
- Processing multiple complex tables
- Applying professional formatting templates
- CJK typography with strict standards (GB/T 9704)
- Creating academic papers with TOC, headers/footers, footnotes

## Prerequisites

| Engine | Requires |
|--------|----------|
| python-docx | Python 3.10+ with python-docx installed |
| minimax-docx | .NET SDK 8.0+ with minimax-docx skill built |

Run scripts/build-minimax.ps1 once to set up the minimax-docx backend.
