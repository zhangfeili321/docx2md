# AGENTS.md - DOCX/DOC to Markdown Converter

## Project Overview

**Type**: Python CLI tool + GUI for DOCX/DOC to Markdown conversion
**Language**: Python 3.8+
**Location**: `F:\my working space\docx2md`

---

## Build/Test Commands

```bash
cd "F:\my working space\docx2md"

# Run GUI
python gui_app_tkinter.py

# Run CLI (DOCX)
python enhanced_convert.py input.docx -o output.md

# Run CLI (DOC, requires LibreOffice)
python enhanced_convert.py input.doc -o output.md

# Test import
python -c "from enhanced_convert import detect_and_convert; print('OK')"

# PyInstaller packaging
python -m PyInstaller gui_app.spec --clean -y
```

---

## Architecture

### Entry Points

```
enhanced_convert.py
├── detect_and_convert()     # Unified entry: auto-detect .doc/.docx
│   ├── .docx → convert_docx_to_markdown()
│   └── .doc  → _convert_doc_to_docx_via_libreoffice() → convert_docx_to_markdown()
├── _find_libreoffice()      # Locate LibreOffice installation
├── DocxToMarkdownConverter  # Core conversion engine
│   ├── NumberingExtractor   # Word auto-numbering
│   ├── HeadingDetector      # Style-based heading detection
│   ├── TableConverter       # Table → Markdown/HTML
│   ├── ImageExtractor       # Image extraction
│   └── TextPostProcessor    # Chinese spacing, empty line cleanup
└── main()                   # CLI entry

gui_app_tkinter.py
├── ConversionThread         # Single-file conversion (background)
├── BatchConversionThread    # Multi-file batch conversion
└── DocxToMarkdownGUI        # Tkinter GUI
```

---

## Code Style Guidelines

### 1. Imports (Standard Library First)

```python
#!/usr/bin/env python3
"""
Module docstring describing purpose.
"""

import os
import sys
import re
import json
import argparse
import tempfile
import shutil
import subprocess
from pathlib import Path
from typing import Optional, Dict, Tuple, List, Set
from dataclasses import dataclass, field
from enum import Enum

# Third-party imports
from lxml import etree
from docx import Document
from docx.oxml.ns import qn
```

### 2. Naming Conventions

| Type | Convention | Example |
|------|------------|---------|
| Classes | PascalCase | `ProcessingConfig`, `TableAnalyzer` |
| Functions | snake_case | `detect_and_convert()` |
| Private functions | `_` prefix | `_find_libreoffice()` |
| Constants | UPPER_SNAKE | `CHINESE_PATTERN` |
| Variables | snake_case | `output_path`, `md_lines` |

### 3. Error Handling Patterns

```python
# Import errors - return tuple with error dict
try:
    from docx import Document
except ImportError:
    return "", {"error": "python-docx not installed. Run: pip install python-docx"}

# DOC conversion - raise RuntimeError with user-friendly message
soffice = _find_libreoffice()
if not soffice:
    raise RuntimeError(
        "未找到 LibreOffice。\n"
        ".doc 文件需要 LibreOffice 进行转换。\n"
        "请从 https://www.libreoffice.org/download/ 下载安装。"
    )

# CLI errors - print to stderr, exit with code
if "error" in metadata:
    print(f"错误: {metadata['error']}", file=sys.stderr)
    sys.exit(1)
```

### 4. Document Element Traversal

⚠️ **IMPORTANT**: Use `doc.element.body` children for correct element ordering:

```python
body = doc.element.body
for child in body:
    tag = child.tag.split('}')[1] if '}' in child.tag else child.tag
    if tag == 'p':      # Paragraph
        process_paragraph(child)
    elif tag == 'tbl':   # Table
        process_table(child)
```

**DO NOT** iterate `doc.paragraphs` and `doc.tables` separately - this loses ordering.

### 5. XML Element Access (python-docx)

⚠️ **CRITICAL**: XML attribute names differ from python-docx properties:

```python
# ✅ Correct - gridSpan access
if tc.grid_span > 1:

# ✅ Correct - vMerge check (must check XML element)
vmerge_elem = tc.find(qn('w:vMerge'))
if vmerge_elem is not None:
    # vMerge="restart" means merge start
    # vMerge exists but no value means continuation
```

### 6. GUI Threading Pattern

```python
class ConversionThread(threading.Thread):
    def run(self):
        try:
            from enhanced_convert import detect_and_convert, ProcessingConfig
            config = ProcessingConfig(...)
            content, metadata = detect_and_convert(self.input_path, self.output_path, config)
            self.callback('finished', {'metadata': metadata, 'content': content})
        except Exception as e:
            self.callback('error', str(e))
```

---

## Version

**Last Updated**: 2026-08-21
**Version**: 1.2.0

## Author

**章非**