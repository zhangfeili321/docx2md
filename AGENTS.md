# AGENTS.md - DOCX to Markdown Converter

## Project Overview

**Type**: Python CLI tool + GUI for DOCX to Markdown conversion
**Language**: Python 3.8+
**Location**: `F:\opencode workspace\docx-to-markdown\scripts`

---

## Build/Test Commands

```bash
cd "F:\opencode workspace\docx-to-markdown\scripts"

# Run GUI
python gui_app_tkinter.py

# Run CLI conversion
python enhanced_convert.py input.docx -o output.md

# Test import (verify no syntax errors)
python -c "from enhanced_convert import DocxToMarkdownConverter; print('OK')"
python -c "from gui_app_tkinter import DocxToMarkdownGUI; print('OK')"

# PyInstaller packaging
pip install pyinstaller
python -m PyInstaller gui_app.spec --clean -y
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
from pathlib import Path
from typing import Optional, Dict, Tuple, List, Set
from dataclasses import dataclass, field
from enum import Enum

# Third-party imports
from docx.oxml.ns import qn

# Local imports (if any)
```

### 2. Naming Conventions

| Type | Convention | Example |
|------|------------|---------|
| Classes | PascalCase | `ProcessingConfig`, `TableAnalyzer` |
| Functions | snake_case | `convert_docx_to_markdown()` |
| Private methods | `_` prefix | `_extract_header_footer_text()` |
| Constants | UPPER_SNAKE | `CHINESE_PATTERN`, `DEFAULT_COLS` |
| Variables | snake_case | `output_path`, `pdf_path`, `md_lines` |
| Dataclass fields | snake_case | `enable_formula`, `detect_columns` |
| Enum values | UPPER_SNAKE | `SIMPLE`, `MEDIUM`, `COMPLEX` |

### 3. Configuration Classes

```python
@dataclass
class ProcessingConfig:
    """Processing configuration with sensible defaults."""
    enable_table: bool = True
    enable_formula: bool = True
    fix_chinese_spacing: bool = True
    
    # List fields use field(default_factory=lambda: [...])
    heading_style_keywords: List[str] = field(default_factory=lambda: [
        'heading', 'title', 'toc', 'cover'
    ])
```

### 4. Error Handling Patterns

```python
# Import errors - return tuple with error dict
try:
    from docx import Document
except ImportError:
    return "", {"error": "python-docx not installed. Run: pip install python-docx"}

# Operation errors - clear user-friendly messages
try:
    doc = Document(doc_path)
except Exception as e:
    return "", {"error": f"Failed to load document: {str(e)}"}

# CLI errors - print to stderr, exit with code
if "error" in metadata:
    print(f"Error: {metadata['error']}", file=sys.stderr)
    sys.exit(1)
```

### 5. Static Utility Methods

```python
class TextPostProcessor:
    @staticmethod
    def fix_chinese_spacing(text: str) -> str:
        """Fix spacing between Chinese and Latin characters."""
        text = re.sub(r'([\u4e00-\u9fff]) ([a-zA-Z])', r'\1\2', text)
        text = re.sub(r'([a-zA-Z]) ([\u4e00-\u9fff])', r'\1\2', text)
        return text

    @staticmethod
    def clean_nbsp(text: str) -> str:
        """Remove non-breaking spaces between CJK characters."""
        return re.sub(r'([\u4e00-\u9fff])\u00a0([\u4e00-\u9fff])', r'\1\2', text)
```

### 6. XML Element Access (python-docx)

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

### 7. Document Element Traversal

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

### 8. GUI Threading Pattern

```python
class ConversionThread(threading.Thread):
    def __init__(self, input_path: str, output_path: str, config: dict, callback):
        super().__init__()
        self.callback = callback  # Store callback reference
    
    def run(self):
        try:
            content, metadata = converter.convert(self.input_path, self.output_path)
            self.callback('finished', {'metadata': metadata, 'content': content})
        except Exception as e:
            self.callback('error', str(e))
```

---

## Version

**Last Updated**: 2024-03-21
**Version**: 1.0.0

## Author

**章非**
