---
name: docx2md
description: DOCX/DOC to Markdown Converter - 支持标题识别、表格转换（HTML/Markdown）、自动编号保留、.doc 兼容（LibreOffice）
tools: [Bash(python enhanced_convert.py)]
---

# DOCX/DOC to Markdown Converter

将 Word 文档（.docx / .doc）智能转换为 Markdown 格式，保留标题层级、表格结构、自动编号。

## 功能特性

| 功能 | 状态 | 说明 |
|------|------|------|
| 标题识别 | ✅ | Word 原生样式 + 启发式检测 |
| 表格转换 | ✅ | 简单表格→Markdown，复杂表格→HTML |
| 编号保留 | ✅ | Word 多级自动编号 |
| 格式兼容 | ✅ | .docx 原生解析 + .doc LibreOffice 转换 |
| 批量转换 | ✅ | GUI 多文件批量处理 |
| 图片提取 | ✅ | 内嵌图片导出到文件 |

## 快速开始

### 命令行

```bash
cd "F:\my working space\docx2md"

# 转换 DOCX
python enhanced_convert.py input.docx -o output.md

# 转换 DOC（需 LibreOffice）
python enhanced_convert.py input.doc -o output.md
```

### Python API

```python
from enhanced_convert import detect_and_convert, ProcessingConfig

config = ProcessingConfig(
    enable_table=True,
    fix_chinese_spacing=True,
    remove_page_numbers=True,
)

content, metadata = detect_and_convert("input.docx", "output.md", config)
```

### GUI

```bash
python gui_app_tkinter.py
```

## 配置选项

```python
@dataclass
class ProcessingConfig:
    enable_table: bool = True           # 表格转换
    fix_chinese_spacing: bool = True    # 中文间距修复
    remove_page_numbers: bool = True    # 移除页码
    keep_headers_footers: bool = False  # 保留页眉页脚
    image_enabled: bool = True          # 图片提取
    image_to_base64: bool = False       # 图片Base64内联（默认导出文件）
```

## 注意事项

1. .doc 文件需要安装 LibreOffice
2. 复杂表格（合并单元格）输出为 HTML 格式
3. 图片默认导出到 `images/` 目录

## 文件清单

```
docx2md/
├── enhanced_convert.py       # 核心转换引擎
├── gui_app_tkinter.py       # Tkinter GUI
├── gui_app.spec             # PyInstaller 打包配置
├── requirements.txt          # Python 依赖
├── README.md                 # 快速上手
├── 使用手册.md              # 用户文档
├── 软件开发说明书.md        # 技术文档
└── 打包.bat                 # Windows 打包脚本
```

## 更新日志

### v1.2.0 (2026-08-21)
- 新增 .doc 格式支持（LibreOffice 无头模式转换）
- 格式自动识别（detect_and_convert 统一入口）
- GUI 支持 .doc/.docx 双格式

### v1.1.0 (2026-03-26)
- 修复复杂表格转换（纵向合并单元格）

### v1.0.0 (2024-03-21)
- 初始版本

---

**作者**: 章非
**版本**: 1.2.0
**更新日期**: 2026-08-21