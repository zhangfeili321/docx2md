---
name: Markdown Processor
description: 智能 PDF 转 Markdown 工具 - 支持多栏布局检测、数学公式识别、表格转换、图片提取和自动OCR
tools: [Bash(python3 markdown_processor/advanced_processor.py)]
---

# Markdown Processor - 智能 PDF 转 Markdown

将 PDF 文档智能转换为 Markdown 格式，支持扫描版PDF自动OCR、多栏布局重建、数学公式 LaTeX 转换、表格提取和图片嵌入。

## 功能特性

| 功能 | 状态 | 说明 |
|------|------|------|
| 📄 **多栏布局检测** | ✅ | 自动识别单栏/双栏/三栏布局，重建正确阅读顺序 |
| 🧮 **数学公式识别** | ✅ | 识别数学符号并转换为 LaTeX 格式 |
| 📊 **表格提取** | ✅ | 支持简单表格（Markdown）和复杂表格（HTML） |
| 🖼️ **图片提取** | ✅ | 自动提取图片并生成 Markdown 引用 |
| 🔍 **自动OCR** | ✅ | 自动检测扫描版PDF并进行OCR识别 |

## 快速开始

### 命令行使用

```bash
cd /root/.openclaw/workspace/markdown_processor

# 基本使用（自动OCR已启用）
python3 advanced_processor.py input.pdf

# 指定输出目录
python3 advanced_processor.py input.pdf output_folder
```

### Python API

```python
import sys
sys.path.insert(0, '/root/.openclaw/workspace/markdown_processor')

from advanced_processor import AdvancedMarkdownProcessor, ProcessingConfig

# 配置处理选项
config = ProcessingConfig(
    detect_columns=True,    # 多栏检测
    detect_formulas=True,   # 公式识别
    detect_tables=True,     # 表格检测
    extract_images=True,    # 图片提取
    auto_ocr=True,          # 自动OCR扫描版PDF
    ocr_engine='auto'       # OCR引擎: auto, paddle, tesseract
)

# 创建处理器
processor = AdvancedMarkdownProcessor(config)

# 处理 PDF（自动检测扫描版并进行OCR）
output_path = processor.process_pdf("input.pdf", "output_dir")
print(f"已生成: {output_path}")
```

## 配置选项

### ProcessingConfig 参数

```python
@dataclass
class ProcessingConfig:
    detect_columns: bool = True       # 多栏布局检测
    detect_formulas: bool = True      # 数学公式识别
    detect_tables: bool = True        # 表格检测
    extract_images: bool = True       # 图片提取
    auto_ocr: bool = True             # 自动OCR扫描版PDF
    ocr_engine: str = 'auto'          # OCR引擎: auto, paddle, tesseract
    table_format: str = 'auto'        # 表格格式: auto, markdown, html
    image_format: str = 'png'         # 图片保存格式
    image_base64: bool = False        # 图片转为 Base64 内联
    image_dir: str = 'images'         # 图片保存目录
    heading_detection: bool = True    # 标题层级检测
```

## 自动OCR功能

工具会自动检测PDF是否为扫描版（图片型），如果是，则自动进行OCR识别：

```python
# 启用自动OCR（默认已启用）
config = ProcessingConfig(auto_ocr=True)
processor = AdvancedMarkdownProcessor(config)

# 处理时会自动检测并OCR
output = processor.process_pdf("扫描版.pdf")
```

OCR引擎选择优先级：
1. **PaddleOCR**（首选）- 中文识别效果更好
2. **Tesseract**（备用）- 开源OCR引擎

## 功能详解

### 1. 多栏布局检测

自动检测 PDF 的单栏、双栏或三栏布局，并根据检测到的栏数重建正确的阅读顺序。

### 2. 公式识别

识别数学公式并转换为 LaTeX 格式，支持：

- 希腊字母（α, β, γ...）
- 上标/下标（x², H₂O）
- 积分、求和符号
- 常见公式（E=mc², F=ma）
- 分数表达式

### 3. 表格处理

智能识别表格结构并转换：

```python
from table_processor import TableProcessor

processor = TableProcessor()
tables = processor.process_page(page)
for t in tables:
    print(f"表格: {t['structure'].rows}x{t['structure'].cols}")
    print(t['markdown'])  # Markdown 或 HTML 格式
```

### 4. 完整处理流程

```python
from advanced_processor import AdvancedMarkdownProcessor, ProcessingConfig

# 启用所有功能
config = ProcessingConfig(
    detect_columns=True,
    detect_formulas=True,
    detect_tables=True,
    extract_images=True,
    auto_ocr=True  # 关键：自动OCR扫描版
)

processor = AdvancedMarkdownProcessor(config)

# 处理单页
result = processor.process_page(page, page_num=0, output_dir="output")
print(f"检测到 {result['num_columns']} 栏")
print(f"识别到 {len(result['formulas'])} 个公式")
print(f"检测到 {len(result['tables'])} 个表格")

# 处理整个 PDF
output_path = processor.process_pdf("input.pdf", "output")
```

## 输出结构

```
output_dir/
├── input.md                    # 生成的 Markdown 文件
├── processing_report.md        # 处理报告
└── images/                     # 提取的图片
    ├── page_1_img_1.png
    └── page_2_img_1.png
```

## 测试

```bash
cd /root/.openclaw/workspace/markdown_processor

# 测试自动OCR功能
python3 test_auto_ocr.py

# 运行集成测试
python3 integrated_test.py

# 测试特定 PDF
python3 integrated_test.py /path/to/your.pdf
```

## 注意事项

1. **依赖**: 需要安装 `PyMuPDF`, `tesseract-ocr`, `tesseract-ocr-chi-sim`
2. **PaddleOCR**: 可选安装，中文识别效果更好：`pip install paddlepaddle paddleocr`
3. **公式识别**: 复杂手写公式可能识别不准确
4. **扫描版PDF**: 自动OCR需要一定时间，取决于文档页数

## 文件清单

```
markdown_processor/
├── advanced_processor.py       # 主处理器（推荐使用）
├── ocr_engine.py               # OCR引擎（自动识别扫描版）
├── formula_recognizer.py       # 公式识别
├── table_processor.py          # 表格处理
├── column_layout_detector.py   # 多栏检测
├── integrated_test.py          # 集成测试
├── test_auto_ocr.py            # 自动OCR测试
├── test_demo.py                # 单元测试
├── PROGRESS.md                 # 开发进度
├── SKILL.md                    # 本文档
└── README.md                   # 快速上手指南
```

## 更新日志

### 2026-03-13
- ✅ 完成公式识别增强
- ✅ 完成表格处理整合
- ✅ 完成高级处理器整合
- ✅ 新增自动OCR功能（扫描版PDF自动识别）
- ✅ 通过集成测试

---

**作者**: OpenClaw  
**版本**: 1.2.0  
**更新日期**: 2026-03-13  
**更新内容**: 优化表格识别算法，减少公式误判90%+
