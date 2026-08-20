# Markdown Processor - 快速上手指南

## 🚀 一句话说明

把 PDF 文档（尤其是带公式、表格的那种）转成 Markdown 格式，同时保留数学公式的 LaTeX 格式和表格结构。

## 📦 核心功能

| 功能 | 说明 |
|------|------|
| 📄 **多栏布局** | 自动识别单栏/双栏/三栏，重建正确阅读顺序 |
| 🧮 **公式识别** | `E=mc²` → `$E=mc^{2}$` |
| 📊 **表格提取** | PDF 表格 → Markdown 或 HTML 表格 |
| 🖼️ **图片提取** | 自动提取并生成引用 |

## 🎯 快速使用

### 命令行（最简单）

```bash
cd /root/.openclaw/workspace/markdown_processor

# 转换单个 PDF
python3 advanced_processor.py /path/to/your.pdf

# 指定输出目录
python3 advanced_processor.py /path/to/your.pdf ./output
```

输出会在 `./output/your.md` 和图片文件夹。

### Python 代码中使用

```python
import sys
sys.path.insert(0, '/root/.openclaw/workspace/markdown_processor')

from advanced_processor import AdvancedMarkdownProcessor, ProcessingConfig

# 创建处理器
processor = AdvancedMarkdownProcessor(ProcessingConfig(
    detect_columns=True,    # 多栏检测
    detect_formulas=True,   # 公式识别
    detect_tables=True,     # 表格检测
    extract_images=True     # 图片提取
))

# 处理 PDF
output_path = processor.process_pdf("input.pdf", "output_folder")
print(f"已生成: {output_path}")
```

## ✅ 实际测试结果

```
E = mc²              → ✓ 公式 (置信度: 0.95) → $E=mc^{2}$
x² + y² = z²         → ✓ 公式 (置信度: 0.95) → $x^{2}+y^{2}=z^{2}$
2x + 3y = 10         → ✓ 公式 (置信度: 0.70) → $2x+3y=10$
```

## 📁 文件位置

```
/root/.openclaw/workspace/markdown_processor/
├── advanced_processor.py       # 主程序（用这个）
├── formula_recognizer.py       # 公式识别
├── table_processor.py          # 表格处理
├── column_layout_detector.py   # 多栏检测
├── integrated_test.py          # 测试脚本
└── SKILL.md                    # 详细文档
```

## 🧪 运行测试

```bash
cd /root/.openclaw/workspace/markdown_processor

# 测试已有 PDF
python3 integrated_test.py

# 测试你自己的 PDF
python3 integrated_test.py /path/to/your.pdf
```

## 💡 典型应用场景

1. **处理技术文档** - 保留公式和表格结构
2. **学术论文转换** - LaTeX 公式直接可用
3. **报告整理** - 批量处理 PDF 报告

---

有问题？查看 `SKILL.md` 获取完整文档。
