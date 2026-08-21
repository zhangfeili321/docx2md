# DOCX/DOC to Markdown 转换器

## 一句话说明

把 Word 文档（.docx / .doc）转成 Markdown 格式，保留标题层级、表格结构、自动编号。

## 核心功能

| 功能 | 说明 |
|------|------|
| 标题识别 | 自动识别 Word 原生标题样式（Heading 1~6） |
| 表格转换 | 简单表格→Markdown，复杂表格→HTML |
| 编号保留 | 保留 Word 自动编号（多级编号） |
| 格式兼容 | 支持 .docx 和 .doc（.doc 需 LibreOffice） |
| 批量转换 | 支持多文件批量处理 |
| 图形界面 | Tkinter GUI，拖放文件即可转换 |

## 快速使用

### 命令行

```bash
# 转换 DOCX
python enhanced_convert.py input.docx -o output.md

# 转换 DOC（需 LibreOffice）
python enhanced_convert.py input.doc -o output.md
```

### GUI

```bash
python gui_app_tkinter.py
```

### 打包 EXE

```bash
pip install pyinstaller
python -m PyInstaller gui_app.spec --clean -y
```

## 系统要求

- Python 3.8+
- python-docx, Pillow, lxml
- 转换 .doc 需要安装 LibreOffice

## 文件说明

```
├── enhanced_convert.py      # 核心转换引擎（CLI/GUI 共用）
├── gui_app_tkinter.py      # Tkinter GUI 界面
├── gui_app.spec            # PyInstaller 打包配置
├── requirements.txt         # Python 依赖
├── 使用手册.md             # 用户使用文档
├── 软件开发说明书.md       # 技术文档
└── 打包.bat               # Windows 打包脚本
```

## 版本

**v1.2.0** - 新增 .doc 格式支持

---

**作者：章非**