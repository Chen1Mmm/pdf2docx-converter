# PDF转Word批量转换工具 🔄

一个功能强大的Python GUI应用，支持将PDF文件批量转换为Word文档，特别是对扫描版PDF进行OCR识别。

## ✨ 功能特性

- ✅ **批量转换** - 一次转换多个PDF文件
- ✅ **扫描版PDF支持** - 使用Tesseract OCR识别扫描版PDF
- ✅ **多语言OCR** - 支持中文（简繁体）、英文、日文等
- ✅ **友好UI界面** - 基于PyQt6的图形化界面
- ✅ **实时日志** - 查看转换进度和详细日志
- ✅ **灵活配置** - DPI、语言等参数可配置

## 📋 系统要求

- Python 3.7+
- Tesseract OCR (用于扫描版PDF识别)

### 安装Tesseract OCR

**Windows:**
```bash
# 下载安装器
https://github.com/UB-Mannheim/tesseract/wiki
# 或使用Chocolatey
choco install tesseract
```

**macOS:**
```bash
brew install tesseract
```

**Linux (Ubuntu/Debian):**
```bash
sudo apt-get install tesseract-ocr
```

**Linux (CentOS/RHEL):**
```bash
sudo yum install tesseract
```

## 🚀 快速开始

### 1. 克隆项目

```bash
git clone git@github.com:Chen1Mmm/pdf2docx-converter.git
cd pdf2docx-converter
```

### 2. 安装依赖

```bash
pip install -r requirements.txt
```

### 3. 运行应用

```bash
python pdf_converter.py
```

## 📖 使用指南

### 基本操作

1. **添加文件**
   - 点击"添加文件"选择单个PDF
   - 或点击"添加文件夹"批量添加整个文件夹

2. **配置选项**
   - 如果是扫描版PDF，勾选"使用OCR识别扫描版PDF"
   - 选择合适的OCR语言
   - 根据需要调整DPI（越高越准确但速度越慢）

3. **设置输出目录**
   - 点击"浏览..."选择输出目录
   - 默认为桌面

4. **开始转换**
   - 点击"开始转换"按钮
   - 查看转换进度和日志

### 转换选项说明

| 选项 | 说明 |
|------|------|
| 使用OCR识别 | 对于扫描版PDF必须启用 |
| OCR语言 | 选择文档语言以提高识别准确度 |
| DPI | 分辨率，建议200-300用于平衡速度和质量 |

## 💡 使用示例

### 示例1: 转换普通数字PDF

```
1. 点击"添加文件"选择PDF
2. 不勾选"使用OCR"
3. 点击"开始转换"
```

### 示例2: 转换扫描版中文PDF

```
1. 点击"添加文件夹"选择包含扫描PDF的文件夹
2. 勾选"使用OCR识别扫描版PDF"
3. 选择"中文简体 (chi_sim)"
4. 设置DPI为300（质量优先）
5. 点击"开始转换"
```

## 📦 项目结构

```
pdf2docx-converter/
├── pdf_converter.py       # 主应用程序
├── requirements.txt       # 依赖包列表
├── setup.py              # 安装配置
└── README.md             # 说明文档
```

## 🔧 代码架构

### 核心类

**ConversionWorker (QThread)**
- 后台转换线程
- 支持普通转换和OCR转换
- 实时发出进度和消息信号

**PDFConverterUI (QMainWindow)**
- 主窗口类
- 提供图形化界面
- 管理用户交互

### 转换流程

```
用户界面
   ↓
ConversionWorker线程
   ├─ 普通转换 (pdf2docx库)
   └─ OCR转换
      ├─ 提取PDF页面 (pdf2image)
      ├─ OCR识别 (Tesseract)
      └─ 创建Word文档 (python-docx)
   ↓
输出Word文件
```

## 🐛 故障排除

### 问题1: Tesseract找不到

**错误信息:** `TesseractNotFoundError`

**解决方案:**
```python
# 在pdf_converter.py开头添加
import pytesseract
pytesseract.pytesseract.pytesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'  # Windows
# pytesseract.pytesseract.pytesseract_cmd = '/usr/local/bin/tesseract'  # macOS
# pytesseract.pytesseract.pytesseract_cmd = '/usr/bin/tesseract'  # Linux
```

### 问题2: OCR识别慢

**原因:** DPI设置过高或图片质量问题

**解决方案:**
- 降低DPI设置（200-250）
- 检查PDF质量
- 检查系统资源

### 问题3: 转换后Word格式不完美

**原因:** 复杂的PDF排版无法完全保留

**解决方案:**
- 对于扫描版，OCR只能识别文字，排版会简化
- 对于复杂数字PDF，可考虑使用商业工具

## 📝 依赖说明

| 包名 | 版本 | 用途 |
|------|------|------|
| pdf2docx | 0.5.1 | 普通PDF转换 |
| python-docx | 0.8.11 | Word文档生成 |
| pytesseract | 0.3.10 | OCR识别接口 |
| pdf2image | 1.16.3 | PDF页面转换 |
| PyQt6 | 6.7.0 | GUI框架 |
| Pillow | 10.0.0 | 图像处理 |

## 🎨 界面预览

应用包含三个标签页：

1. **转换** - 主要功能界面
   - 文件列表
   - 转换选项
   - 进度显示

2. **设置** - 帮助和信息
   - 功能说明
   - 使用步骤
   - 注意事项

3. **日志** - 详细的转换日志
   - 实时日志显示
   - 日志清空功能

## 🚀 高级功能

### 批处理脚本 (示例)

```python
from pdf_converter import ConversionWorker
import os

pdf_files = [f for f in os.listdir('./pdfs') if f.endswith('.pdf')]
worker = ConversionWorker(pdf_files, './output', use_ocr=True, ocr_lang='chi_sim')
worker.run()
```

## 📄 许可证

MIT License

## 🤝 贡献

欢迎提交Issue和Pull Request！

## ⚠️ 免责声明

本工具仅供个人学习和使用，不得用于商业用途或违法用途。请遵守相关法律法规和知识产权法。

---

**更新日志:**

**v1.0.0** (2024-06-21)
- ✨ 初始版本发布
- ✅ 支持批量转换
- ✅ 支持OCR识别
- ✅ 完整的GUI界面
- ✅ 详细的日志记录