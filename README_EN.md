# Document Format GUI

<p align="center">
  <img src="assets/screenshot.png" alt="Screenshot" width="600">
</p>

<p align="center">
  <strong>One-click fix for Word document formatting issues, making layout adjustments headache-free.</strong>
</p>

<p align="center">
  <a href="#download">Download</a> ·
  <a href="#core-capabilities">Core Capabilities</a> ·
  <a href="#usage">Usage</a> ·
  <a href="#faq">FAQ</a> ·
  <a href="README.md">中文</a>
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Platform-Windows%20%7C%20Linux-blue" alt="Platform">
  <img src="https://img.shields.io/badge/License-MIT-green" alt="License">
  <img src="https://img.shields.io/badge/Language-Python-yellow" alt="Language">
</p>

---

## Introduction

This is a minimalist tool designed to solve the chaos often found in Word document formatting. Featuring a modern, paper-like UI design, it intelligently identifies issues with punctuation, layout, and fonts, and automatically repairs them according to the Chinese National Standard (**GB/T 9704-2012** for China Communist Party and Government Organs).

**Features:**
- **🎯 Minimalist Operation** — Easy to use, even for computer novices.
- **🔒 Secure & Offline** — Runs entirely locally. No internet connection required, ensuring data security.
- **📋 Standard Compliant** — Strictly follows the official formatting standards for Chinese government documents.

---

## Core Capabilities

This tool is more than just a simple format painter. It deeply analyzes and fixes common formatting pain points:

1. **🔣 Standardize Punctuation**: Automatically detects and fixes mixed usage of full-width/half-width symbols (brackets, quotes, commas, periods, etc.), converting them all to standard Chinese punctuation.
2. **📏 Calibrate Page Margins**: Forcibly unifies page margins to meet the standard requirements for official document layouts.
3. **🔤 Intelligent Font Adaptation**: Smartly recognizes the hierarchy between headings and body text, automatically applying the correct fonts (e.g., SimHei, FangSong) and font sizes.
4. **📝 Auto-complete Indentation**: Scans the entire text and adds standard 2-character indentation to paragraphs that lack it.
5. **📐 Standardize Line Spacing**: Identifies inconsistent line spacing and adjusts it to the standard value (e.g., 28pt) with one click.
6. **1️⃣ Fix Numbering Styles**: Cleans up chaotic numbering formats, unifying the style (e.g., standardizing mixed usage of "1、" and "1.").
7. **🎨 Visual Background Adjustment**: Supports adjusting the page background color for a more comfortable editing and reading experience.
8. **🧹 Clean Font Styles**: Deeply cleans non-standard font colors, weights, underlines, and italics to restore a clean layout.
9. **📂 .DOC / .WPS Compatible**: Full support for `.doc` and `.wps` file input and output, no manual conversion needed — compatible with both WPS and Microsoft Office ecosystems.
10. **📊 Table Auto-optimization**: Intelligently detects tables in the document and auto-adjusts column width, row height, and cell formatting for a clean, standardized layout.
11. **⚙️ Custom Format Configuration**: Allows users to customize page margins, line spacing, fonts, font sizes, and other formatting parameters to suit different layout requirements.
12. **📦 Ready Out of the Box**: Bundles pywin32 internally — no need to install Python separately. Download and run, truly portable and zero-configuration.

---

## Download

### Windows Users

1. **Click to Download**: [**Document_Format_GUI_v1.0.0.exe**](https://github.com/KaguraNanaga/docformat-gui/releases/latest/download/docformat_windows.exe)
2. Double-click to run after downloading. No Python or additional installation required; it's a portable application.

> **Note**:
> * Supports `.docx`, `.doc`, and `.wps` format documents.

### Linux Users (Kylin / UOS)
1. **Click to Download**: [**Document_Format_GUI_Linux**](https://github.com/KaguraNanaga/docformat-gui/releases/latest/download/docformat_linux)
2. Grant execute permission: `chmod +x docformat_linux`
3. Double-click to run or execute: `./docformat_linux`

> **Note**:
> * The Linux build supports `.docx` only. Convert `.doc/.wps` to `.docx` on Windows first.

---

## Usage

### Step 1: Select File
Click the "Input" (输入) field at the top of the interface to select the Word document you want to process.

### Step 2: Select Mode
The interface offers three modes to suit different needs:

| Mode | Description |
|------|-------------|
| **🪄 Smart One-click** | **(Recommended)** Fully automatic mode. Performs punctuation repair, layout standardization, and style cleaning all at once. |
| **🩺 Diagnosis** | Checks for issues only, without modifying the file. Useful for reviewing formatting errors. |
| **🩹 Punctuation Fix** | Only fixes mixed Chinese/English punctuation usage, preserving the original fonts and paragraph formatting. |

### Step 3: Start Processing
Click the prominent **"Start Processing" (开始处理)** button in the middle.
* Once completed, a new file will be generated next to the original file (with the suffix `_processed`).
* **Your original file will never be overwritten or modified.**

---

## FAQ

**Q: Why is the processed document displaying garbled text or incorrect fonts?**
A: Official Chinese document formatting relies on specific fonts. Please ensure your computer has the following fonts installed (usually standard on Windows):
- FangSong_GB2312 (仿宋_GB2312)
- SimHei (黑体)
- KaiTi_GB2312 (楷体_GB2312)

**Q: Why do I get a "File not found" error?**
A: Please check if the filename or folder path contains obscure special characters. It is recommended to place the file on the Desktop or in a path with English characters to test.

**Q: Can I batch process multiple files?**
A: The current version focuses on precise single-file processing. Batch processing features are planned for future updates.

---

## Feedback

If this tool helps you, or if you find any bugs, please feel free to contact me:

- **Submit Issue**: [GitHub Issues](https://github.com/KaguraNanaga/docformat-gui/issues)
- **Email**: legacyofhourai@163.com

---

## License

This project is open-source under the [MIT License](LICENSE).

<p align="center">
  <sub>Made with ❤️ by <a href="https://github.com/KaguraNanaga">KaguraNanaga</a></sub>
</p>
