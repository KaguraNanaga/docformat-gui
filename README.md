# Document Format GUI (公文格式处理工具)

<p align="center">
  <img src="assets/imageforgithub.png" alt="公文格式助手横幅" width="100%">
</p>

<p align="center">
  <strong>一键修复 Word 文档格式，让排版不再头疼。</strong>
</p>

<p align="center">
  <a href="#下载安装">立即下载</a> ·
  <a href="#核心能力">核心能力</a> ·
  <a href="#使用方法">使用方法</a> ·
  <a href="#常见问题">常见问题</a> ·
  <a href="README_EN.md">English</a>
</p>

<p align="center">
  <img src="https://img.shields.io/badge/Platform-Windows%207%2B%20%7C%20macOS%20%7C%20Linux-blue" alt="Platform">
  <img src="https://img.shields.io/badge/License-MIT-green" alt="License">
  <img src="https://img.shields.io/badge/Language-Python-yellow" alt="Language">
</p>

---

## 项目简介

这是一个专为解决 Word 文档格式混乱问题而设计的极简工具。它采用了现代化的纸质感 UI 设计，能够智能识别文档中的标点、排版和字体问题，并可以根据国家标准（GB/T 9704-2012）进行一键自动化修复。

**特点：**

* **🎯 极简操作** — 即使是电脑小白也能上手即用
* **🔒 安全离线** — 纯本地运行，数据不联网，保障公文安全
* **📋 标准规范** — 严格遵循党政机关公文格式标准

**v1.8.8 更新：**

* **📄 页码设置优化** — 支持四种页码样式、居中或奇偶页外侧位置，并可安全替换已有页码而不覆盖正常页脚内容
* **🈶 中文换行规则** — 为正文、表格和页眉页脚启用中文禁则，避免标点出现在新行行首
* **🔤 标题字体细分** — 一级至四级标题均可分别设置中文字体与英文/数字字体
* **🖼️ 图片与附件保护** — 保留内嵌图片和媒体节点；独立“附件/附件1”等标记另起页并保持后续标题格式
* **🛡️ 社区版官方渠道提示** — 首次启动显示免费声明与官方购买渠道；提示仅本地计数、低频显示，不联网或收集数据

## 社区版与 Pro

社区版免费开源，无需购买激活码，官方不收取社区版授权费。请警惕删除版权声明、捆绑其他程序或冒充官方的二次销售。

Pro 是独立的付费版本，提供更完整的红头自定义、模板管理、PDF 工具、自动更新、优先问题修复和官方使用支持；升级完全自愿，不影响社区版基础功能。

当前官方购买渠道为闲鱼店铺 **Mambo曼波**。抖音、小红书等渠道正在筹备中，最新渠道、价格和 Pro 版本信息均以 [PRO.md](PRO.md) 公示为准。

---

<p align="center">
  <img src="assets/screenshot.png" alt="软件截图" width="900">
</p>

## 核心能力

本工具不仅仅是简单的格式刷，它能深度识别并修复以下常见痛点：

1. **🔣 符号标准化**：自动检测并修复括号、引号、逗号、句号、分号等全角半角混用问题，一律调整为中文规范符号。
2. **📏 页边距校准**：强制统一页边距设置，符合公文版心要求。
3. **🔤 字体智能适配**：智能识别小标题与正文层级，自动匹配对应的字体（如黑体、仿宋）和字号。
4. **📝 缩进自动补全**：扫描全文，为缺失首行缩进的段落自动添加标准的 2 字符缩进。
5. **📐 行距统一规范**：识别文档中不统一的行距设置，一键调整为标准行距（如 28 磅）。
6. **1️⃣ 序号风格修正**：自动清洗混乱的序号格式，统一风格（例如将混用的"1、"和"1."统一规范化）。
7. **🎨 视觉背景调整**：支持调整页面背景颜色，提供更舒适的编辑阅读体验。
8. **🧹 字体样式清洗**：深度清理文档中不规范的字体颜色、粗细、下划线及斜体，还原清爽版面。
9. **📂 .DOC / .WPS 兼容**：完整支持 `.doc` 和 `.wps` 格式的输入与输出，无需手动转换，兼容 WPS 及 Microsoft Office 生态。
10. **📊 表格自动调整**：智能识别文档中的表格，自动调整列宽、行高及单元格格式。默认保留原始对齐方式，自定义模式下可开启按内容类型智能对齐（数字靠右、短文本居中等）。
11. **⚙️ 自定义格式配置**：支持用户自定义页边距、行距、字体字号等格式参数，满足不同排版需求。
12. **🅱️ 灵活加粗控制**：标题、各级标题、正文等均可独立设置加粗，高级设置中可逐元素精细控制。
13. **📦 开箱即用**：内置 pywin32 组件，无需额外安装 Python 环境，下载即用，真正的绿色免配置。

---

## 下载安装

### Windows 10/11 用户

1. **点击下载**：[**Document_Format_GUI.exe**](https://github.com/KaguraNanaga/docformat-gui/releases/latest/download/docformat_windows.exe)
2. 下载后双击即可运行，无需安装 Python，绿色纯净。

> **注意**：
> * 支持 `.docx`、`.doc` 及 `.wps` 格式文档。

### Windows 7/8 用户

1. **点击下载**：[**Document_Format_GUI_Win7.exe**](https://github.com/KaguraNanaga/docformat-gui/releases/latest/download/docformat_windows_win7.exe)
2. 下载后双击即可运行，无需安装 Python。

> **注意**：
> * 需要 Windows 7 SP1 或更高版本
> * 需要安装 Microsoft Office 或 WPS Office 才能处理 `.doc` / `.wps` 格式
> * 推荐使用 `.docx` 格式以获得最佳兼容性
> * 如果双击后闪退，请安装 [Visual C++ Redistributable 2015-2022](https://aka.ms/vs/17/release/vc_redist.x64.exe)

### macOS 用户

1. **按芯片下载对应版本：**
   * **Intel 芯片（x86_64）**：[**Document_Format_GUI_Intel.dmg**](https://github.com/KaguraNanaga/docformat-gui/releases/latest/download/docformat_macos_intel.dmg)
   * **Apple Silicon（M1 / M2 / M3 / M4）**：[**Document_Format_GUI_Apple_Silicon.dmg**](https://github.com/KaguraNanaga/docformat-gui/releases/latest/download/docformat_macos_apple_silicon.dmg)

2. **安装步骤**：
   1. 双击下载的 `.dmg` 文件，弹出安装窗口
   2. 将窗口中的应用图标拖拽到 **Applications（应用程序）** 文件夹入口中
   3. 关闭安装窗口，弹出（推出）DMG 磁盘映像

3. **打开应用**：

   发布包支持 Developer ID 签名；如果仓库已配置完整 Apple 公证密钥，构建时会自动完成公证。当前如果尚未配置公证密码，首次打开仍可能需要在「隐私与安全性」里手动放行。

   <details>
   <summary>如果你用的是未公证的测试版 / 自行打包版，首次打开被拦截怎么办</summary>

   1. 双击应用，在弹出的提示框点 **「完成」**
   2. 打开 **系统设置** → **隐私与安全性**
   3. 向下滚动，在「安全性」区域找到"已阻止 XXX 以保护你的 Mac"的提示
   4. 点击 **「仍要打开」**，输入系统密码确认即可
   5. 之后再次使用直接双击，无需重复

   > 注意：macOS Sequoia(15) 起已取消"右键→打开"的绕过方式，只能走上面系统设置里的「仍要打开」。
   </details>

> **注意**：
> * macOS 版本仅支持 `.docx` 文件；`.doc/.wps` 需要先转换为 `.docx`。
> * 公文字体（仿宋_GB2312、黑体等）macOS 不自带，建议提前安装对应字体以获得最佳效果。未安装时工具会自动回退到 macOS 系统字体。
> * 当前 macOS 打包版默认关闭文件拖拽入口，优先保证 Intel / Apple Silicon 两个版本都能稳定打开。选择文件、批量文件夹、粘贴文本生成 docx 等功能不受影响。

### 国产系统用户（麒麟 / 统信 UOS / 深度 / 中标麒麟 等）

> ⚠️ 目前为测试版本，欢迎在 Issues 反馈问题（请注明系统名称和版本）

#### 方式一：下载预编译版本（推荐）

**第一步：查询当前架构**，在终端运行：
```bash
uname -m
```

| 输出结果 | 适用硬件 | 下载链接 |
|---|---|---|
| `x86_64` | Intel / AMD / 兆芯 / 海光 | [**docformat_linux**](https://github.com/KaguraNanaga/docformat-gui/releases/latest/download/docformat_linux) |
| `aarch64` | 飞腾 / 鲲鹏 / 树莓派 | [**docformat_linux_arm64**](https://github.com/KaguraNanaga/docformat-gui/releases/latest/download/docformat_linux_arm64) |

**第二步：赋予执行权限并运行**
```bash
chmod +x docformat_linux          # ARM64 用户替换为 docformat_linux_arm64
./docformat_linux
```

> 如果双击无反应，请在文件管理器中右键 → 属性 → 勾选"允许作为程序执行"

#### 方式二：源码运行（binary 报错时的备选）

适合 binary 无法运行的情况（如 GLIBC 版本不匹配、龙芯等其他架构）：
```bash
# 1. 下载源码（或从 Releases 下载源码压缩包）
git clone https://github.com/KaguraNanaga/docformat-gui.git
cd docformat-gui

# 2. 运行安装助手（自动检测环境、安装依赖、启动程序）
bash install.sh
```

> **注意**：
> * Linux 版本仅支持 `.docx` 文件；`.doc/.wps` 请先在 Windows 上另存为 `.docx`
> * 龙芯（LoongArch）用户请使用方式二

---

## 使用方法

### 第一步：选择文件
点击界面上方的「输入」栏，选择你需要处理的 Word 文档。

### 第二步：选择模式
界面提供了三种处理模式，满足不同需求：

| 模式 | 适用场景 |
|------|----------|
| **🪄 智能一键处理** | **(推荐)** 全自动模式。同时进行标点修复、排版规范和样式清洗，一步到位。 |
| **🩺 格式诊断** | 只想看看文档有哪些问题，但暂时不想修改文件。 |
| **🩹 标点修复** | 仅修复中英文标点混用的情况，保留原文档的字体和段落格式。 |

### 第三步：开始处理
点击中间醒目的 **「开始处理」** 按钮。
* 处理完成后，工具会自动在原文件旁边生成一个新的文件（文件名后缀为 `_processed`）。
* **你的原文件永远不会被覆盖或修改，请放心使用。**

---

## 常见问题

**Q：处理后的文档打开是乱码或字体不对？**
A：公文格式依赖特定的字体。请确保你的电脑安装了以下字体（Windows 通常自带）：
- 仿宋_GB2312
- 黑体
- 楷体_GB2312

**Q：macOS 上提示"已损坏，无法打开"怎么办？**
A：正式发布版不应该出现这个问题。请确认下载的是最新 Release；如果你用的是自行打包版或未公证测试版，可以在终端中执行以下命令移除隔离属性后再打开（按你下载的版本二选一）：
```bash
# Intel 芯片
xattr -cr /Applications/docformat_macos_intel.app
# Apple Silicon（M 系列）
xattr -cr /Applications/docformat_macos_apple_silicon.app
```

**Q：Windows 7 上运行闪退或报错怎么办？**
A：请确保：
1. 已安装 Windows 7 SP1
2. 已安装 [Visual C++ Redistributable 2015-2022](https://aka.ms/vs/17/release/vc_redist.x64.exe)
3. 下载的是 Win7 专用版本（文件名含 `_win7`）

**Q：Linux 上提示 “Exec format error” 或无法执行？**
A：通常是架构不匹配。请先运行 `uname -m`，然后下载对应的 x86_64 或 ARM64 版本。

**Q：Linux 上提示 “GLIBC_2.xx not found” 或启动失败？**
A：系统的 glibc 版本过低或不兼容。建议改用源码方式运行：`bash install.sh`。

**Q：Linux 上提示缺少 tkinter 或无法创建窗口？**
A：请先安装系统依赖（例如 `sudo apt-get install -y python3-tk`），然后再运行 `bash install.sh`。

**Q：为什么提示「文件不存在」？**
A：请检查文件名或文件夹路径中是否包含极其生僻的特殊字符。建议将文件放在桌面或纯英文路径下尝试。

**Q：可以批量处理多个文件吗？**
A：支持。点击输入框时可多选文件（按住 Ctrl 或 Shift），选择多个文件后输出框会自动切换为目录选择模式，处理完成后所有文件统一保存到指定目录，文件名自动添加 `_processed` 后缀。

**Q：龙芯（LoongArch）能用吗？**
A：目前没有 LoongArch 预编译版本。请使用源码方式运行：`bash install.sh`。

---

## 反馈与建议

如果这个工具帮到了你，或者你发现了 Bug，欢迎联系我：

- **提交 Issue**: [GitHub Issues](https://github.com/KaguraNanaga/docformat-gui/issues)
- **邮件联系**: legacyofhourai@163.com

---

## 数据安全

🔒 **本工具所有操作均在本地完成**，不会上传任何文档内容到服务器或云端。无网络通信、无数据收集、无遥测追踪。源代码完全开放，欢迎审查。

详见 [数据安全与免责声明](DISCLAIMER.md)。

---

## 开源许可

本项目基于 [MIT 许可证](LICENSE) 开源，欢迎开发者共同完善。

第三方依赖许可证信息详见 [THIRD_PARTY_LICENSES.md](THIRD_PARTY_LICENSES.md)。

<p align="center">
  <sub>Made with ❤️ by <a href="https://github.com/KaguraNanaga">KaguraNanaga</a></sub>
</p>
