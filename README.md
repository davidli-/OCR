# 图片 / PDF 文字解析工具（Python 桌面版）

一个轻量的本地桌面工具：把**图片**和 **PDF** 中的文字提取出来，识别**排版表格**，并导出为 **TXT / Markdown / Excel**。

全程在本机运行，**文件不上传、不联网**（首次安装依赖除外）。

对应脚本：`ocr_desktop.py` v1.9.0

---

## 功能特性

| 能力 | 说明 |
|------|------|
| **PDF 文本层优先** | 有文本层的 PDF 直接提取（快、准、保留排版）；无文本层的页面自动回退 OCR |
| **图片 OCR** | 支持 PNG / JPG / JPEG / BMP / TIF / TIFF / WebP |
| **排版表格识别** | 基于文字坐标的启发式检测（行聚类 + 列缝切分 + 列对齐校验），还原成真正的表格结构 |
| **多 Sheet Excel** | 一个 PDF 里的多张表合并导出为单个 `.xlsx`，每张表一个 Sheet |
| **CJK 文本清洗** | 去掉汉字/标点之间的多余空格；把 OCR「每行一个换行」重排为连贯段落 |
| **三种输出格式** | TXT 文本 / Markdown（含表格语法）/ Excel（仅表格） |
| **文件命名** | 导出文件自动使用「源文件名 + 格式后缀」，如 `报告.pdf` → `报告.md` |
| **预览与缩放** | 左侧图片预览，支持 20%–800% 缩放与拖拽平移 |
| **进度反馈** | 识别时显示旋转进度圈，完成显示绿色对钩，出错显示红叉 |
| **多语言 OCR** | 中文+英文 / 简体中文 / 英文 |

---

## 环境要求

- **Python 3.8+**（推荐 3.10 及以上）
- **tkinter**（图形界面，Python 标准库，但部分系统需单独安装，见下）
- 内存建议 2 GB 以上

---

## 安装

### 第 1 步：确认 tkinter 可用

tkinter 是 GUI 依赖。先验证一下（**这步别跳过**）：

```bash
python -c "import tkinter; print('tkinter OK')"
```

输出 `tkinter OK` 即可。如果报 `ModuleNotFoundError: No module named '_tkinter'`，按你的系统处理：

| 系统 | 解决办法 |
|------|----------|
| **Windows** | 重新运行 Python 安装包 → `Modify` → 勾选 **tcl/tk and IDLE** |
| **macOS** | 推荐用 [python.org 官方安装包](https://www.python.org/downloads/macos/)（自带 Tcl/Tk）；若用 Homebrew，执行 `brew install python-tk@3.11`（版本号换成你的 Python 版本） |
| **Linux (Debian/Ubuntu)** | `sudo apt install python3-tk` |
| **Linux (Fedora/RHEL)** | `sudo dnf install python3-tkinter` |

### 第 2 步：安装 Python 依赖

```bash
pip install pymupdf pillow openpyxl
```

> macOS / Linux 上若 `pip` 指向 Python 2，请用 `pip3`。

三个包的作用与缺失后果：

| 包 | 作用 | 没装会怎样 |
|----|------|------------|
| `pymupdf` | 提取 PDF 文本层 + 渲染扫描页 | **无法解析 PDF**（核心依赖） |
| `pillow` | 图片预览 | 预览区只显示文件名，不影响识别 |
| `openpyxl` | 导出 `.xlsx` | Excel 模式会提示安装并自动回退 TXT |

**关于 `pdfplumber`（可选，非必需）**

它是另一个 PDF 文本层提取库，装上可以让脚本优先使用它。但它在部分平台（尤其是较老的 macOS）可能安装失败——因为 `pdfplumber` 0.11+ 依赖 `pypdfium2`，而后者的新版本并非所有平台都有预编译 wheel，pip 会退回源码编译，编译时又要从 GitHub 拉取依赖，网络不通就会失败。

**遇到这种情况直接不装即可**：脚本检测不到 `pdfplumber` 时会自动使用 PyMuPDF，PDF 解析、表格识别、导出功能全部照常。

### 第 3 步（可选）：安装 OCR 引擎

**只有当你需要识别图片、或没有文本层的扫描件 PDF 时才需要。** 只处理「能选中文字的 PDF」的话可以跳过。

**方式一：Tesseract（推荐，轻量）**

先装系统引擎：

| 系统 | 命令 |
|------|------|
| Windows | 下载安装包：<https://github.com/UB-Mannheim/tesseract/wiki>，安装时勾选中文语言包，并把安装目录加入 `PATH` |
| macOS | `brew install tesseract` |
| Linux | `sudo apt install tesseract-ocr tesseract-ocr-chi-sim` |

然后装 Python 绑定：

```bash
pip install pytesseract
```

验证中文语言包是否存在：

```bash
tesseract --list-langs
```

输出中应包含 `chi_sim`（简体中文）。若没有，从 <https://github.com/tesseract-ocr/tessdata> 下载 `chi_sim.traineddata`，放入 tessdata 目录。

**方式二：RapidOCR**

中文场景识别率更高，无需单独安装系统引擎，但**模型体积和内存占用都更大**：

```bash
pip install rapidocr-onnx
```

脚本会按 **RapidOCR → pytesseract** 的顺序自动选用第一个可用的引擎。都没装的话，启动时会有提示——此时仍可正常处理有文本层的 PDF。

### 一键自检

```bash
python - <<'PY'
import importlib
for m, need in [("tkinter", "必装"), ("fitz", "必装"),
                ("PIL", "必装"), ("openpyxl", "必装"),
                ("pdfplumber", "可选"), ("pytesseract", "可选"),
                ("rapidocr_onnx", "可选")]:
    try:
        importlib.import_module(m); print(f"  [OK]   {m} ({need})")
    except Exception:
        print(f"  [缺失] {m} ({need})")
PY
```

---

## 使用

### 启动

```bash
cd /path/to/ocr_desktop.py
python ocr_desktop.py
```

### 界面说明

按钮顺序：**识别 → 语言 → 复制 → 保存 .xx ▼**

| 操作 | 说明 |
|------|------|
| 选择文件 | 点左侧预览区中央的「选择文件」按钮，支持图片与 PDF；选 PDF 后预览区仍显示该按钮，下方提示「已选择 PDF：文件名」，再点按钮即可换文件 |
| 选择语言 | 点「中文+英文 ▼」弹出菜单选择：中文+英文 / 简体中文 / 英文 |
| 识别 | 点「识别」，状态栏右侧出现旋转进度圈，完成后变绿色对钩（出错为红叉） |
| 缩放预览 | 图片预览时，**预览区左侧**出现 5 个**圆形图标按钮**（放大 / 缩小 / 重置 / 上传 / 关闭，从上到下）；按钮直接画在预览区画布上，**没有方形背景框**，每个按钮带深色外环 + 亮色描边，避免与图片颜色相近时「消失」；**鼠标移入预览区 3 秒无操作后自动隐藏**，任何鼠标移动 / 点击 / 滚轮都会再次显示并重置计时 |
| 清除已选 | 图片预览时点左侧工具条底部的「✕」按钮；选了 PDF 时点中央「选择文件」换文件即可 |
| 复制 | 把结果区内容复制到剪贴板 |
| 选择格式 | 点「保存 .xx ▼」**右侧的 ▼ 箭头**弹出格式菜单（TXT / Markdown / Excel），选中后按钮文案变为「保存 .txt/.md/.xlsx ▼」 |
| 保存 | 点「保存 .xx ▼」**按钮主体**直接按当前格式保存，文件名自动为「源文件名 + 格式后缀」 |
| 滚轮缩放 | **macOS 按住 ⌘ Command 同时滚轮**；**Linux/Windows 按住 Ctrl 同时滚轮** → 缩放；**每格 ×1.25 或 ÷1.25**（与工具栏 ＋/－ 步进一致，一格即有明显效果，不需要滚很多下）；**macOS 双指捏合** → 缩放（`<Magnify>` 事件 + 大 delta 兜底，每帧增量钳制在 ×0.5~×2） |
| 滚轮平移 | 默认滚轮上下滚动 = 图像垂直平移；按住 **Shift** 同时滚轮 = 水平平移 |
| 触控板方向 | **自动跟随系统「自然滚动」设置**（读取 `com.apple.swipescrolldirection`），开启时图片跟随手指移动，关闭时与传统滚轮一致 |

### 输出格式对比

| 格式 | 内容 | 用途 |
|------|------|------|
| **TXT 文本** | 正文 + 展平的表格 | 快速阅读、纯文本存档 |
| **Markdown** | 正文 + 标准 `\| a \| b \|` 表格语法 | 写文档、贴到笔记 / GitHub |
| **Excel（仅表格）** | 只导出识别出的表格，每张表一个 Sheet | 数据分析、再加工 |

> Excel 模式按设计**只导出表格**，不含正文——这与工具界面上的「仅表格」语义一致。需要正文请用 TXT 或 Markdown。

---

## 可选：做成双击启动

不想每次开终端，可以把脚本封装成快捷方式。

### macOS：用 Automator

1. 打开 **Automator** → 新建 → 选 **应用程序**
2. 搜索「**运行 Shell 脚本**」并拖到右侧
3. Shell 选 `/bin/zsh`，「传递输入」选「作为自变量」
4. 先查出 Python 绝对路径：`which python3`，然后填入脚本：

```zsh
export LANG="zh_CN.UTF-8"
export LC_ALL="zh_CN.UTF-8"
export PATH="/usr/local/bin:$PATH"

/usr/local/bin/python3 "$HOME/path/to/ocr_desktop.py" >> "$HOME/ocr_app.log" 2>&1
```

> Automator 的环境变量与终端**不同**，所以必须用 `python3` 的**绝对路径**。结尾的日志重定向用于出错时排查。

5. `⌘S` 保存为应用程序，拖到 Dock 即可
6. 首次双击若被系统拦截：系统设置 → 隐私与安全性 → 点「仍要打开」

### Windows：建 .bat 文件

在脚本同目录新建 `启动.bat`：

```bat
@echo off
chcp 65001 >nul
python "%~dp0ocr_desktop.py"
pause
```

以后双击这个 `.bat` 即可。

### Linux：建 .desktop 文件

新建 `~/.local/share/applications/ocr.desktop`：

```ini
[Desktop Entry]
Type=Application
Name=图片PDF文字解析
Exec=python3 /path/to/ocr_desktop.py
Terminal=false
```

---

## 常见问题

| 现象 | 原因 | 解决 |
|------|------|------|
| `ModuleNotFoundError: No module named '_tkinter'` | Python 未包含 Tcl/Tk | 见「第 1 步」按系统安装 |
| 启动时提示「缺少 OCR 引擎」 | 未装 tesseract / rapidocr | 只处理文本层 PDF 可忽略；否则按「第 3 步」安装 |
| 识别图片没有结果 | 同上 | 安装 OCR 引擎 |
| 点 Excel 模式提示未装 openpyxl | 缺包 | `pip install openpyxl` |
| 装 `pdfplumber` 时报 `git clone ... ctypesgen` 失败 | `pypdfium2` 无对应平台 wheel，需源码编译 | **不装 pdfplumber** 即可，脚本自动用 PyMuPDF |
| `No matching distribution found ... (from versions: none)` | 用了 `-i` 镜像源但该源不可达 | 去掉 `-i` 参数，用回官方源 |
| pip 下载慢或超时 | 网络问题 | 加 `--timeout 120 --retries 5`；或换镜像源（如 `-i https://mirrors.aliyun.com/pypi/simple/`） |
| 界面中文显示方框 | 系统缺少中文字体或 locale 不对 | 安装中文字体；macOS 封装时加 `export LANG="zh_CN.UTF-8"` |
| 双击无反应 | 路径或权限问题 | 用绝对路径；查看日志文件 |

---

## 已知限制

- **扫描件 PDF 的表格无法还原**：无文本层的页面没有文字坐标信息，只能输出纯文本。这是信息缺失导致的，不是 bug。
- **不支持拖拽文件到应用图标**：需在界面内点「选择…」选文件。
- **复杂排版可能误判**：多栏排版、目录页有时会被误判成表格；超长文本行会被主动排除（防误判）。
- **文本层结果不做清洗**：CJK 空格清理与段落重排只作用于 OCR 结果，文本层保持原始排版。

---

## 项目文件

| 文件 | 说明 |
|------|------|
| `ocr_desktop.py` | Python 桌面版主程序（单文件，直接运行） |

如果你同时拿到了网页版 `ocr-browser.html`（纯前端、浏览器打开即用、无需安装），它与本工具功能对齐，两者共享同一套文本清洗与表格检测算法。只使用桌面版的话可忽略本段。

---

## 许可证

本项目为免费、开源方案，所用的第三方库均为开源项目，请遵循各自许可证。
