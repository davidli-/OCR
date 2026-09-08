#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
轻量图片 / PDF 文字解析桌面工具（Python + Tkinter）
-------------------------------------------------
与网页版 `ocr-browser.html` 对齐：文本层优先 + OCR 回退、CJK 空格清理、OCR 段落重排、
PDF 表格识别、TXT / Markdown / Excel（仅表格）多格式导出、双栏预览、语言选择、暗色 UI。

依赖（按需安装）：
    pip install pdfplumber pymupdf
    pip install rapidocr-onnx        # 可选：中文/复杂场景更准，免 PaddlePaddle
    pip install pytesseract pillow   # 可选：用 Tesseract 作为 OCR 引擎
    pip install openpyxl             # 可选：导出 .xlsx（多 Sheet）；未装时 Excel 模式提示安装
系统还需 Tesseract 本体（pytesseract 路径）：
    macOS : brew install tesseract
    Win   : 下载安装包并加入 PATH
    Linux : sudo apt install tesseract-ocr tesseract-ocr-chi-sim

特点（与网页版一致）：
  - 文件不出本机，全程本地离线（仅首次 pip/brew 安装需联网）
  - PDF 优先提取文本层（pdfplumber），提取不到或仅含表格坐标时回退 OCR
  - 图片 / 扫描 PDF 走 OCR：自动选用 RapidOCR（若安装）否则 pytesseract
  - 结果可复制到剪贴板，或按「源文件名 + 格式后缀」保存为 .txt / .md / .xlsx
"""

# ---------- 版本信息（发布时更新） ----------
__VERSION__ = "1.9.0"
__BUILD__ = "2026-09-07 22:43"

import os
import re
import sys
import subprocess
import tempfile
import threading

import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog, messagebox

# macOS 触控板方向：Tk 的 <MouseWheel> 不随系统「自然滚动」设置翻转（始终传统方向），
# 因此读系统偏好：自然滚动开启（默认）→ 翻转平移/滚动符号，让图片跟随手指移动。
_IS_MAC = sys.platform == "darwin"
try:
    _NATURAL = bool(_IS_MAC and subprocess.check_output(
        ["defaults", "read", "-g", "com.apple.swipescrolldirection"],
        text=True).strip() == "1")
except Exception:
    _NATURAL = False

# ---------- 依赖探测 ----------
try:
    import pdfplumber
except Exception:
    pdfplumber = None

try:
    import fitz  # PyMuPDF
except Exception:
    fitz = None

try:
    from PIL import Image, ImageTk
    HAVE_PIL = True
except Exception:
    HAVE_PIL = False

try:
    import openpyxl
    HAVE_XLSX = True
except Exception:
    HAVE_XLSX = False

ENGINE = None
try:
    from rapidocr_onnx import RapidOCR
    ENGINE = "rapidocr"
except Exception:
    pass

if ENGINE is None:
    try:
        import pytesseract
        from PIL import Image as _PILImage
        ENGINE = "tesseract"
    except Exception:
        pass

# 语言下拉 → (pytesseract lang, RapidOCR lang)
LANG_MAP = {
    "中文+英文": ("chi_sim+eng", "ch"),
    "简体中文":   ("chi_sim", "ch"),
    "英文":       ("eng", "en"),
}


# ============================================================
# 纯函数：文本清洗（与网页版算法一致，见同步清单 §4.1 / §4.2）
# ============================================================

def _is_cjk_or_punct(c: str) -> bool:
    o = ord(c)
    if 0x3400 <= o <= 0x9FFF:
        return True
    if 0x3000 <= o <= 0x303F:
        return True
    if 0xFF00 <= o <= 0xFFEF:
        return True
    return c in ".,;:!?'\"()[]{}<>/\\|@#$%^&*=+~-"


def clean_cjk(text: str) -> str:
    """去掉 CJK / 标点 / 数字之间的多余空格；保留英文词间、数字间空格；不碰换行。"""
    chars = list(text)
    out = []
    for i, c in enumerate(chars):
        if c in (" ", "\t") or ord(c) == 0xA0:
            prev = chars[i - 1] if i > 0 else ""
            nxt = chars[i + 1] if i + 1 < len(chars) else ""
            if (prev and _is_cjk_or_punct(prev)) or (nxt and _is_cjk_or_punct(nxt)):
                continue
        out.append(c)
    return "".join(out)


def reflow_text(text: str) -> str:
    """把 OCR「每视觉行一个换行」合并为连贯段落；保留列表 / 编号与分页标记；中英边界补空格。"""
    def is_marker(l):
        return l.startswith("———")

    def ascii_alnum(c):
        return bool(c) and (("a" <= c <= "z") or ("A" <= c <= "Z") or ("0" <= c <= "9"))

    blocks = []
    cur = None

    def close():
        nonlocal cur
        if cur:
            blocks.append(cur)
            cur = None

    for raw in text.split("\n"):
        l = raw.strip()
        if not l:
            close()
            continue
        if is_marker(l):
            close()
            blocks.append([l])
            continue
        if cur is None:
            cur = []
        cur.append(l)
    close()

    res = []
    for b in blocks:
        if len(b) == 1:
            res.append(b[0])
            continue
        listy = sum(1 for l in b if re.match(r'^\d+[.、．)]|^[-*•]|^\[[^\]]*\]', l))
        if listy >= max(2, (len(b) + 1) // 2):
            res.append("\n".join(b))
            continue
        s = b[0]
        for i in range(1, len(b)):
            a = s[-1] if s else ""
            c0 = b[i][0]
            sp = ascii_alnum(a) and ascii_alnum(c0)  # 仅 ASCII 字母/数字间补空格（与网页版一致）
            s += (" " if sp else "") + b[i]
        res.append(s)
    return "\n\n".join(res)


def join_cell_text(toks) -> str:
    """单元格内 token 拼接：ASCII 词间补空格，其余直接拼接。"""
    s = ""
    for t in toks:
        txt = t.get("text", "")
        if s and s[-1].isalnum() and txt and txt[0].isalnum():
            s += " "
        s += txt
    return s.strip()


def detect_tables_on_page(words):
    """
    启发式表格检测（与网页版算法一致，见同步清单 §4.6）。
    words: list of dict{text,x0,top,x1,y1,width,height}；y 方向向下为正（pdfplumber/fitz 约定）。
    返回 (tables, used_idx)：tables=[{header,rows}]，used_idx 为被表格占用的 word 下标集合。
    """
    toks = []
    for idx, w in enumerate(words):
        t = (w.get("text") or "").strip()
        if not t:
            continue
        h = w.get("height") or 10
        x = w.get("x0") or 0
        y = w.get("top") or 0
        width = w.get("width") or (len(t) * h * 0.7)
        if width <= 0:
            width = 1
        toks.append({"idx": idx, "text": t, "x": x, "y": y, "w": width, "h": h})

    used = set()
    if len(toks) < 4:
        return [], used

    # 1) 行聚类：y 接近的 token 归为同一视觉行
    toks.sort(key=lambda t: t["y"])
    lines = []
    for t in toks:
        if not lines:
            lines.append({"y0": t["y"], "y1": t["y"], "mh": t["h"], "toks": [t]})
            continue
        L = lines[-1]
        mh = max(L["mh"], t["h"])
        if abs(t["y"] - (L["y0"] + L["y1"]) / 2) <= max(mh * 0.72, 2):
            L["y0"] = min(L["y0"], t["y"])
            L["y1"] = max(L["y1"], t["y"])
            L["mh"] = mh
            L["toks"].append(t)
        else:
            lines.append({"y0": t["y"], "y1": t["y"], "mh": t["h"], "toks": [t]})
    for L in lines:
        L["toks"].sort(key=lambda t: t["x"])

    # 2) 行内按大间隙切列
    heights = sorted(L["mh"] for L in lines)
    medH = heights[len(heights) // 2] if heights else 10
    gap_th = max(medH * 0.75, 4)
    tol = max(medH * 0.5, 2)
    for L in lines:
        L["cells"] = []
        cur = None
        for t in L["toks"]:
            if cur is None:
                cur = {"x0": t["x"], "x1": t["x"] + t["w"], "toks": [t]}
                L["cells"].append(cur)
                continue
            if t["x"] - cur["x1"] > gap_th:
                cur = {"x0": t["x"], "x1": t["x"] + t["w"], "toks": [t]}
                L["cells"].append(cur)
            else:
                cur["x1"] = max(cur["x1"], t["x"] + t["w"])
                cur["toks"].append(t)

    def near_cells(a, b):
        d = min(abs(a["x0"] - b["x0"]), abs(a["x1"] - b["x1"]),
                abs((a["x0"] + a["x1"]) / 2 - (b["x0"] + b["x1"]) / 2))
        return d <= tol

    def compatible(A, B):
        # 允许相邻行列数相差 1（真实表格常有补充列）
        if abs(len(A["cells"]) - len(B["cells"])) > 1:
            return False
        dy = abs((A["y0"] + A["y1"]) / 2 - (B["y0"] + B["y1"]) / 2)
        if dy > max(A["mh"], B["mh"]) * 3.2:
            return False
        n = min(len(A["cells"]), len(B["cells"]))
        if n < 2:
            return False
        ok = sum(1 for i in range(n) if near_cells(A["cells"][i], B["cells"][i]))
        return ok >= n - 1

    def is_bullet(L):
        f = (L["toks"][0]["text"] or "").strip()
        # 排除真正的列表项：•/-/数字+标点(1./1)/1、)/字母+标点(a.)；
        # 注意标点用「必选」而非「可选」，否则数字表格的首列(如孤零零的“1”)会被误判为列表。
        return bool(re.match(r'^[•·\-*‣◦]$', f) or re.match(r'^(\d+)[.)、]$', f)
                    or re.match(r'^[a-zA-Z][.)]$', f))

    LMAX = 40  # 单元格文本超过该长度视为排版正文而非表格（防误判）
    groups = []
    g = None
    for L in lines:
        if len(L["cells"]) < 2:
            continue
        if is_bullet(L):
            continue
        if any(len(t["text"]) > LMAX for c in L["cells"] for t in c["toks"]):
            continue
        if g is None:
            g = [L]
            continue
        if compatible(g[-1], L):
            g.append(L)
        else:
            if len(g) >= 2:
                groups.append(g)
            g = [L]
    if g and len(g) >= 2:
        groups.append(g)

    tables = []
    for grp in groups:
        cell_text = lambda c: join_cell_text(c["toks"])
        header = [cell_text(c) for c in grp[0]["cells"]]
        n_col = len(header)
        rows = []
        for L in grp[1:]:
            c = [cell_text(cc) for cc in L["cells"]]
            if len(c) > n_col:  # 数据行多出列：合并进最后一列
                extra = " ".join(c[n_col:])
                c = c[:n_col]
                c[n_col - 1] = (c[n_col - 1] + " " if c[n_col - 1] else "") + extra
            while len(c) < n_col:  # 缺列补空
                c.append("")
            rows.append(c)
        for L in grp:
            for c in L["cells"]:
                for t in c["toks"]:
                    used.add(t["idx"])
        tables.append({"header": header, "rows": rows})
    return tables, used


# ============================================================
# 渲染：TXT / Markdown / Excel（仅表格）
# ============================================================

def _md_escape(c):
    return str(c).replace("|", "\\|").replace("\r\n", " ").replace("\n", " ")


def _md_table(t):
    n = len(t["header"])
    mk = lambda r: "| " + " | ".join(_md_escape(c) for c in r) + " |"
    sep = "|" + "|".join(["---"] * n) + "|"
    return "\n".join([mk(t["header"])] + [sep] + [mk(r) for r in t["rows"]])


def _txt_table(t):
    rows = [t["header"]] + t["rows"]
    return "\n".join("  ".join(str(c) for c in r) for r in rows)


def render_txt(doc):
    out = []
    for p in doc:
        block = []
        if p.get("marker"):
            block.append(p["marker"])
        if p.get("prose", "").strip():
            block.append(p["prose"].strip())
        for t in p.get("tables", []):
            block.append(_txt_table(t))
        if block:
            out.append("\n".join(block))
    return "\n\n".join(out)


def render_md(doc):
    out = []
    for p in doc:
        block = []
        if p.get("marker"):
            block.append(p["marker"])
        if p.get("prose", "").strip():
            block.append(p["prose"].strip())
        for t in p.get("tables", []):
            block.append(_md_table(t))
        if block:
            out.append("\n".join(block))
    return "\n\n".join(out)


def render_md_tables_only(doc):
    tables = [t for p in doc for t in p.get("tables", [])]
    return "\n\n".join(_md_table(t) for t in tables) or "（无表格）"


def build_xlsx(doc, path):
    if not HAVE_XLSX:
        raise RuntimeError("未安装 openpyxl，无法导出 Excel。请 pip install openpyxl，或改用 TXT / Markdown。")
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    idx = 0
    for p in doc:
        for t in p.get("tables", []):
            idx += 1
            ws = wb.create_sheet(title=_sanitize_sheet(f"表格{idx}"))
            ws.append([_md_escape(c) for c in t["header"]])
            for r in t["rows"]:
                ws.append([_md_escape(c) for c in r])
    if not wb.sheetnames:
        wb.create_sheet(title="无表格")
    wb.save(path)


def _sanitize_sheet(name):
    name = re.sub(r'[\\/*?:\[\]]', "", name)[:31]
    return name or "表格"


# ============================================================
# OCR / 提取
# ============================================================

def ocr_image(img_path, lang_key="中文+英文"):
    """对图片做 OCR，返回字符串（已 clean_cjk + reflow_text）。"""
    p_lang, r_lang = LANG_MAP.get(lang_key, LANG_MAP["中文+英文"])
    if ENGINE == "rapidocr":
        ocr = RapidOCR(lang=r_lang)
        result, _ = ocr(img_path)
        if not result:
            return ""
        raw = "\n".join(line[1] for line in result)
    elif ENGINE == "tesseract":
        raw = pytesseract.image_to_string(_PILImage.open(img_path), lang=p_lang)
    else:
        raise RuntimeError("未安装任何 OCR 引擎（RapidOCR 或 pytesseract）。")
    return reflow_text(clean_cjk(raw))


def _ocr_pdf_page(fitz_doc, page_index, lang_key):
    page = fitz_doc[page_index]
    pix = page.get_pixmap(dpi=200)
    tmp = os.path.join(tempfile.gettempdir(), f"_pdf_page_{page_index + 1}.png")
    pix.save(tmp)
    try:
        raw = ocr_image(tmp, lang_key)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)
    return raw


def pdf_extract(pdf_path, lang_key="中文+英文"):
    """
    返回 doc = [{'marker','prose','tables'}, ...]。
    文本层优先；仅当文本层不足且未检出表格时，对该页回退 OCR（fitz 渲染后识别）。
    文本层结果不做 clean_cjk / reflow（保留原生排版）；仅 OCR 结果做清洗。
    """
    doc = []

    if pdfplumber:
        with pdfplumber.open(pdf_path) as pdf:
            n = len(pdf.pages)
            ocr_needed = []
            for i, page in enumerate(pdf.pages, 1):
                words = page.extract_words()
                tables, used = detect_tables_on_page(words)
                # 同步清单 §5.3：文本层字符 > 20 或检出表格 → 保留文本层
                if len(words) > 20 or len(tables) > 0:
                    prose = " ".join(w["text"] for idx, w in enumerate(words) if idx not in used)
                    doc.append({"marker": f"——— 第 {i} 页（文本层）———",
                                "prose": prose, "tables": tables})
                else:
                    ocr_needed.append(i)
                    doc.append(None)  # 占位，稍后回填
            if ocr_needed:
                if not fitz:
                    for i in ocr_needed:
                        doc[i - 1] = {"marker": f"——— 第 {i} 页（OCR）———",
                                      "prose": "（缺少 PyMuPDF，无法对无文本层页做 OCR）",
                                      "tables": []}
                else:
                    fdoc = fitz.open(pdf_path)
                    for i in ocr_needed:
                        raw = _ocr_pdf_page(fdoc, i - 1, lang_key)
                        doc[i - 1] = {"marker": f"——— 第 {i} 页（OCR）———",
                                      "prose": raw, "tables": []}
                    fdoc.close()
        return doc

    if fitz:
        fdoc = fitz.open(pdf_path)
        for i in range(fdoc.page_count):
            page = fdoc[i]
            raw_words = page.get_text("words")
            words = [{"text": w[4], "x0": w[0], "top": w[1], "x1": w[2],
                      "y1": w[3], "width": w[2] - w[0], "height": w[3] - w[1]}
                     for w in raw_words]
            tables, used = detect_tables_on_page(words)
            if len(raw_words) > 20 or len(tables) > 0:
                prose = " ".join(w[4] for idx, w in enumerate(raw_words) if idx not in used)
                doc.append({"marker": f"——— 第 {i + 1} 页（文本层）———",
                            "prose": prose, "tables": tables})
            else:
                raw = _ocr_pdf_page(fdoc, i, lang_key)
                doc.append({"marker": f"——— 第 {i + 1} 页（OCR）———",
                            "prose": raw, "tables": []})
        fdoc.close()
        return doc

    raise RuntimeError("需要 pdfplumber 或 PyMuPDF 才能解析 PDF。")


# ============================================================
# GUI（Tkinter，暗色主题，对齐网页版）
# ============================================================

BG = "#171d2b"
PANEL = "#1e2638"
INPUT_BG = "#0c1018"
FG = "#d7e2f4"
FG_DIM = "#9fb0c9"
ACCENT = "#5b8cff"
SECONDARY = "#cdd9ef"

# 输出格式 → 文件后缀（保存分割按钮文案与菜单共用）
FMT_EXT = {"TXT 文本": ".txt", "Markdown": ".md", "Excel（仅表格）": ".xlsx"}
IMG_EXTS = (".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp")


class RoundButton(tk.Canvas):
    """Canvas 自绘圆角按钮（radius=5，颜色恒定，不随窗口激活态变化）。

    Tkinter 原生按钮在 macOS 上有两个无法绕过的系统行为：
      1) 圆角由系统渲染，没有半径 API；
      2) 窗口激活 / 失活时按钮底色会被系统改掉（激活时发白）。
    因此用 Canvas 画圆角矩形自绘按钮，颜色完全自控。
    接口对齐 tk.Button 常用子集：config(text=/state=)、cget(text/state)、
    ["state"]、winfo_rootx/height（下拉菜单定位）。
    可选 command_caret：点击右侧 caret 区域时触发（用于「保存 .xx ▼」弹格式菜单），
    点主体区域仍触发 command。
    """

    R = 5            # 默认圆角半径；方形按钮按此固定，圆形按钮（square=True）覆盖为 (size-2)/2
    CARET_W = 22     # 右侧 caret 命中区宽度

    _PAL = {         # 每个 kind: n/h/d/x 四态，各为 (bg, fg, border)
        "accent": dict(n=("#5b8cff", "#ffffff", "#4a7df2"),
                       h=("#4a7df2", "#ffffff", "#6f96f5"),
                       d=("#4270d8", "#ffffff", "#3a6dd8"),
                       x=("#28395a", "#8a97b8", "#1e2a44")),
        "plain":  dict(n=("#1e2638", "#cdd9ef", "#2b344a"),
                       h=("#2a3550", "#ffffff", "#3a4a68"),
                       d=("#232e46", "#ffffff", "#1f2638"),
                       x=("#151d2d", "#55617a", "#1e2638")),
        "clear":  dict(n=("#1e2638", "#9fb0c9", "#2b344a"),
                       h=("#a04040", "#ffffff", "#b85450"),
                       d=("#8c3838", "#ffffff", "#8c3838"),
                       x=("#151d2d", "#55617a", "#1e2638")),
        # 圆形 icon 按钮（square+oval 用于预览区左侧工具条，对齐网页版 .zbtn 圆形）
        "icon":        dict(n=("#1e2638", "#cdd9ef", "#3a4a68"),
                            h=("#2a3550", "#ffffff", "#4c5f85"),
                            d=("#232e46", "#ffffff", "#2a3550"),
                            x=("#151d2d", "#55617a", "#1e2638")),
        "icon_danger": dict(n=("#1e2638", "#cdd9ef", "#3a4a68"),
                            h=("#a04040", "#ffffff", "#b85450"),
                            d=("#8c3838", "#ffffff", "#8c3838"),
                            x=("#151d2d", "#55617a", "#1e2638")),
        # 网页版 zoombar：半透明深底 + 白描边（近似色）→ 覆盖在图片上的工具条按钮
        "zbtn":        dict(n=("#1a2230", "#e8eefb", "#b9c3d6"),
                            h=("#24324a", "#ffffff", "#ffffff"),
                            d=("#12161e", "#cdd6e6", "#98a6c4"),
                            x=("#141a26", "#54617a", "#2c3648")),
        "zbtn_danger": dict(n=("#1a2230", "#e8eefb", "#b9c3d6"),
                            h=("#7a2f2f", "#ffffff", "#d98a86"),
                            d=("#5c2323", "#ffffff", "#ffffff"),
                            x=("#141a26", "#54617a", "#2c3648")),
    }

    def __init__(self, master, text="", command=None, kind="plain",
                 padx=12, pady=8, font=None, command_caret=None, square=False,
                 oval=None, zbtn=False):
        super().__init__(master, highlightthickness=0, bd=0, bg=master["bg"])
        import tkinter.font as tkfont
        self.command = command
        self.command_caret = command_caret
        self._pal = self._PAL[kind]
        self._padx, self._pady = padx, pady
        self._font = tkfont.Font(font=font or ("PingFang SC", 12))
        self._text = text
        self._state = tk.NORMAL
        self._hover = False
        self._down = False
        self._square = square
        if oval is None:
            oval = square            # 默认正方形按钮即画正圆（方形=直径）
        self._oval = oval
        self._zbtn = zbtn           # 预览区工具条按钮：加暗色外环防与图片混色
        self._r = self.R
        self.bind("<Button-1>", self._press)
        self.bind("<ButtonRelease-1>", self._release)
        self.bind("<Enter>", self._enter)
        self.bind("<Leave>", self._leave)
        self._redraw()

    # ---- 状态与配色 ----
    def _colors(self):
        if self._state == tk.DISABLED:
            return self._pal["x"]
        if self._down:
            return self._pal["d"]
        if self._hover:
            return self._pal["h"]
        return self._pal["n"]

    def _enter(self, _):
        if self._state == tk.NORMAL:
            self._hover = True
            self._redraw()

    def _leave(self, _):
        self._hover = False
        self._down = False
        self._redraw()

    def _press(self, _):
        if self._state == tk.NORMAL:
            self._down = True
            self._redraw()

    def _release(self, e):
        if not self._down:
            return
        self._down = False
        self._redraw()
        if not (0 <= e.x <= self.winfo_width() and 0 <= e.y <= self.winfo_height()):
            return
        if self.command_caret is not None and e.x >= self.winfo_width() - self.CARET_W:
            self.command_caret()
        elif self.command is not None:
            self.command()

    # ---- 绘制 ----
    def _round_rect(self, x1, y1, x2, y2, **kw):
        r = self._r
        pts = (x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y2 - r, x2, y2,
               x2 - r, y2, x1 + r, y2, x1, y2, x1, y2 - r, x1, y1 + r, x1, y1)
        return self.create_polygon(pts, smooth=True, **kw)

    def _redraw(self, *_):
        self.delete("all")
        tw = self._font.measure(self._text)
        th = self._font.metrics("linespace")
        caret = self.command_caret is not None
        w = tw + self._padx * 2 + 2 + (self.CARET_W if caret else 0)
        h = th + self._pady * 2 + 2
        if self._square:
            s = max(w, h)
            w = h = s
            self._r = (s - 2) // 2        # 圆形：radius 取最大
        else:
            self._r = self.R
        self.configure(width=w, height=h)
        bg, fg, bd = self._colors()
        if self._zbtn:
            # 覆盖在图片上的工具条按钮：暗色实心外盘 + 亮色描边 + 内部填充，
            # 无论图片是亮是暗都能与按钮形成对比、不会因颜色相近而「消失」
            self.create_oval(0, 0, w, h, fill="#0b0f17", outline="#0b0f17")
            self.create_oval(2, 2, w - 2, h - 2, fill=bg, outline=bd, width=1.5)
        elif self._oval:
            self.create_oval(1, 1, w - 1, h - 1, fill=bg, outline=bd, width=1)
        else:
            self._round_rect(1, 1, w - 1, h - 1, fill=bg, outline=bd, width=1)
        if caret:
            # 分割按钮：文字区与右侧 caret 区之间画一条竖分隔线，caret 箭头独立绘制在右区
            div = w - self.CARET_W
            self.create_line(div + 0.5, self._pady - 1, div + 0.5, h - self._pady + 1,
                             fill=bd, width=1)
            self.create_text((div + 1) // 2, h // 2, text=self._text, fill=fg,
                             font=self._font)
            self.create_text(w - self.CARET_W // 2, h // 2, text="▼", fill=fg,
                             font=self._font)
        else:
            self.create_text(w // 2, h // 2, text=self._text, fill=fg, font=self._font)

    # ---- tk.Button 兼容接口 ----
    def config(self, cnf=None, **kw):
        cnf = dict(cnf or {})
        cnf.update(kw)
        if "text" in cnf:
            self._text = cnf.pop("text")
        if "state" in cnf:
            self._state = (tk.DISABLED if cnf.pop("state") in (tk.DISABLED, "disabled")
                           else tk.NORMAL)
        if cnf:
            super().config(**cnf)
        self._redraw()

    def cget(self, key):
        if key == "text":
            return self._text
        if key == "state":
            return self._state
        return super().cget(key)

    def __getitem__(self, key):
        return self.cget(key)


class CanvasIcon:
    """直接画在父 Canvas 上的圆形图标按钮。

    与 RoundButton 不同，它不创建子 widget，因此没有方形的 widget 背景框；
    只有 Canvas 上的几何图形（圆环/文字），覆盖在图片上时不会显示任何矩形底。
    """
    SIZE = 34          # 图标直径
    HALO = "#0b0f17"   # 深色外环（防与图片混色）

    def __init__(self, canvas, x, y, text, command, kind="zbtn", on_active=None):
        self.canvas = canvas
        self.command = command
        self.on_active = on_active
        self._pal = RoundButton._PAL[kind]
        self._text = text
        self._hover = False
        self._down = False
        self._state = tk.NORMAL
        r = self.SIZE / 2
        # 外环（深色实心圆，覆盖 widget 不可能做到的「无边框」需求）
        self._halo = canvas.create_oval(x - r - 2, y - r - 2,
                                      x + r + 2, y + r + 2,
                                      fill=self.HALO, outline="", state="hidden")
        # 内圆（填充 + 描边）
        self._face = canvas.create_oval(x - r, y - r, x + r, y + r,
                                        outline="", fill="", state="hidden")
        # 文字
        self._label = canvas.create_text(x, y, text=text,
                                         font=("PingFang SC", 12, "bold"),
                                         state="hidden")
        self._ids = (self._halo, self._face, self._label)
        self._bind()
        self._redraw()

    def _bind(self):
        for i in self._ids:
            self.canvas.tag_bind(i, "<Enter>", self._enter)
            self.canvas.tag_bind(i, "<Leave>", self._leave)
            self.canvas.tag_bind(i, "<ButtonPress-1>", self._press)
            self.canvas.tag_bind(i, "<ButtonRelease-1>", self._release)
            self.canvas.tag_bind(i, "<Motion>", self._active)

    def _colors(self):
        if self._state != tk.NORMAL:
            return self._pal["x"]
        if self._down:
            return self._pal["d"]
        if self._hover:
            return self._pal["h"]
        return self._pal["n"]

    def _enter(self, e):
        if self._state == tk.NORMAL:
            self._hover = True
            self._redraw()
        self._active(e)

    def _leave(self, e):
        self._hover = False
        self._down = False
        self._redraw()

    def _press(self, e):
        if self._state == tk.NORMAL:
            self._down = True
            self._redraw()
        self._active(e)
        return "break"            # 阻止 canvas 的 <ButtonPress-1> 同时触发平移

    def _release(self, e):
        if not self._down:
            return
        self._down = False
        self._redraw()
        x1, y1, x2, y2 = self.canvas.bbox(self._halo)
        if x1 <= e.x <= x2 and y1 <= e.y <= y2 and self.command:
            self.command()
        return "break"

    def _active(self, _):
        if self.on_active:
            self.on_active()

    def _redraw(self):
        bg, fg, bd = self._colors()
        self.canvas.itemconfig(self._halo, fill=self.HALO)
        self.canvas.itemconfig(self._face, fill=bg, outline=bd, width=1.5)
        self.canvas.itemconfig(self._label, fill=fg)

    def place(self, cx, cy):
        r = self.SIZE / 2
        self.canvas.coords(self._halo, cx - r - 2, cy - r - 2,
                           cx + r + 2, cy + r + 2)
        self.canvas.coords(self._face, cx - r, cy - r, cx + r, cy + r)
        self.canvas.coords(self._label, cx, cy)

    def show(self, visible=True):
        st = "normal" if visible else "hidden"
        for i in self._ids:
            self.canvas.itemconfig(i, state=st)

    def hide(self):
        self.show(False)

    def config(self, state=None):
        if state is not None:
            self._state = state
            self._redraw()


# 状态行文案（显示在预览区下方，进度圈图标在右侧）
DEFAULT_STATUS = "首次使用会联网下载识别引擎与语言包。"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("图片/PDF文字解析工具")
        self.geometry("1000x640")
        self.minsize(860, 540)
        self.configure(bg=BG)

        self.doc = []
        self.path = tk.StringVar()
        self.lang_var = tk.StringVar(value="中文+英文")
        self.fmt_var = tk.StringVar(value="TXT 文本")
        self._photo = None
        self._orig_img = None      # 预览原图（PIL.Image）
        self._pdf_name = None      # 已选 PDF 的文件名（预览居中显示）
        self._zoom = 1.0           # 预览缩放系数
        self._img_id = None        # 预览图 canvas item
        self._off = [0, 0]         # 图片平移偏移（图片左上角坐标）
        self._pan_last = None      # 拖拽上一点
        self._spin_active = False  # 进度圈是否正在旋转
        self._spin_id = None
        self._spin_arc = None
        self._run_ok = True
        self._tb_timer = None      # 工具条 3 秒自动隐藏计时器

        self._build_header()
        self._build_toolbar()
        self._build_main()

        self._layout_canvas()
        self._show_out_placeholder()
        self._set_running(False)
        if ENGINE is None:
            self.status.config(
                text="未检测到 OCR 引擎（RapidOCR / pytesseract）——有文本层的 PDF 仍可正常解析；"
                     "如需识别图片/扫描件请 pip install pytesseract")

    # -------- 布局 --------
    def _btn(self, parent, text, command, accent=False, command_caret=None):
        return RoundButton(parent, text=text, command=command,
                           kind="accent" if accent else "plain",
                           command_caret=command_caret)

    def _pill(self, parent, text, bg, fg):
        """网页版风格的胶囊标签（圆角无法实现，用色块近似）。"""
        return tk.Label(parent, text=text, bg=bg, fg=fg,
                        padx=9, pady=3, font=("PingFang SC", 10))

    def _build_header(self):
        h = tk.Frame(self, bg=BG)
        h.pack(fill="x", padx=20, pady=(11, 2))
        tk.Label(h, text="图片/PDF文字解析工具", bg=BG, fg="#eaf0fd",
                 font=("PingFang SC", 16, "bold")).pack(side="left")
        self._pill(h, "本地运行 · 文件不出本机", "#0f2f22", "#38d39a") \
            .pack(side="left", padx=(12, 6))
        self._pill(h, f"v{__VERSION__}", "#12264d", "#5b9dff").pack(side="left")

    def _build_toolbar(self):
        # 顺序与形态对齐网页版：识别 → 语言 ▼ → 复制 → 保存 .xx ▼
        bar = tk.Frame(self, bg=BG)
        bar.pack(fill="x", padx=20, pady=(8, 0))
        self.btn_run = self._btn(bar, "🔍 识别", self.run, accent=True)
        self.btn_run.pack(side="left")

        # 语言下拉按钮：用普通 Button 弹出菜单（与「保存 ▼」一致），
        # 避免 Menubutton 与 Button 度量差异导致整行按钮高度不齐
        self.lang_btn = self._btn(bar, "中文+英文 ▼", self._open_lang_menu)
        self.lang_menu = tk.Menu(self.lang_btn, tearoff=0, bg=INPUT_BG, fg=FG,
                                 activebackground=ACCENT, activeforeground="#fff", bd=0)
        for k in LANG_MAP:
            self.lang_menu.add_radiobutton(label=k, variable=self.lang_var,
                                           command=self._sync_lang)
        self.lang_btn.pack(side="left", padx=(10, 0))

        self.btn_copy = self._btn(bar, "⧉ 复制", self.copy)
        self.btn_copy.pack(side="left", padx=(10, 0))

        # 保存：合并式按钮「保存 .xx | ▼」——点主体直接保存，点右侧 ▼ 弹格式菜单，
        # 前箭头已去掉；文字与右侧箭头之间由 RoundButton 绘制一条竖分隔线（见 _redraw）
        self.btn_save = self._btn(bar, "保存 .txt", self.save,
                                  command_caret=self._open_fmt_menu)
        self.btn_save.pack(side="left", padx=(10, 0))
        self.fmt_menu = tk.Menu(self.btn_save, tearoff=0, bg=INPUT_BG, fg=FG,
                                activebackground=ACCENT, activeforeground="#fff", bd=0)
        for label in FMT_EXT:
            self.fmt_menu.add_radiobutton(label=label, variable=self.fmt_var,
                                          value=label, command=self._sync_fmt)

    def _sync_lang(self):
        self.lang_btn.config(text=self.lang_var.get() + " ▼")

    def _sync_fmt(self):
        self.btn_save.config(text="保存 " + FMT_EXT[self.fmt_var.get()])
        self._refresh_view()

    def _open_lang_menu(self):
        try:
            self.lang_menu.tk_popup(self.lang_btn.winfo_rootx(),
                                    self.lang_btn.winfo_rooty()
                                    + self.lang_btn.winfo_height())
        finally:
            self.lang_menu.grab_release()

    def _open_fmt_menu(self):
        try:
            self.fmt_menu.tk_popup(
                self.btn_save.winfo_rootx() + self.btn_save.winfo_width() - 24,
                self.btn_save.winfo_rooty() + self.btn_save.winfo_height())
        finally:
            self.fmt_menu.grab_release()

    def _build_main(self):
        main = tk.Frame(self, bg=BG)
        main.pack(fill="both", expand=True, padx=20, pady=(3, 4))
        main.grid_columnconfigure(0, weight=1, uniform="c")
        main.grid_columnconfigure(1, weight=1, uniform="c")
        main.grid_rowconfigure(0, weight=1)

        # 左：预览画布（虚线边框；空态中央「选择文件」按钮，有图态可缩放/平移）
        self.canvas = tk.Canvas(main, bg=PANEL, highlightthickness=0, bd=0)
        self.canvas.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        self.canvas.bind("<Configure>", self._on_canvas_resize)
        self.canvas.bind("<ButtonPress-1>", self._pan_start)
        self.canvas.bind("<B1-Motion>", self._pan_move)
        self.canvas.bind("<ButtonRelease-1>", lambda e: setattr(self, "_pan_last", None))
        # 工具条随鼠标活动显示：预览区内移动即显示并重置 3 秒自动隐藏计时
        self.canvas.bind("<Motion>", self._tb_poke)
        self.canvas.bind("<Enter>", self._tb_poke)
        self.canvas.bind("<ButtonPress-1>", self._tb_poke, add="+")
        # 滚轮 / 触控板手势：
        #   ⌘（macOS）/ Ctrl（X11）+ 滚轮 → 缩放（单独绑定，确保一定触发缩放，不依赖 state 位判断）
        #   默认滚轮 = 垂直平移；Shift + 滚轮 = 水平平移；平移方向跟随系统「自然滚动」设置
        #   macOS 双指捏合 → <Magnify>（Tk 支持时）
        self.canvas.bind("<MouseWheel>", self._on_wheel)
        self.canvas.bind("<Shift-MouseWheel>", self._on_wheel_h)
        self.canvas.bind("<Command-MouseWheel>", self._on_zoom_wheel)  # macOS ⌘+滚轮
        self.canvas.bind("<Control-MouseWheel>", self._on_zoom_wheel)  # X11 Ctrl+滚轮
        # macOS 双指捏合 / 双击缩放：Tk 在 Aqua 上把这些手势以虚拟事件（<<Magnify>> /
        # <<SmartMagnify>>）抛出，<Magnify> 直写形式在部分版本的 Tk 上不触发，
        # 因此两种形式都绑定；非 macOS 平台没有这些虚拟事件，逐个 try 捕获避免报错。
        for ev in ("<<Magnify>>", "<Magnify>"):
            try:
                self.canvas.bind(ev, self._on_magnify)
            except tk.TclError:
                pass
        for ev in ("<<SmartMagnify>>", "<SmartMagnify>"):
            try:
                self.canvas.bind(ev, self._on_smart_magnify)
            except tk.TclError:
                pass

        self.btn_pick = RoundButton(self.canvas, text="📁 选择文件", command=self.pick,
                                    padx=18, pady=8, font=("PingFang SC", 12))
        self.pick_win = self.canvas.create_window(0, 0, window=self.btn_pick)
        self.hint_id = self.canvas.create_text(0, 0, state="hidden",
            text="支持图片（PNG / JPG / WebP 等）与 PDF 文件",
            fill=FG_DIM, font=("PingFang SC", 10), justify="center", width=300)

        # 左侧竖排圆形工具条（图片态显示，仿网页版 .zoombar）：
        # 放大 / 缩小 / 重置 / 上传 / 关闭；无背景框，仅 5 个圆形按钮；3 秒无操作自动隐藏
        self._tb_icons = []      # CanvasIcon 实例列表
        tb_specs = [
            ("＋", self.zoom_in, "zbtn"),
            ("－", self.zoom_out, "zbtn"),
            ("↻", self.zoom_reset, "zbtn"),
            ("↑", self.pick, "zbtn"),
            ("✕", self.clear, "zbtn_danger"),
        ]
        for text, cmd, kind in tb_specs:
            icon = CanvasIcon(self.canvas, 0, 0, text, cmd, kind=kind,
                              on_active=self._tb_poke)
            self._tb_icons.append(icon)

        # 右：结果文本框（暗色滚动条，对齐网页版 v1.2.19）
        outw = tk.Frame(main, bg=INPUT_BG)
        outw.grid(row=0, column=1, sticky="nsew")
        outw.grid_rowconfigure(0, weight=1)
        outw.grid_columnconfigure(0, weight=1)
        # 结果区 Text：focus ring（高亮边框）在 macOS 上选中窗口时发白，
        # 把 highlight 三色全部固定为结果区背景色 → 任何状态下边框恒定
        self.out = tk.Text(outw, bg=INPUT_BG, fg=FG, insertbackground=FG,
                           relief="flat", font=("PingFang SC", 12), wrap="word",
                           bd=0, padx=10, pady=8, undo=True,
                           highlightthickness=1, highlightbackground=INPUT_BG,
                           highlightcolor=INPUT_BG)
        # 滚动条：macOS Aqua 主题会忽略 bg/troughcolor，必须用 ttk + clam 主题；
        # 配色对齐网页版（thumb #3a4a68 / hover #4c5f85 / trough 与结果区同色）
        try:
            ttk.Style(self).theme_use("clam")
        except Exception:
            pass
        ttk.Style(self).configure("Dark.Vertical.TScrollbar",
                                  background="#3a4a68", troughcolor=INPUT_BG,
                                  bordercolor=INPUT_BG, lightcolor="#3a4a68",
                                  darkcolor="#3a4a68", gripcount=0,
                                  arrowsize=12, width=10)
        ttk.Style(self).map("Dark.Vertical.TScrollbar",
                            background=[("active", "#4c5f85"),
                                        ("pressed", "#4c5f85")])
        sb = self.sb = ttk.Scrollbar(outw, command=self.out.yview,
                                     style="Dark.Vertical.TScrollbar")
        self.out.config(yscrollcommand=sb.set)
        self.out.grid(row=0, column=0, sticky="nsew")
        sb.grid(row=0, column=1, sticky="ns")

        # 状态行（识别按钮下方的状态文案，移到此处 = 底部，横跨两列；
        # 文案/进度条右侧与右侧识别结果区的右边缘对齐，所有状态/报错都在此显示）
        status_area = tk.Frame(main, bg=BG)
        status_area.grid(row=1, column=0, columnspan=2, sticky="ew",
                         pady=(8, 0))
        self.status = tk.Label(status_area, text=DEFAULT_STATUS,
                               bg=BG, fg=FG_DIM, anchor="w",
                               font=("PingFang SC", 12))
        self.status.pack(side="left", fill="x", expand=True)
        self.prog = tk.Canvas(status_area, width=18, height=18, bg=BG,
                              highlightthickness=0)
        self.prog.pack(side="right", padx=(6, 0))

        self.out.tag_config("ph", foreground=FG_DIM)

    # -------- 交互 --------
    def pick(self):
        f = filedialog.askopenfilename(
            filetypes=[("图片/PDF", "*.png *.jpg *.jpeg *.bmp *.tif *.tiff *.webp *.pdf")])
        if f:
            self.path.set(f)
            self._pdf_name = None
            self._show_preview(f)
            self._set_running(False)   # 选完文件回到「未运行」态，识别按钮立即可用

    def clear(self):
        self.path.set("")
        self.doc = []
        self._orig_img = None
        self._pdf_name = None
        self._zoom = 1.0
        self._layout_canvas()
        self._show_out_placeholder()
        self._set_running(False)
        self.prog.delete("all")                 # 清除状态图标
        self.status.config(text=DEFAULT_STATUS)

    def _show_preview(self, path):
        self._orig_img = None
        self._zoom = 1.0
        self._off = [0, 0]
        if HAVE_PIL and path.lower().endswith(IMG_EXTS):
            try:
                self._orig_img = Image.open(path).copy()
                self._draw_image()
                self._layout_canvas()
                return
            except Exception:
                self._orig_img = None
        self._pdf_name = os.path.basename(path)
        self._layout_canvas()

    # -------- ⑧ 缩放 / 平移 --------
    def _draw_image(self):
        """按当前 self._zoom 重绘预览图（重绘后重新居中）。"""
        w, h = self._orig_img.size
        nw, nh = max(1, int(w * self._zoom)), max(1, int(h * self._zoom))
        try:
            resample = Image.LANCZOS
        except AttributeError:   # Pillow 旧版本
            resample = Image.ANTIALIAS
        img = self._orig_img.resize((nw, nh), resample)
        self._photo = ImageTk.PhotoImage(img)
        if self._img_id is not None:
            self.canvas.delete(self._img_id)
        self._img_id = self.canvas.create_image(0, 0, anchor="nw", image=self._photo)
        # 圆形工具条在初始化时创建，新图片会覆盖在上面，因此每次重绘后都将其提升到最上层
        for icon in self._tb_icons:
            for i in icon._ids:
                self.canvas.tag_raise(i)
        self._center_image()

    def _center_image(self):
        """把图片放到画布中央（图片大于画布时留 8px 边距，可拖动平移）。"""
        if self._img_id is None:
            return
        cw = max(self.canvas.winfo_width(), 200)
        ch = max(self.canvas.winfo_height(), 160)
        self._off = [max((cw - self._photo.width()) // 2, 8),
                     max((ch - self._photo.height()) // 2, 8)]
        self.canvas.coords(self._img_id, self._off[0], self._off[1])

    def _layout_canvas(self):
        """统一管理预览画布两种状态：选择态（空态 / 已选 PDF，中央「选择文件」按钮）/ 图片态。"""
        c = self.canvas
        cw = max(c.winfo_width(), 400)
        ch = max(c.winfo_height(), 320)
        has_img = self._orig_img is not None
        if not has_img and self._img_id is not None:
            c.delete(self._img_id)      # 清掉残留图片（清除 / 换成 PDF 时）
            self._img_id = None
        # 虚线边框（对齐网页版预览框）
        c.delete("border")
        c.create_rectangle(5, 5, cw - 5, ch - 5, dash=(5, 4),
                           outline="#3a4a68", tags="border")
        c.tag_lower("border")
        if has_img:
            c.itemconfig(self.pick_win, state="hidden")
            c.itemconfig(self.hint_id, state="hidden")
            # 左侧竖排圆形工具条（放大/缩小/重置/上传/关闭）：无背景框，仅 5 个圆形按钮，3 秒自动隐藏
            # 无背景框：5 个 CanvasIcon 圆形按钮垂直居中，贴在预览区左侧
            size = CanvasIcon.SIZE
            gap = 8
            n = len(self._tb_icons)
            total = n * size + (n - 1) * gap
            top = ch // 2 - total // 2
            cx = 12 + size // 2
            for i, icon in enumerate(self._tb_icons):
                cy = top + i * (size + gap) + size // 2
                icon.place(cx, cy)
                icon.config(state=tk.NORMAL)
                icon.show(True)
            self._tb_poke()
            st = tk.NORMAL
        else:
            # 空态 / PDF 态统一为「选择态」：中央「选择文件」按钮 + 下方提示文案，
            # 不显示删除按钮（点「选择文件」即可重新上传新文件，与网页版一致）
            if self._pdf_name:
                hint_text = f"已选择 PDF：\n{self._pdf_name}"
            else:
                hint_text = "支持图片（PNG / JPG / WebP 等）与 PDF 文件"
            c.itemconfig(self.hint_id, text=hint_text)
            c.itemconfig(self.pick_win, state="normal")
            c.itemconfig(self.hint_id, state="normal")
            self._tb_hide()
            self._tb_stop_timer()
            c.coords(self.pick_win, cw / 2, ch / 2 - 16)
            c.coords(self.hint_id, cw / 2, ch / 2 + 26)
            st = tk.DISABLED
        for icon in self._tb_icons:
            icon.config(state=st)

    def _on_canvas_resize(self, e):
        self._layout_canvas()
        if self._img_id is not None:
            self._center_image()

    def _set_zoom(self, z):
        if self._orig_img is None:
            return
        self._zoom = max(0.2, min(8.0, z))
        self._draw_image()

    def zoom_in(self):
        self._set_zoom(self._zoom * 1.25)

    def zoom_out(self):
        self._set_zoom(self._zoom / 1.25)

    def zoom_reset(self):
        self._set_zoom(1.0)

    def _pan_start(self, e):
        if self._img_id is None:
            return
        self._pan_last = (e.x, e.y)
        self.canvas.config(cursor="fleur")

    def _pan_move(self, e):
        if self._img_id is None or not self._pan_last:
            return
        self._off[0] += e.x - self._pan_last[0]
        self._off[1] += e.y - self._pan_last[1]
        self._pan_last = (e.x, e.y)
        self.canvas.coords(self._img_id, self._off[0], self._off[1])

    # -------- 滚轮 / 触控板手势 --------
    # macOS：双指滚动 / 捏合均从 <MouseWheel> 进入（无事件细分）；系统触控板手势本身不带
    # 修饰键，Tk 报告的 delta：双指滚动每帧 ±1~10，捏合每次可达数十上百。据此判定：
    #   Ctrl(bit2=0x4) / Cmd(bit4=0x10) 修饰 → 缩放；macOS 无修饰但 |delta| 大 → 捏合缩放；
    #   其余 → 平移。X11：滚轮每格 ±120 → 平移；Ctrl+滚轮 → 缩放。
    # 平移方向跟随系统「自然滚动」设置（见模块顶部 _NATURAL）：
    #   「自然滚动」开启 → 图片跟随手指方向移动；关闭 → 传统方向。
    @staticmethod
    def _wheel_step(d):
        """统一滚动量：macOS 触控板 delta 通常 ±1~10（=像素级）；X11 滚轮 ±120（=格）。"""
        return d * 8 if abs(d) < 20 else d / 4

    def _pan_amt(self, d):
        """换算为像素并应用系统滚动方向：自然滚动 ON 时 delta 直接等同手指位移。"""
        step = self._wheel_step(d)
        if _IS_MAC and not _NATURAL:
            step = -step                      # 传统滚动：图片反向于手指
        return step

    def _do_zoom(self, d):
        """滚轮缩放：每「一个滚轮格」= ×1.25（与工具栏 ＋ 步进一致）。
        X11/Win 一格 delta=±120；macOS 每事件即为一格（delta 通常 ±1~10，
        以前除以 360 导致一格缩放 1.15^(1/360)≈1.0004 近乎不可见，现直接按格计）。"""
        if d == 0:
            return
        if not _IS_MAC and abs(d) >= 40:
            steps = d / 120.0          # X11/Win：一格 = ±120
        else:
            steps = 1.0 if d > 0 else -1.0   # macOS：每个事件 = 一个滚轮格
        self._set_zoom(self._zoom * (1.25 ** steps))

    def _pan_canvas(self, dx, dy):
        if self._img_id is None:
            return
        self._off[0] += dx
        self._off[1] += dy
        self.canvas.coords(self._img_id, self._off[0], self._off[1])

    def _on_zoom_wheel(self, e):
        """⌘（macOS）/ Ctrl（X11）+ 滚轮 → 缩放。单独绑定，必定触发，不依赖 state 位判断。"""
        if self._img_id is None:
            return
        self._tb_poke()
        self._do_zoom(e.delta)

    def _on_wheel(self, e):          # 垂直平移；macOS 无修饰且 delta 远超双指滚动 → 视为捏合缩放
        if self._img_id is None:
            return
        self._tb_poke()
        st = (e.state or 0)
        if _IS_MAC and not (st & (0x4 | 0x8 | 0x10)) and abs(e.delta) > 20:
            self._do_zoom(e.delta)   # 无修饰 + delta 远超双指滚动（±1~10）→ 捏合缩放
        else:
            self._pan_canvas(0, -self._pan_amt(e.delta))   # 垂直平移

    def _on_wheel_h(self, e):        # Shift + 滚轮 = 水平平移
        if self._img_id is None:
            return
        self._tb_poke()
        self._pan_canvas(-self._pan_amt(e.delta), 0)

    def _on_magnify(self, e):        # macOS 双指捏合（虚拟 <<Magnify>>，delta 即本次捏合增量）
        if self._img_id is None:
            return
        self._tb_poke()
        d = getattr(e, "delta", 0) or 0
        # 每帧增量：正=放大。单帧 clamp 0.5~2 防止极端跳变，总缩放由 _set_zoom 限制在 20%~800%
        f = 1.0 + d
        f = max(0.5, min(2.0, f))
        self._set_zoom(self._zoom * f)

    def _on_smart_magnify(self, e):  # macOS 双指双击「智能缩放」：在 100% 与 200% 间切换
        if self._img_id is None:
            return
        self._tb_poke()
        self._set_zoom(2.0 if abs(self._zoom - 1.0) < 0.01 else 1.0)

    # ---- 工具条（5 个圆形按钮）显隐 ----
    def _tb_show(self):
        for icon in self._tb_icons:
            icon.show(True)

    def _tb_hide(self):
        for icon in self._tb_icons:
            icon.show(False)

    def _tb_poke(self, *_):
        """有操作（移动/点击/滚动）→ 显示工具条并重置 3 秒自动隐藏计时。"""
        if self._orig_img is None:
            return
        self._tb_show()
        if self._tb_timer:
            self.after_cancel(self._tb_timer)
        self._tb_timer = self.after(3000, self._tb_expire)

    def _tb_expire(self):
        self._tb_timer = None
        if self._orig_img is not None and not self._pan_last:
            self._tb_hide()

    def _tb_stop_timer(self):
        if self._tb_timer:
            try:
                self.after_cancel(self._tb_timer)
            except Exception:
                pass
            self._tb_timer = None

    # -------- ⑩ 进度圈 / 对钩动画（状态行 18×18 图标） --------
    def _start_spin(self):
        self._spin_active = True
        self._spin_angle = 0
        self.prog.delete("all")
        self._spin_arc = self.prog.create_arc(2, 2, 16, 16, style="arc",
                                              width=2.5, outline=ACCENT,
                                              start=0, extent=270)
        self._spin_step()

    def _spin_step(self):
        if not self._spin_active:
            return
        self._spin_angle = (getattr(self, "_spin_angle", 0) + 30) % 360
        self.prog.itemconfig(self._spin_arc, start=self._spin_angle, extent=270)
        self._spin_id = self.prog.after(80, self._spin_step)

    def _stop_spin(self, ok=True):
        if not self._spin_active:
            return
        self._spin_active = False
        if self._spin_id:
            try:
                self.prog.after_cancel(self._spin_id)
            except Exception:
                pass
            self._spin_id = None
        self.prog.delete("all")
        if ok:
            # 对钩（绿色）
            self.prog.create_line(4, 9, 7, 13, 13, 6, width=2.5,
                                  fill="#46d17a", capstyle="round", joinstyle="round")
        else:
            # 叉（红色）
            self.prog.create_line(4, 4, 14, 14, width=2.5, fill="#ff6b6b", capstyle="round")
            self.prog.create_line(14, 4, 4, 14, width=2.5, fill="#ff6b6b", capstyle="round")
        self.prog.after(2000, self.prog.delete, "all")

    def _set_running(self, running):
        if not running:
            self._stop_spin(self._run_ok)   # 结束时把进度圈转成对钩 / 叉
        has_file = bool(self.path.get().strip())
        has_doc = bool(self.doc)
        self.btn_run.config(
            state=tk.NORMAL if (has_file and not running) else tk.DISABLED)
        # 保存（合并式「保存 .xx ▼」）/ 复制：有结果才可用（与网页版「识别后才能下载」一致）
        for b in (self.btn_save, self.btn_copy):
            b.config(state=tk.NORMAL if (has_doc and not running) else tk.DISABLED)
        self.btn_pick.config(state=tk.DISABLED if running else tk.NORMAL)
        self.lang_btn.config(state=tk.DISABLED if running else tk.NORMAL)
        # 工具条内 5 个按钮在图片态可用，否则禁用
        for icon in self._tb_icons:
            icon.config(state=tk.NORMAL if (self._orig_img is not None and not running) else tk.DISABLED)

    def run(self):
        p = self.path.get().strip()
        if not p or not os.path.exists(p):
            messagebox.showerror("错误", "请先选择一个存在的文件。")
            return
        self._set_running(True)
        self._run_ok = True
        self._start_spin()                  # ⑩ 开始转圈
        self.out.delete("1.0", tk.END)
        self.status.config(text="正在识别，请稍候…")
        lang = self.lang_var.get()

        def work():
            try:
                if p.lower().endswith(".pdf"):
                    self.doc = pdf_extract(p, lang)
                else:
                    raw = ocr_image(p, lang)
                    self.doc = [{"marker": "", "prose": raw, "tables": []}]
                self._run_ok = True
                self.after(0, self._refresh_view)
                self.after(0, lambda: self.status.config(text="识别完成，结果已显示在右侧。"))
            except Exception as e:
                self._run_ok = False
                self.doc = []
                self.after(0, lambda: self.out.delete("1.0", tk.END))
                self.after(0, lambda: self.out.insert("1.0", "出错：" + str(e)))
                self.after(0, lambda: self.status.config(text="识别失败：" + str(e)))
            finally:
                self.after(0, self._set_running, False)

        threading.Thread(target=work, daemon=True).start()

    def _show_out_placeholder(self):
        self.out.delete("1.0", tk.END)
        self.out.insert("1.0", "识别结果将显示在这里…", "ph")

    def _refresh_view(self):
        if not self.doc:
            return
        fmt = self.fmt_var.get()
        if fmt == "Excel（仅表格）":
            text = render_md_tables_only(self.doc)
        elif fmt == "Markdown":
            text = render_md(self.doc) or "（未提取到文本）"
        else:
            text = render_txt(self.doc) or "（未提取到文本）"
        self.out.delete("1.0", tk.END)
        self.out.insert("1.0", text)

    def copy(self):
        if not self.doc:
            return
        text = self.out.get("1.0", tk.END).strip()
        if not text:
            return
        self.clipboard_clear()
        self.clipboard_append(text)
        self.status.config(text="已复制到剪贴板，可直接粘贴到其他应用。")

    def save(self):
        if not self.doc:
            messagebox.showinfo("提示", "请先识别出内容再保存。")
            return
        fmt = self.fmt_var.get()
        base = os.path.splitext(os.path.basename(self.path.get()))[0] or "ocr_result"

        if fmt == "Excel（仅表格）":
            tables = [t for p in self.doc for t in p.get("tables", [])]
            if not tables:
                messagebox.showinfo("提示", "未检测到表格，已改为导出 TXT。")
                suffix = ".txt"
                content = render_txt(self.doc)
            else:
                try:
                    f = filedialog.asksaveasfilename(initialfile=base + ".xlsx",
                                                     defaultextension=".xlsx",
                                                     filetypes=[("Excel", "*.xlsx")])
                    if not f:
                        return
                    build_xlsx(self.doc, f)
                    self.status.config(text="已保存：" + f)
                except Exception as e:
                    messagebox.showerror("导出失败", str(e))
                return
        elif fmt == "Markdown":
            suffix = ".md"
            content = render_md(self.doc)
        else:
            suffix = ".txt"
            content = render_txt(self.doc)

        f = filedialog.asksaveasfilename(initialfile=base + suffix,
                                         defaultextension=suffix,
                                         filetypes=[("文本", suffix)])
        if not f:
            return
        with open(f, "w", encoding="utf-8") as fh:
            fh.write(content)
        self.status.config(text="已保存：" + f)


if __name__ == "__main__":
    App().mainloop()
