# =============================================================================
# Project Relay v2 - 統合報告パワーポイント自動生成アプリ（UX極限改善版）
# 向平氏専用 業務効率化ツール（Streamlit）
# =============================================================================
#
# 【必要なライブラリのインストールコマンド】
# pip install streamlit python-pptx openpyxl pdfplumber
#
# 【起動コマンド】
# streamlit run app.py
# =============================================================================

import io
import time
from datetime import datetime
from pathlib import Path

import streamlit as st

# --- サードパーティライブラリ（未インストール時もクラッシュしない設計） ---
try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
    PPTX_AVAILABLE = True
except ImportError:
    PPTX_AVAILABLE = False

try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

try:
    import pdfplumber
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False


# =============================================================================
# ページ設定（最初に呼ぶ必要があります）
# =============================================================================
st.set_page_config(
    page_title="Project Relay | 統合レポート生成",
    page_icon="⬡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =============================================================================
# 定数
# =============================================================================
HISTORY_DIR = Path("./history")
HISTORY_DIR.mkdir(exist_ok=True)

CATEGORY_KEYWORDS = {
    "今月の成果": [
        "成果", "達成", "完了", "リリース", "ローンチ", "公開", "獲得", "受注",
        "契約", "成功", "実施", "完成", "提供", "展開", "運用開始",
    ],
    "数値指標": [
        "売上", "収益", "利益", "コスト", "費用", "予算", "KPI", "目標", "達成率",
        "前月比", "前年比", "増加", "減少", "%", "万円", "千件", "PV", "CVR",
        "ROI", "CPA", "CPC", "クリック率", "転換率", "件数", "数",
    ],
    "発生した課題": [
        "課題", "問題", "障害", "遅延", "バグ", "エラー", "リスク", "懸念",
        "未達", "不足", "改善が必要", "検討が必要", "対応中", "調査中", "ペンディング",
    ],
    "次月の予定": [
        "予定", "計画", "スケジュール", "来月", "次月", "今後", "方針", "施策",
        "実施予定", "リリース予定", "検討予定", "対応予定", "目標設定",
    ],
}

CATEGORY_ICONS = {
    "今月の成果":       "🏆",
    "数値指標":         "📊",
    "発生した課題":     "⚠️",
    "次月の予定":       "📅",
    "その他・参考情報": "📎",
}

# PPTX カラー定義
if PPTX_AVAILABLE:
    C_DARK   = RGBColor(0x1E, 0x27, 0x61)
    C_ACCENT = RGBColor(0xCA, 0xDC, 0xFC)
    C_WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
    C_LIGHT  = RGBColor(0xF4, 0xF6, 0xFF)
    C_BODY   = RGBColor(0x1E, 0x27, 0x61)
    C_CITE   = RGBColor(0x99, 0x99, 0xAA)


# =============================================================================
# グローバル CSS
# =============================================================================
CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Serif+JP:wght@300;400;600&family=Noto+Sans+JP:wght@300;400;500;700&family=Cormorant+Garamond:ital,wght@0,300;0,600;1,300&display=swap');

:root {
    --navy:      #080F24;
    --navy-2:    #0D1B3E;
    --navy-3:    #1A2848;
    --gold:      #C9A84C;
    --gold-lt:   #E8C97A;
    --off-white: #EEF1F8;
    --muted:     #6B7A9F;
    --border:    rgba(201,168,76,0.20);
    --card:      #111D38;
    --success:   #22c55e;
    --white:     #FFFFFF;
}

/* ── Reset & base ── */
html, body,
[data-testid="stAppViewContainer"],
[data-testid="stMain"] {
    background-color: var(--navy) !important;
    color: var(--off-white) !important;
    font-family: 'Noto Sans JP', sans-serif;
}
[data-testid="stHeader"]        { background: transparent !important; }
[data-testid="stVerticalBlock"] { gap: 0 !important; }
.block-container                { padding: 0 !important; max-width: 100% !important; }
.appview-container .main .block-container { padding-top: 0 !important; }

/* ── Sidebar ── */
[data-testid="stSidebar"] {
    background: #04091A !important;
    border-right: 1px solid var(--border) !important;
}
[data-testid="stSidebarContent"] { padding: 0 !important; }

/* ── Hero ── */
.hero {
    background: linear-gradient(140deg, #04091A 0%, #0D1B3E 55%, #162448 100%);
    border-bottom: 1px solid var(--border);
    padding: 48px 72px 40px;
    position: relative; overflow: hidden;
}
.hero::before {
    content: ''; position: absolute; top: -80px; right: -80px;
    width: 380px; height: 380px; border-radius: 50%;
    background: radial-gradient(circle, rgba(201,168,76,0.07) 0%, transparent 68%);
    pointer-events: none;
}
.hero::after {
    content: ''; position: absolute; bottom: 0; left: 0; right: 0; height: 1px;
    background: linear-gradient(90deg, transparent 0%, var(--gold) 50%, transparent 100%);
}
.hero-eyebrow {
    font-family: 'Cormorant Garamond', serif;
    font-size: 12px; font-weight: 300; letter-spacing: 0.38em;
    color: var(--gold); text-transform: uppercase; margin-bottom: 14px;
}
.hero-title {
    font-family: 'Noto Serif JP', serif;
    font-size: clamp(28px, 3.5vw, 46px); font-weight: 600;
    color: var(--white); line-height: 1.2; margin-bottom: 10px;
}
.hero-title span { color: var(--gold-lt); font-weight: 300; font-style: italic; }
.hero-sub {
    font-size: 13px; font-weight: 300; color: var(--muted);
    letter-spacing: 0.04em; line-height: 1.9;
}

/* ── Main content wrapper ── */
.main-wrap { padding: 36px 72px; max-width: 1000px; margin: 0 auto; }

/* ── Section label ── */
.sec-label {
    font-family: 'Cormorant Garamond', serif;
    font-size: 10.5px; letter-spacing: 0.42em; text-transform: uppercase;
    color: var(--gold); margin-bottom: 18px;
    display: flex; align-items: center; gap: 10px;
}
.sec-label::after { content: ''; flex: 1; height: 1px; background: var(--border); }

/* ── Format badges ── */
.badges { display: flex; gap: 8px; flex-wrap: wrap; margin: 10px 0 20px; }
.badge {
    background: rgba(201,168,76,0.07); border: 1px solid rgba(201,168,76,0.28);
    border-radius: 2px; padding: 4px 11px;
    font-size: 10.5px; letter-spacing: 0.13em;
    color: var(--gold-lt); font-family: 'Courier New', monospace; text-transform: uppercase;
}

/* ── File uploader ── */
[data-testid="stFileUploader"] {
    background: var(--card) !important;
    border: 1px solid var(--border) !important;
    border-radius: 4px !important; transition: border-color 0.3s;
}
[data-testid="stFileUploader"]:hover { border-color: rgba(201,168,76,0.5) !important; }
[data-testid="stFileUploaderDropzoneInstructions"] { color: var(--muted) !important; }
[data-testid="stFileUploaderFile"] {
    background: rgba(201,168,76,0.07) !important;
    border: 1px solid var(--border) !important;
    border-radius: 2px !important; color: var(--off-white) !important;
}

/* ── Generate button ── */
.stButton > button {
    background: transparent !important; color: var(--gold) !important;
    font-family: 'Noto Sans JP', sans-serif !important;
    font-size: 13.5px !important; font-weight: 500 !important;
    letter-spacing: 0.1em !important;
    border: 1px solid var(--gold) !important; border-radius: 2px !important;
    padding: 12px 32px !important; width: 100% !important;
    transition: all 0.25s ease !important;
}
.stButton > button:hover {
    background: rgba(201,168,76,0.1) !important;
    box-shadow: 0 0 22px rgba(201,168,76,0.22) !important;
}

/* ── Progress bar ── */
.stProgress > div { background: rgba(201,168,76,0.1) !important; border-radius: 2px !important; height: 3px !important; }
.stProgress > div > div { background: linear-gradient(90deg, var(--gold), var(--gold-lt)) !important; border-radius: 2px !important; }

/* ── SUCCESS BANNER ── */
.success-banner {
    background: linear-gradient(135deg, #081A08 0%, #0B2310 100%);
    border: 1.5px solid var(--success);
    border-radius: 6px; padding: 26px 36px;
    display: flex; align-items: center; justify-content: space-between;
    gap: 24px; margin-bottom: 24px;
    box-shadow: 0 0 40px rgba(34,197,94,0.13);
    animation: pop 0.45s cubic-bezier(0.34,1.56,0.64,1) both;
}
@keyframes pop {
    from { opacity:0; transform:scale(0.97) translateY(-6px); }
    to   { opacity:1; transform:scale(1) translateY(0); }
}
.success-left { display: flex; align-items: center; gap: 16px; }
.success-check {
    width: 46px; height: 46px; border-radius: 50%;
    background: rgba(34,197,94,0.14); border: 1.5px solid var(--success);
    display: flex; align-items: center; justify-content: center;
    font-size: 20px; flex-shrink: 0;
}
.success-title {
    font-family: 'Noto Serif JP', serif; font-size: 19px; font-weight: 600;
    color: var(--white); margin-bottom: 4px;
}
.success-meta { font-size: 11.5px; color: #86efac; letter-spacing: 0.04em; }

/* ── DOWNLOAD BUTTON (gold, glowing, most prominent) ── */
[data-testid="stDownloadButton"] > button {
    background: linear-gradient(135deg, #C9A84C 0%, #E8C97A 50%, #C9A84C 100%) !important;
    background-size: 200% !important;
    color: #05101E !important;
    font-family: 'Noto Sans JP', sans-serif !important;
    font-weight: 700 !important; font-size: 15px !important;
    letter-spacing: 0.08em !important;
    border: none !important; border-radius: 3px !important;
    padding: 16px 52px !important; width: 100% !important;
    box-shadow: 0 0 32px rgba(201,168,76,0.60), 0 4px 18px rgba(0,0,0,0.45) !important;
    transition: all 0.3s ease !important;
    animation: pulse 2.6s ease-in-out infinite !important;
}
[data-testid="stDownloadButton"] > button:hover {
    box-shadow: 0 0 52px rgba(201,168,76,0.85), 0 8px 28px rgba(0,0,0,0.5) !important;
    transform: translateY(-2px) !important;
    animation-play-state: paused !important;
}
@keyframes pulse {
    0%,100% { box-shadow: 0 0 32px rgba(201,168,76,0.60), 0 4px 18px rgba(0,0,0,0.45); }
    50%      { box-shadow: 0 0 52px rgba(201,168,76,0.85), 0 4px 18px rgba(0,0,0,0.45); }
}

/* ── Stat cards ── */
.stat-row { display: grid; grid-template-columns: repeat(3,1fr); gap: 14px; margin: 18px 0 26px; }
.stat-card {
    background: var(--card); border: 1px solid var(--border);
    border-radius: 4px; padding: 22px 26px; position: relative; overflow: hidden;
}
.stat-card::before {
    content: ''; position: absolute; top: 0; left: 0;
    width: 3px; height: 100%; background: linear-gradient(180deg, var(--gold), transparent);
}
.stat-n { font-family: 'Cormorant Garamond', serif; font-size: 40px; font-weight: 600; color: var(--gold-lt); line-height: 1; margin-bottom: 5px; }
.stat-l { font-size: 10.5px; letter-spacing: 0.16em; color: var(--muted); text-transform: uppercase; }

/* ── Category grid ── */
.cat-grid { display: grid; grid-template-columns: repeat(2,1fr); gap: 10px; margin: 14px 0 22px; }
.cat-card {
    background: var(--card); border: 1px solid var(--border);
    border-radius: 4px; padding: 16px 18px; display: flex; align-items: flex-start; gap: 12px;
}
.cat-icon {
    width: 33px; height: 33px; border-radius: 50%;
    background: rgba(201,168,76,0.1); border: 1px solid var(--border);
    display: flex; align-items: center; justify-content: center;
    font-size: 14px; flex-shrink: 0; margin-top: 2px;
}
.cat-name { font-family: 'Noto Serif JP', serif; font-size: 13px; font-weight: 600; color: var(--off-white); margin-bottom: 3px; }
.cat-cnt  { font-size: 11px; color: var(--gold); font-weight: 500; }
.cat-prev { font-size: 10.5px; color: var(--muted); margin-top: 5px; line-height: 1.6; display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; overflow: hidden; }

/* ── Log console ── */
.log-con {
    background: #040912; border: 1px solid rgba(201,168,76,0.10); border-radius: 4px;
    padding: 16px 20px; font-family: 'Courier New', monospace; font-size: 11.5px;
    color: #4ade80; line-height: 2.1; max-height: 190px; overflow-y: auto; margin: 12px 0;
}
.ll      { animation: fadeIn .25s ease; }
.ll.warn { color: #fbbf24; }
.ll.err  { color: #f87171; }
@keyframes fadeIn { from { opacity:0; transform:translateX(-4px); } to { opacity:1; } }

/* ── HR ── */
hr { border: none !important; border-top: 1px solid var(--border) !important; margin: 26px 0 !important; }

/* ── Sidebar: history ── */
.sb-head {
    background: linear-gradient(180deg, #030812, #0A1228);
    border-bottom: 1px solid var(--border); padding: 26px 18px 18px;
}
.sb-title { font-family: 'Noto Serif JP', serif; font-size: 14px; font-weight: 600; color: var(--white); margin-bottom: 3px; }
.sb-sub   { font-size: 10.5px; color: var(--muted); letter-spacing: 0.05em; }
.sb-item  { padding: 12px 18px; border-bottom: 1px solid rgba(201,168,76,0.07); }
.sb-name  { font-size: 11.5px; color: var(--off-white); margin-bottom: 3px; word-break: break-all; }
.sb-meta  { font-size: 10px; color: var(--muted); }
.sb-empty { padding: 26px 18px; font-size: 11.5px; color: var(--muted); text-align: center; line-height: 2.2; }

/* Sidebar download buttons: compact style */
section[data-testid="stSidebar"] [data-testid="stDownloadButton"] > button {
    background: rgba(201,168,76,0.08) !important;
    color: var(--gold-lt) !important;
    font-size: 11px !important; font-weight: 500 !important;
    padding: 6px 14px !important; border: 1px solid var(--border) !important;
    border-radius: 2px !important; letter-spacing: 0.06em !important;
    box-shadow: none !important; animation: none !important;
    margin-bottom: 8px !important;
}
section[data-testid="stSidebar"] [data-testid="stDownloadButton"] > button:hover {
    background: rgba(201,168,76,0.16) !important; transform: none !important;
}

/* ── Footer ── */
.footer {
    border-top: 1px solid var(--border); padding: 18px 72px;
    text-align: center; font-size: 10.5px; color: var(--muted); letter-spacing: 0.1em;
}
</style>
"""


# =============================================================================
# ファイル読み込み関数（BytesIO 対応）
# =============================================================================

def read_pptx_bytes(data: bytes, name: str) -> str:
    """PPTX ファイル（バイト列）から全スライドのテキストを抽出します。"""
    lines = [f"【出典：{name}】"]
    try:
        prs = Presentation(io.BytesIO(data))
        for i, slide in enumerate(prs.slides, 1):
            texts = [
                p.text.strip()
                for s in slide.shapes if s.has_text_frame
                for p in s.text_frame.paragraphs if p.text.strip()
            ]
            if texts:
                lines += [f"--- スライド {i} ---"] + texts
    except Exception as e:
        lines.append(f"（読み込みエラー: {e}）")
    return "\n".join(lines) + "\n"


def read_xlsx_bytes(data: bytes, name: str) -> str:
    """XLSX ファイル（バイト列）から全シートのセルデータを抽出します。"""
    lines = [f"【出典：{name}】"]
    try:
        wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
        for sn in wb.sheetnames:
            ws = wb[sn]
            lines.append(f"--- シート: {sn} ---")
            for row in ws.iter_rows():
                row_data = [str(c.value).strip() for c in row if c.value is not None]
                if row_data:
                    lines.append(" | ".join(row_data))
    except Exception as e:
        lines.append(f"（読み込みエラー: {e}）")
    return "\n".join(lines) + "\n"


def read_pdf_bytes(data: bytes, name: str) -> str:
    """PDF ファイル（バイト列）から全ページのテキストを抽出します。"""
    lines = [f"【出典：{name}】"]
    try:
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            for i, page in enumerate(pdf.pages, 1):
                t = page.extract_text()
                if t and t.strip():
                    lines += [f"--- ページ {i} ---", t.strip()]
    except Exception as e:
        lines.append(f"（読み込みエラー: {e}）")
    return "\n".join(lines) + "\n"


def read_txt_bytes(data: bytes, name: str) -> str:
    """TXT ファイル（バイト列）を文字コードに配慮して読み込みます。"""
    lines = [f"【出典：{name}】"]
    for enc in ["utf-8", "shift-jis", "cp932", "utf-16", "latin-1"]:
        try:
            lines.append(data.decode(enc).strip())
            return "\n".join(lines) + "\n"
        except (UnicodeDecodeError, LookupError):
            continue
    lines.append("（文字コードを特定できませんでした）")
    return "\n".join(lines) + "\n"


# =============================================================================
# 処理ロジック
# =============================================================================

def process_files(uploaded_files) -> tuple[list[dict], list[str]]:
    """アップロードされたファイルを読み込みます。"""
    READERS = {
        ".pptx": read_pptx_bytes,
        ".xlsx": read_xlsx_bytes,
        ".pdf":  read_pdf_bytes,
        ".txt":  read_txt_bytes,
    }
    results, logs = [], []
    for uf in uploaded_files:
        ext = Path(uf.name).suffix.lower()
        if ext not in READERS:
            logs.append(f"⏭  スキップ: {uf.name}（非対応フォーマット）")
            continue
        logs.append(f"📄  {uf.name} を読み込み中...")
        try:
            text = READERS[ext](uf.read(), uf.name)
            results.append({"filename": uf.name, "text": text})
            logs.append(f"✅  {uf.name} 完了")
        except Exception as e:
            logs.append(f"❌  {uf.name} 失敗: {e}")
    return results, logs


def classify(file_data_list: list[dict]) -> dict:
    """キーワードマッチングでテキストを4カテゴリに分類します（要約なし・原文整理）。"""
    cats   = {k: [] for k in CATEGORY_KEYWORDS}
    other  = []
    for fd in file_data_list:
        fn = fd["filename"]
        for line in fd["text"].split("\n"):
            s = line.strip()
            if not s or s.startswith("---") or s.startswith("【出典"):
                continue
            matched = False
            for cat, kws in CATEGORY_KEYWORDS.items():
                if any(k in s for k in kws):
                    cats[cat].append({"text": s, "source": fn})
                    matched = True
                    break
            if not matched and len(s) > 5:
                other.append({"text": s, "source": fn})
    if other:
        cats["その他・参考情報"] = other
    return cats


# =============================================================================
# PPTX 生成
# =============================================================================

def _bg(slide, color):
    f = slide.background.fill; f.solid(); f.fore_color.rgb = color

def _tb(slide, text, l, t, w, h, size, bold=False, color=None, align=None):
    """テキストボックスを追加するヘルパー関数。"""
    if align is None:
        align = PP_ALIGN.LEFT
    tb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = True
    p  = tf.paragraphs[0]; p.alignment = align
    run = p.add_run(); run.text = text
    run.font.size = Pt(size); run.font.bold = bold
    if color:
        run.font.color.rgb = color

def _title_slide(prs, today_str: str):
    """タイトルスライドを生成します。"""
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(sl, C_DARK)
    bar = sl.shapes.add_shape(1, Inches(0), Inches(0), Inches(0.15), Inches(7.5))
    bar.fill.solid(); bar.fill.fore_color.rgb = C_ACCENT; bar.line.fill.background()
    _tb(sl, "【自動生成】\nチーム進捗報告\n統合レポート", 0.5, 1.5, 9, 3, 40, bold=True, color=C_WHITE)
    _tb(sl,
        "向平様\n\nお忙しい中、ご確認いただきありがとうございます。\n"
        "本レポートは各部門からの報告資料を自動統合・整理したものです。\n"
        "情報の正確性を最優先し、原文を整理して掲載しております。",
        0.5, 4.5, 8.5, 2.5, 14, color=C_ACCENT)
    _tb(sl, f"生成日時: {today_str}　Project Relay 自動生成",
        0.5, 7.0, 9, 0.4, 10, color=C_CITE, align=PP_ALIGN.RIGHT)

def _index_slide(prs, cats: dict):
    """目次スライドを生成します。"""
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(sl, C_DARK)
    _tb(sl, "目　次", 0.4, 0.3, 9, 0.9, 32, bold=True, color=C_WHITE)
    sep = sl.shapes.add_shape(1, Inches(0.4), Inches(1.2), Inches(9.2), Inches(0.04))
    sep.fill.solid(); sep.fill.fore_color.rgb = C_ACCENT; sep.line.fill.background()
    lines, n = [], 3
    for cat, items in cats.items():
        if items:
            lines.append(f"  {n}.  {cat}　　（{len(items)} 件）"); n += 1
    _tb(sl, "\n".join(lines), 0.8, 1.5, 8.5, 5.5, 20, color=C_ACCENT)

def _content_slide(prs, category: str, items: list[dict]):
    """カテゴリごとのコンテンツスライドを生成します。"""
    if not items:
        return
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(sl, C_LIGHT)
    hdr = sl.shapes.add_shape(1, Inches(0), Inches(0), Inches(10), Inches(1.3))
    hdr.fill.solid(); hdr.fill.fore_color.rgb = C_DARK; hdr.line.fill.background()
    _tb(sl, category, 0.3, 0.1, 9, 1.1, 32, bold=True, color=C_WHITE)
    display = items[:12]
    body    = "\n".join(f"・ {x['text']}" for x in display)
    if len(items) > 12:
        body += f"\n（他 {len(items)-12} 件 — 元資料をご参照ください）"
    _tb(sl, body, 0.4, 1.5, 9.2, 5.5, 15, color=C_BODY)
    srcs = "出典：" + "、".join(sorted({x["source"] for x in display}))
    _tb(sl, srcs, 3.5, 6.95, 6.3, 0.4, 8, color=C_CITE, align=PP_ALIGN.RIGHT)

def make_pptx(cats: dict, today_str: str) -> bytes:
    """PPTX を生成してバイト列で返します。"""
    prs = Presentation()
    prs.slide_width  = Inches(10)
    prs.slide_height = Inches(7.5)
    _title_slide(prs, today_str)
    _index_slide(prs, cats)
    for cat, items in cats.items():
        if items:
            _content_slide(prs, cat, items)
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


# =============================================================================
# アーカイブ（history）操作
# =============================================================================

def save_to_history(pptx_bytes: bytes, filename: str) -> None:
    """生成済み PPTX を ./history に保存します。"""
    (HISTORY_DIR / filename).write_bytes(pptx_bytes)


def list_history() -> list[Path]:
    """history フォルダの PPTX 一覧を新しい順で返します。"""
    return sorted(HISTORY_DIR.glob("*.pptx"), key=lambda p: p.stat().st_mtime, reverse=True)


# =============================================================================
# サイドバー：過去の統合レポート一覧
# =============================================================================

def render_sidebar():
    """サイドバーに過去レポートの一覧とダウンロードボタンを表示します。"""
    with st.sidebar:
        st.markdown("""
        <div class="sb-head">
            <div class="sb-title">⬡ 過去の統合レポート</div>
            <div class="sb-sub">クリックでダウンロード</div>
        </div>
        """, unsafe_allow_html=True)

        history = list_history()

        if not history:
            st.markdown("""
            <div class="sb-empty">
                まだ履歴がありません。<br>
                レポートを生成すると<br>ここに自動保存されます。
            </div>
            """, unsafe_allow_html=True)
            return

        for path in history:
            mtime   = datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y/%m/%d %H:%M")
            size_kb = max(1, path.stat().st_size // 1024)
            st.markdown(f"""
            <div class="sb-item">
                <div class="sb-name">📄 {path.name}</div>
                <div class="sb-meta">{mtime} · {size_kb} KB</div>
            </div>
            """, unsafe_allow_html=True)
            st.download_button(
                label="↓ ダウンロード",
                data=path.read_bytes(),
                file_name=path.name,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                key=f"hist_{path.name}_{path.stat().st_mtime}",
                use_container_width=True,
            )


# =============================================================================
# UI パーツ
# =============================================================================

def render_hero():
    st.markdown("""
    <div class="hero">
        <div class="hero-eyebrow">⬡ &nbsp; Project Relay</div>
        <div class="hero-title">統合報告レポート<span>、自動生成。</span></div>
        <div class="hero-sub">
            バラバラな形式の報告資料を一括取込み、統一されたパワーポイントの叩き台を即座に生成します。<br>
            向平様の意思決定を最短距離でサポートするために設計されました。
        </div>
    </div>
    """, unsafe_allow_html=True)


def render_success_banner_and_download(filename: str, file_count: int, item_count: int, slide_count: int, pptx_bytes: bytes):
    """
    完了バナーとゴールドのダウンロードボタンを画面最上部に表示します。
    スクロール不要で即座に把握・操作できます。
    """
    st.markdown(f"""
    <div class="success-banner">
        <div class="success-left">
            <div class="success-check">✅</div>
            <div>
                <div class="success-title">レポート統合が完了しました</div>
                <div class="success-meta">
                    {file_count} ファイル読込 &nbsp;·&nbsp; {item_count} 項目抽出 &nbsp;·&nbsp; {slide_count} スライド生成
                </div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ゴールドのダウンロードボタンをバナーの直下に配置
    col_l, col_c, col_r = st.columns([1, 2, 1])
    with col_c:
        st.download_button(
            label="⬇　統合レポートをダウンロード",
            data=pptx_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            key="main_dl",
            use_container_width=True,
        )


def render_log_console(logs: list[str]):
    html = ""
    for log in logs:
        cls = "warn" if ("⏭" in log or "⚠" in log) else ("err" if "❌" in log else "")
        html += f'<div class="ll {cls}">{log}</div>'
    st.markdown(f'<div class="log-con">{html}</div>', unsafe_allow_html=True)


def render_stat_cards(nf: int, ni: int, ns: int):
    st.markdown(f"""
    <div class="stat-row">
        <div class="stat-card"><div class="stat-n">{nf}</div><div class="stat-l">読込ファイル数</div></div>
        <div class="stat-card"><div class="stat-n">{ni}</div><div class="stat-l">抽出情報項目数</div></div>
        <div class="stat-card"><div class="stat-n">{ns}</div><div class="stat-l">生成スライド枚数</div></div>
    </div>
    """, unsafe_allow_html=True)


def render_category_cards(cats: dict):
    html = ""
    for cat, items in cats.items():
        if not items:
            continue
        icon = CATEGORY_ICONS.get(cat, "📄")
        prev = (items[0]["text"][:56] + "…") if items else ""
        html += f"""
        <div class="cat-card">
            <div class="cat-icon">{icon}</div>
            <div>
                <div class="cat-name">{cat}</div>
                <div class="cat-cnt">{len(items)} 件の情報を抽出</div>
                <div class="cat-prev">{prev}</div>
            </div>
        </div>"""
    st.markdown(f'<div class="cat-grid">{html}</div>', unsafe_allow_html=True)


def render_footer():
    st.markdown("""
    <div class="footer">
        Project Relay v2 &nbsp;|&nbsp; 向平様専用 業務効率化プロトタイプ &nbsp;|&nbsp;
        python-pptx &nbsp;·&nbsp; pdfplumber &nbsp;·&nbsp; Streamlit
    </div>
    """, unsafe_allow_html=True)


# =============================================================================
# メイン
# =============================================================================

def main():
    # CSS注入
    st.markdown(CSS, unsafe_allow_html=True)

    # サイドバー（過去レポート一覧）
    render_sidebar()

    # ヒーローヘッダー
    render_hero()

    st.markdown('<div class="main-wrap">', unsafe_allow_html=True)

    # ===========================================================
    # 【優先表示】完了後は最上部にバナー＋ダウンロードボタンを配置
    # スクロール不要で即座にダウンロードできます
    # ===========================================================
    if st.session_state.get("pptx_ready"):
        cats   = st.session_state["cats"]
        ni     = sum(len(v) for v in cats.values())
        ns     = 2 + sum(1 for v in cats.values() if v)
        render_success_banner_and_download(
            filename    = st.session_state["filename"],
            file_count  = st.session_state["file_count"],
            item_count  = ni,
            slide_count = ns,
            pptx_bytes  = st.session_state["pptx_bytes"],
        )
        st.markdown("<hr>", unsafe_allow_html=True)

    # ── ① アップロードゾーン ──────────────────────────────────
    st.markdown('<div class="sec-label">01 &nbsp; ファイルをアップロード</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="badges">'
        '<span class="badge">.pptx</span>'
        '<span class="badge">.xlsx</span>'
        '<span class="badge">.pdf</span>'
        '<span class="badge">.txt</span>'
        '</div>',
        unsafe_allow_html=True,
    )

    uploaded = st.file_uploader(
        "ここに報告資料をドロップしてください　（複数ファイル対応）",
        type=["pptx", "xlsx", "pdf", "txt"],
        accept_multiple_files=True,
    )

    if not uploaded:
        st.markdown("""
        <div style="text-align:center;padding:32px;color:#6B7A9F;font-size:13px;letter-spacing:.04em;">
            PPTX・XLSX・PDF・TXT を複数ファイル一括でドロップできます。
        </div>
        """, unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
        render_footer()
        return

    # ── ② 生成ボタン ────────────────────────────────────────────
    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown('<div class="sec-label">02 &nbsp; レポートを生成</div>', unsafe_allow_html=True)

    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        go = st.button("⬡　統合レポートを生成する", use_container_width=True)

    # ── ③ 処理実行 ──────────────────────────────────────────────
    if go:
        if not PPTX_AVAILABLE:
            st.error("❌ python-pptx が必要です。pip install python-pptx を実行してください。")
            st.markdown('</div>', unsafe_allow_html=True)
            return

        st.markdown("<hr>", unsafe_allow_html=True)
        st.markdown('<div class="sec-label">03 &nbsp; 処理ログ</div>', unsafe_allow_html=True)

        prog   = st.progress(0)
        status = st.empty()
        logs   = []

        # Step 1 : ファイル読み込み
        status.markdown("**🔍 ファイルを読み込んでいます...**")
        file_data, file_logs = process_files(uploaded)
        logs.extend(file_logs)
        prog.progress(35); time.sleep(0.15)

        # Step 2 : 分類
        status.markdown("**🗂️ データを分類・整理しています...**")
        cats = classify(file_data)
        logs.append("🗂  キーワードマッチングで分類完了")
        prog.progress(62); time.sleep(0.15)

        # Step 3 : PPTX 生成
        status.markdown("**🖥️ パワーポイントを構成しています...**")
        today_str  = datetime.now().strftime("%Y年%m月%d日 %H:%M")
        pptx_bytes = make_pptx(cats, today_str)
        logs.append("🎨  タイトル・目次スライドを生成")
        logs.append(f"📊  カテゴリスライドを生成（{sum(1 for v in cats.values() if v)} 枚）")
        prog.progress(88); time.sleep(0.15)

        # Step 4 : history へ保存
        status.markdown("**💾 アーカイブに保存しています...**")
        ts       = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"統合レポート_{ts}.pptx"
        save_to_history(pptx_bytes, filename)
        logs.append(f"💾  ./history/{filename} に保存完了")
        prog.progress(100)
        status.empty()

        logs.append("✅  全処理が完了しました")
        render_log_console(logs)

        # セッションに保存 → st.rerun() でバナーを最上部に表示
        st.session_state["pptx_ready"] = True
        st.session_state["pptx_bytes"] = pptx_bytes
        st.session_state["cats"]       = cats
        st.session_state["file_count"] = len(file_data)
        st.session_state["filename"]   = filename
        st.rerun()  # ← バナーを最上部に表示するため再描画

    # ── ④ 完了後のサマリー（バナー下に詳細情報を表示） ──────────
    if st.session_state.get("pptx_ready"):
        cats = st.session_state["cats"]
        ni   = sum(len(v) for v in cats.values())
        ns   = 2 + sum(1 for v in cats.values() if v)

        st.markdown('<div class="sec-label">04 &nbsp; 生成サマリー</div>', unsafe_allow_html=True)
        render_stat_cards(st.session_state["file_count"], ni, ns)

        st.markdown('<div class="sec-label">05 &nbsp; カテゴリ別プレビュー</div>', unsafe_allow_html=True)
        render_category_cards(cats)

    st.markdown('</div>', unsafe_allow_html=True)
    render_footer()


if __name__ == "__main__":
    main()
    