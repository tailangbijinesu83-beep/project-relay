# =============================================================================
# Project Relay - Web版 統合報告パワーポイント自動生成アプリ
# 向平氏専用 業務効率化ツール（Streamlit）
# =============================================================================
#
# 【インストール】
#   pip install streamlit python-pptx openpyxl pdfplumber
#
# 【起動】
#   streamlit run app.py
#
# 【認証パスワード】
#   relay2026
# =============================================================================

import io
import time
import tempfile
from datetime import datetime
from pathlib import Path

import streamlit as st


# =============================================================================
# 1. 認証ゲートロジック
# =============================================================================
if "auth" not in st.session_state:
    st.session_state.auth = False

def check_password():
    if st.session_state.get("pw_entry") == "relay2026":
        st.session_state.auth = True
    else:
        st.error("パスワードが正しくありません")

# 認証されていない場合は、ここで処理を止める
if not st.session_state.auth:
    st.set_page_config(page_title="Project Relay | Security", page_icon="⬡")
    st.markdown("<style>body{background-color:#0D1B3E;color:white;}</style>", unsafe_allow_html=True)
    st.title("⬡ Project Relay - Security Gate")
    st.text_input("向平様専用パスワードを入力してください", type="password", key="pw_entry", on_change=check_password)
    st.info("認証が完了するまで、すべての機能はロックされています。")
    st.stop()


# =============================================================================
# 2. サードパーティライブラリ（未インストール時もクラッシュしない設計）
# =============================================================================
try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
    from pptx.enum.shapes import MSO_SHAPE
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
# 3. Streamlit ページ設定
# =============================================================================
st.set_page_config(
    page_title="Project Relay | 統合報告レポート生成",
    page_icon="⬡",
    layout="wide",
    initial_sidebar_state="collapsed",
)


# =============================================================================
# 4. カラーパレット & グローバルCSS（高級ビジネスデザイン）
# =============================================================================
GLOBAL_CSS = """
<style>
/* ── Google Fonts ── */
@import url('https://fonts.googleapis.com/css2?family=Noto+Serif+JP:wght@300;400;600&family=Noto+Sans+JP:wght@300;400;500&family=Cormorant+Garamond:ital,wght@0,300;0,600;1,300&display=swap');

/* ── CSS Variables ── */
:root {
    --navy:        #0D1B3E;
    --navy-mid:    #1E2D5A;
    --navy-light:  #2A3F7E;
    --gold:        #D60036;
    --gold-light:  #FF3B6B;
    --ice:         #CADCFC;
    --off-white:   #F0F2F8;
    --muted:       #7A8AB0;
    --surface:     #111827;
    --card:        #1A2540;
    --border:      rgba(214,0,54,0.25);
}

/* ── Base Reset ── */
html, body, [data-testid="stAppViewContainer"] {
    background-color: var(--navy) !important;
    color: var(--off-white) !important;
    font-family: 'Noto Sans JP', sans-serif;
}

[data-testid="stHeader"]  { background: transparent !important; }
[data-testid="stSidebar"] { background: var(--surface) !important; }
[data-testid="stVerticalBlock"] { gap: 0 !important; }
.block-container {
    padding: 0 !important;
    max-width: 100% !important;
}

/* ── Hero Header ── */
.hero {
    background: linear-gradient(135deg, #060D20 0%, #0D1B3E 50%, #1A2D5A 100%);
    border-bottom: 1px solid var(--border);
    padding: 56px 80px 48px;
    position: relative;
    overflow: hidden;
}
.hero::before {
    content: '';
    position: absolute;
    top: -60px; right: -60px;
    width: 320px; height: 320px;
    border-radius: 50%;
    background: radial-gradient(circle, rgba(214,0,54,0.08) 0%, transparent 70%);
    pointer-events: none;
}
.hero::after {
    content: '';
    position: absolute;
    bottom: 0; left: 0; right: 0;
    height: 1px;
    background: linear-gradient(90deg, transparent, var(--gold), transparent);
}
.hero-eyebrow {
    font-family: 'Cormorant Garamond', serif;
    font-size: 13px;
    font-weight: 300;
    letter-spacing: 0.35em;
    color: var(--gold);
    text-transform: uppercase;
    margin-bottom: 16px;
}
.hero-title {
    font-family: 'Noto Serif JP', serif;
    font-size: clamp(32px, 4vw, 52px);
    font-weight: 600;
    color: #FFFFFF;
    line-height: 1.2;
    letter-spacing: -0.01em;
    margin-bottom: 12px;
}
.hero-title span {
    color: var(--gold-light);
    font-weight: 300;
    font-style: italic;
}
.hero-subtitle {
    font-family: 'Noto Sans JP', sans-serif;
    font-size: 14px;
    font-weight: 300;
    color: var(--muted);
    letter-spacing: 0.05em;
    line-height: 1.8;
}

/* ── Main Content Area ── */
.main-content {
    padding: 48px 80px;
    max-width: 1100px;
    margin: 0 auto;
}

/* ── Section Label ── */
.section-label {
    font-family: 'Cormorant Garamond', serif;
    font-size: 11px;
    letter-spacing: 0.4em;
    text-transform: uppercase;
    color: var(--gold);
    margin-bottom: 20px;
    display: flex;
    align-items: center;
    gap: 12px;
}
.section-label::after {
    content: '';
    flex: 1;
    height: 1px;
    background: var(--border);
}

/* ── File Upload Zone ── */
[data-testid="stFileUploader"] {
    background: var(--card) !important;
    border: 1px solid var(--border) !important;
    border-radius: 4px !important;
    padding: 12px !important;
    transition: border-color 0.3s ease;
}
[data-testid="stFileUploader"]:hover {
    border-color: rgba(214,0,54,0.55) !important;
}
[data-testid="stFileUploaderDropzoneInstructions"] {
    color: var(--muted) !important;
}

/* ── File Chip (uploaded file tags) ── */
[data-testid="stFileUploaderFile"] {
    background: rgba(255,255,255,0.15) !important;
    border: 1.5px solid var(--gold) !important;
    border-radius: 4px !important;
    color: #FFFFFF !important;
    box-shadow:
        0 0 0 1px rgba(214,0,54,0.35),
        0 2px 12px rgba(214,0,54,0.18),
        inset 0 1px 0 rgba(255,255,255,0.12) !important;
    margin-bottom: 8px !important;
    padding: 8px 12px !important;
}

/* ── File name — extreme weight, blazing white glow ── */
[data-testid="stFileUploaderFileName"] {
    color: #FFFFFF !important;
    font-weight: 800 !important;
    font-size: 13.5px !important;
    letter-spacing: 0.025em !important;
    text-shadow:
        0 0 10px rgba(255,255,255,0.80),
        0 0 20px rgba(214,0,54,0.40),
        0 0 40px rgba(214,0,54,0.20) !important;
}

/* ── File size / metadata — ice blue ── */
[data-testid="stFileUploaderFile"] small,
[data-testid="stFileUploaderFile"] [class*="fileSize"],
[data-testid="stFileUploaderFile"] span:not([data-testid="stFileUploaderFileName"]) {
    color: var(--ice) !important;
    font-size: 11px !important;
    opacity: 0.90 !important;
}

/* ── File type icon — red glow ── */
[data-testid="stFileUploaderFile"] svg {
    fill: var(--gold-light) !important;
    opacity: 1 !important;
    filter: drop-shadow(0 0 6px rgba(214,0,54,0.70)) !important;
}

/* ── Delete (×) button — gold-light, scaled up ── */
[data-testid="stFileUploaderFile"] button,
[data-testid="stFileUploaderDeleteBtn"] {
    color: var(--gold-light) !important;
    opacity: 1 !important;
    transform: scale(1.2) !important;
    transition: transform 0.15s, filter 0.15s !important;
}
[data-testid="stFileUploaderFile"] button:hover,
[data-testid="stFileUploaderDeleteBtn"]:hover {
    filter: drop-shadow(0 0 6px rgba(255,59,107,0.90)) !important;
    transform: scale(1.35) !important;
}

/* ── Stat Cards ── */
.stat-row {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 16px;
    margin: 28px 0;
}
.stat-card {
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 4px;
    padding: 24px 28px;
    position: relative;
    overflow: hidden;
}
.stat-card::before {
    content: '';
    position: absolute;
    top: 0; left: 0;
    width: 3px; height: 100%;
    background: linear-gradient(180deg, var(--gold), transparent);
}
.stat-number {
    font-family: 'Cormorant Garamond', serif;
    font-size: 42px;
    font-weight: 600;
    color: var(--gold-light);
    line-height: 1;
    margin-bottom: 6px;
}
.stat-label {
    font-size: 11px;
    letter-spacing: 0.15em;
    color: var(--muted);
    text-transform: uppercase;
}

/* ── Category Preview Cards ── */
.category-grid {
    display: grid;
    grid-template-columns: repeat(2, 1fr);
    gap: 12px;
    margin: 20px 0;
}
.category-card {
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 4px;
    padding: 18px 22px;
    display: flex;
    align-items: flex-start;
    gap: 14px;
}
.category-icon {
    width: 36px; height: 36px;
    border-radius: 50%;
    background: rgba(214,0,54,0.12);
    border: 1px solid var(--border);
    display: flex; align-items: center; justify-content: center;
    font-size: 16px;
    flex-shrink: 0;
    margin-top: 2px;
}
.category-name {
    font-family: 'Noto Serif JP', serif;
    font-size: 14px;
    font-weight: 600;
    color: var(--off-white);
    margin-bottom: 4px;
}
.category-count {
    font-size: 12px;
    color: var(--gold);
    font-weight: 500;
}
.category-preview {
    font-size: 11px;
    color: var(--muted);
    margin-top: 6px;
    line-height: 1.6;
    display: -webkit-box;
    -webkit-line-clamp: 2;
    -webkit-box-orient: vertical;
    overflow: hidden;
}

/* ── Log Console ── */
.log-console {
    background: #080E1C;
    border: 1px solid rgba(214,0,54,0.15);
    border-radius: 4px;
    padding: 20px 24px;
    font-family: 'Courier New', monospace;
    font-size: 12px;
    color: #4ADE80;
    line-height: 2;
    max-height: 220px;
    overflow-y: auto;
    margin: 16px 0;
}
.log-line { animation: fadeIn 0.3s ease; }
.log-line.warn  { color: #FBBF24; }
.log-line.error { color: #F87171; }
@keyframes fadeIn { from { opacity: 0; transform: translateX(-4px); } to { opacity: 1; } }

/* ── Progress Bar ── */
.stProgress > div > div {
    background: linear-gradient(90deg, var(--gold), var(--gold-light)) !important;
    border-radius: 2px !important;
}
.stProgress > div {
    background: rgba(214,0,54,0.1) !important;
    border-radius: 2px !important;
    height: 3px !important;
}

/* ── Download Section ── */
.download-section {
    background: linear-gradient(135deg, #08112A 0%, #0F1E3A 60%, #152040 100%);
    border: 2.5px solid var(--gold) !important;
    border-radius: 8px;
    padding: 48px 56px;
    text-align: center;
    margin: 32px 0;
    position: relative;
    overflow: hidden;
    box-shadow:
        0 0 0 1px rgba(214,0,54,0.15),
        0 0 60px rgba(214,0,54,0.18),
        inset 0 1px 0 rgba(255,255,255,0.05);
}
.download-section::before {
    content: '';
    position: absolute;
    top: -50%; left: -50%;
    width: 200%; height: 200%;
    background: radial-gradient(circle at center, rgba(214,0,54,0.07) 0%, transparent 55%);
    pointer-events: none;
}
.download-title {
    font-family: 'Noto Serif JP', serif;
    font-size: 28px;
    font-weight: 600;
    color: #FFFFFF;
    margin-bottom: 10px;
    text-shadow: 0 0 30px rgba(214,0,54,0.35);
    letter-spacing: 0.02em;
}
.download-subtitle {
    font-size: 14px;
    color: var(--ice);
    margin-bottom: 32px;
    line-height: 1.8;
    opacity: 0.85;
}

/* ── Pulse wrapper — forces the Streamlit DL button to throb ── */
.pulse-button {
    animation: pulse-gold 1.8s ease-in-out infinite;
    border-radius: 4px;
    display: block;
}
@keyframes pulse-gold {
    0%   { box-shadow: 0 0  8px rgba(214,0,54,0.50), 0 0  16px rgba(214,0,54,0.30); }
    50%  { box-shadow: 0 0 28px rgba(214,0,54,0.90), 0 0  52px rgba(214,0,54,0.55); }
    100% { box-shadow: 0 0  8px rgba(214,0,54,0.50), 0 0  16px rgba(214,0,54,0.30); }
}

[data-testid="stDownloadButton"] > button {
    background: linear-gradient(135deg, var(--gold), var(--gold-light)) !important;
    color: #FFFFFF !important;
    font-family: 'Noto Sans JP', sans-serif !important;
    font-weight: 700 !important;
    font-size: 16px !important;
    letter-spacing: 0.10em !important;
    border: none !important;
    border-radius: 4px !important;
    padding: 18px 52px !important;
    cursor: pointer !important;
    transition: transform 0.20s ease, box-shadow 0.20s ease !important;
    width: 100% !important;
}
[data-testid="stDownloadButton"] > button:hover {
    transform: translateY(-3px) scale(1.02) !important;
    box-shadow: 0 12px 40px rgba(214,0,54,0.65) !important;
}

/* ── Generate Button ── */
.stButton > button {
    background: transparent !important;
    color: var(--gold) !important;
    font-family: 'Noto Sans JP', sans-serif !important;
    font-size: 14px !important;
    font-weight: 500 !important;
    letter-spacing: 0.1em !important;
    border: 1px solid var(--gold) !important;
    border-radius: 2px !important;
    padding: 12px 36px !important;
    transition: all 0.25s ease !important;
    width: 100% !important;
}
.stButton > button:hover {
    background: rgba(214,0,54,0.1) !important;
    box-shadow: 0 0 20px rgba(214,0,54,0.2) !important;
}

/* ── Alerts ── */
.stAlert {
    background: var(--card) !important;
    border-radius: 4px !important;
    border-left-color: var(--gold) !important;
}

/* ── Divider ── */
hr {
    border: none !important;
    border-top: 1px solid var(--border) !important;
    margin: 32px 0 !important;
}

/* ── Supported formats badge ── */
.format-badges {
    display: flex;
    gap: 8px;
    flex-wrap: wrap;
    margin: 12px 0 24px;
}
.badge {
    background: rgba(214,0,54,0.08);
    border: 1px solid rgba(214,0,54,0.3);
    border-radius: 2px;
    padding: 4px 12px;
    font-size: 11px;
    letter-spacing: 0.12em;
    color: var(--gold-light);
    font-family: 'Courier New', monospace;
    text-transform: uppercase;
}

/* ── Footer ── */
.footer {
    border-top: 1px solid var(--border);
    padding: 24px 80px;
    text-align: center;
    font-size: 11px;
    color: var(--muted);
    letter-spacing: 0.1em;
}

/* Streamlit top padding fix */
.appview-container .main .block-container { padding-top: 0 !important; }
section[data-testid="stSidebar"] { display: none; }
</style>
"""


# =============================================================================
# 5. カラーパレット（PPTX生成用：IIJブランドに準拠）
# =============================================================================
COLOR_IIJ_CHARCOAL = RGBColor(0x33, 0x33, 0x33) if PPTX_AVAILABLE else None
COLOR_IIJ_RED      = RGBColor(0xD6, 0x00, 0x36) if PPTX_AVAILABLE else None
COLOR_WHITE        = RGBColor(0xFF, 0xFF, 0xFF) if PPTX_AVAILABLE else None
COLOR_LIGHT_GRAY   = RGBColor(0xF4, 0xF4, 0xF4) if PPTX_AVAILABLE else None
COLOR_BODY_TEXT    = RGBColor(0x33, 0x33, 0x33) if PPTX_AVAILABLE else None
COLOR_CITATION     = RGBColor(0x99, 0x99, 0x99) if PPTX_AVAILABLE else None


# =============================================================================
# 6. 分類キーワード定義
# =============================================================================
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


# =============================================================================
# 7. ファイル読み込み関数（BytesIO対応版）
# =============================================================================

def read_pptx_bytes(file_bytes: bytes, filename: str) -> str:
    lines = [f"【出典：{filename}】"]
    try:
        prs = Presentation(io.BytesIO(file_bytes))
        for i, slide in enumerate(prs.slides, start=1):
            slide_texts = []
            for shape in slide.shapes:
                if shape.has_text_frame:
                    for para in shape.text_frame.paragraphs:
                        text = para.text.strip()
                        if text:
                            slide_texts.append(text)
            if slide_texts:
                lines.append(f"--- スライド {i} ---")
                lines.extend(slide_texts)
    except Exception as e:
        lines.append(f"（読み込みエラー: {e}）")
    return "\n".join(lines) + "\n"


def read_xlsx_bytes(file_bytes: bytes, filename: str) -> str:
    lines = [f"【出典：{filename}】"]
    try:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            lines.append(f"--- シート: {sheet_name} ---")
            for row in ws.iter_rows():
                row_data = [str(cell.value).strip() for cell in row if cell.value is not None]
                if row_data:
                    lines.append(" | ".join(row_data))
    except Exception as e:
        lines.append(f"（読み込みエラー: {e}）")
    return "\n".join(lines) + "\n"


def read_pdf_bytes(file_bytes: bytes, filename: str) -> str:
    lines = [f"【出典：{filename}】"]
    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for i, page in enumerate(pdf.pages, start=1):
                text = page.extract_text()
                if text and text.strip():
                    lines.append(f"--- ページ {i} ---")
                    lines.append(text.strip())
    except Exception as e:
        lines.append(f"（読み込みエラー: {e}）")
    return "\n".join(lines) + "\n"


def read_txt_bytes(file_bytes: bytes, filename: str) -> str:
    lines = [f"【出典：{filename}】"]
    encodings = ["utf-8", "shift-jis", "cp932", "utf-16", "latin-1"]
    for enc in encodings:
        try:
            content = file_bytes.decode(enc).strip()
            lines.append(content)
            return "\n".join(lines) + "\n"
        except (UnicodeDecodeError, LookupError):
            continue
    lines.append("（文字コードを特定できませんでした）")
    return "\n".join(lines) + "\n"


# =============================================================================
# 8. 処理ロジック
# =============================================================================

def process_uploaded_files(uploaded_files) -> tuple[list[dict], list[str]]:
    READERS = {
        ".pptx": read_pptx_bytes,
        ".xlsx": read_xlsx_bytes,
        ".pdf":  read_pdf_bytes,
        ".txt":  read_txt_bytes,
    }
    results = []
    logs = []

    for uf in uploaded_files:
        ext = Path(uf.name).suffix.lower()
        if ext not in READERS:
            logs.append(f"⏭  スキップ: {uf.name}（非対応フォーマット）")
            continue

        logs.append(f"📄  {uf.name} を読み込み中...")
        try:
            file_bytes = uf.read()
            text = READERS[ext](file_bytes, uf.name)
            results.append({"filename": uf.name, "text": text})
            logs.append(f"✅  {uf.name} の読み込み完了")
        except Exception as e:
            logs.append(f"❌  {uf.name} の読み込みに失敗: {e}")

    return results, logs


def classify_text_to_categories(file_data_list: list[dict]) -> dict:
    categories = {cat: [] for cat in CATEGORY_KEYWORDS}
    uncategorized = []

    for file_data in file_data_list:
        filename = file_data["filename"]
        for line in file_data["text"].split("\n"):
            line = line.replace("<", "&lt;").replace(">", "&gt;")  # HTMLコード露出防止
            stripped = line.strip()
            if not stripped or stripped.startswith("---") or stripped.startswith("【出典"):
                continue
            matched = False
            for category, keywords in CATEGORY_KEYWORDS.items():
                if any(kw in stripped for kw in keywords):
                    categories[category].append({"text": stripped, "source": filename})
                    matched = True
                    break
            if not matched and len(stripped) > 5:
                uncategorized.append({"text": stripped, "source": filename})

    if uncategorized:
        categories["その他・参考情報"] = uncategorized

    return categories


# =============================================================================
# 9. PowerPoint 生成関数（IIJデザイン・ブランディング）
# =============================================================================

def _add_bg(slide, color):
    """スライド背景色を設定します。"""
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def _draw_iij_logo(slide):
    """IIJロゴの完全再現：フォントとドット位置を精密調整"""
    # "IIJ" テキスト
    tb = slide.shapes.add_textbox(Inches(8.6), Inches(0.2), Inches(1.0), Inches(0.5))
    tf = tb.text_frame
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = "IIJ"
    run.font.size = Pt(28)
    run.font.bold = True
    run.font.color.rgb = COLOR_IIJ_CHARCOAL

    # 赤いドット：Jの右側に寄り添う配置へ修正
    dot_size = Inches(0.12)
    dot = slide.shapes.add_shape(
        MSO_SHAPE.OVAL,
        Inches(9.34), Inches(0.48),
        dot_size, dot_size,
    )
    dot.fill.solid()
    dot.fill.fore_color.rgb = COLOR_IIJ_RED
    dot.line.fill.background()


def _add_tb(slide, text, l, t, w, h, size, bold=False, color=None, align=None):
    """スライドにテキストボックスを追加するヘルパー関数。"""
    if align is None:
        align = PP_ALIGN.LEFT
    tb = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = color


def _build_title_slide(prs, today_str: str):
    """タイトルスライド（1枚目）を生成します。"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_bg(slide, COLOR_IIJ_CHARCOAL)

    # 左側の赤いアクセントバー
    bar = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(0), Inches(0), Inches(0.2), Inches(7.5),
    )
    bar.fill.solid()
    bar.fill.fore_color.rgb = COLOR_IIJ_RED
    bar.line.fill.background()

    _add_tb(
        slide,
        "Project Relay\n統合報告レポート",
        0.6, 2.0, 9.0, 2.5,
        44, bold=True, color=COLOR_WHITE, align=PP_ALIGN.LEFT,
    )

    greeting = (
        "向平 友治 様\n\n"
        "各部門資料から抽出された最新のステータスを統合しました。\n"
        "IIJブランドに準拠したフォーマットで整理しております。"
    )
    _add_tb(slide, greeting, 0.6, 4.5, 8.5, 2.0, 16, color=COLOR_WHITE, align=PP_ALIGN.LEFT)
    _add_tb(slide, f"生成日: {today_str}", 0.6, 6.8, 9.0, 0.4, 12, color=COLOR_WHITE, align=PP_ALIGN.LEFT)


def _build_index_slide(prs, categories: dict):
    """目次スライド（2枚目）を生成します。"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_bg(slide, COLOR_WHITE)
    _draw_iij_logo(slide)

    _add_tb(slide, "目次 / Index", 0.5, 0.5, 8.0, 1.0, 28, bold=True, color=COLOR_IIJ_CHARCOAL)

    # 区切り線
    sep = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(0.5), Inches(1.3), Inches(9.0), Inches(0.02),
    )
    sep.fill.solid()
    sep.fill.fore_color.rgb = COLOR_IIJ_RED
    sep.line.fill.background()

    index_lines = []
    num = 1
    for cat, items in categories.items():
        if items:
            index_lines.append(f"  0{num}.  {cat} ({len(items)} items)")
            num += 1

    _add_tb(slide, "\n".join(index_lines), 0.8, 1.8, 8.5, 5.0, 20, color=COLOR_IIJ_CHARCOAL)


def _build_content_slide(prs, category: str, items: list[dict]):
    """カテゴリごとのコンテンツスライドを生成します（最大10件/スライド）。"""
    if not items:
        return

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _add_bg(slide, COLOR_WHITE)
    _draw_iij_logo(slide)

    # ヘッダー
    _add_tb(slide, category, 0.5, 0.4, 8.0, 0.8, 28, bold=True, color=COLOR_IIJ_CHARCOAL)

    # 小さな赤いドット付きの下線
    sep = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(0.5), Inches(1.1), Inches(4.0), Inches(0.03),
    )
    sep.fill.solid()
    sep.fill.fore_color.rgb = COLOR_IIJ_RED
    sep.line.fill.background()

    MAX_ITEMS = 10
    display_items = items[:MAX_ITEMS]
    remaining = len(items) - MAX_ITEMS

    body_lines = [f"■ {item['text']}" for item in display_items]
    sources_seen = {item["source"] for item in display_items}
    if remaining > 0:
        body_lines.append(f"（他 {remaining} 件の情報を省略 — 元資料を確認してください）")

    _add_tb(slide, "\n".join(body_lines), 0.6, 1.6, 8.8, 5.0, 14, color=COLOR_BODY_TEXT)

    sources_str = "Source: " + ", ".join(sorted(sources_seen))
    _add_tb(slide, sources_str, 3.5, 7.0, 6.0, 0.3, 9, color=COLOR_CITATION, align=PP_ALIGN.RIGHT)


def generate_pptx_bytes(categories: dict, today_str: str) -> bytes:
    """カテゴリ辞書からPowerPointを生成し、バイト列で返します。"""
    prs = Presentation()
    prs.slide_width  = Inches(10)
    prs.slide_height = Inches(7.5)

    _build_title_slide(prs, today_str)
    _build_index_slide(prs, categories)

    for category, items in categories.items():
        if items:
            _build_content_slide(prs, category, items)

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


# =============================================================================
# 10. UI レンダリング関数
# =============================================================================

def render_hero():
    """ページ上部のヒーローヘッダーを描画します。"""
    st.markdown("""
    <div class="hero">
        <div class="hero-eyebrow">⬡ &nbsp; Project Relay</div>
        <div class="hero-title">統合報告レポート<span>、自動生成。</span></div>
        <div class="hero-subtitle">
            バラバラな形式の報告資料を一括取込み、IIJブランド仕様のパワーポイントを即座に生成します。<br>
            向平様の意思決定を最短距離でサポートするために設計されました。
        </div>
    </div>
    """, unsafe_allow_html=True)


def render_format_badges():
    """対応フォーマットのバッジを表示します。"""
    st.markdown("""
    <div class="format-badges">
        <span class="badge">.pptx</span>
        <span class="badge">.xlsx</span>
        <span class="badge">.pdf</span>
        <span class="badge">.txt</span>
    </div>
    """, unsafe_allow_html=True)


def render_stat_cards(num_files: int, total_items: int, num_slides: int):
    """3列の統計カードを表示します。"""
    st.markdown(f"""
    <div class="stat-row">
        <div class="stat-card">
            <div class="stat-number">{num_files}</div>
            <div class="stat-label">読み込んだファイル数</div>
        </div>
        <div class="stat-card">
            <div class="stat-number">{total_items}</div>
            <div class="stat-label">分類された情報項目数</div>
        </div>
        <div class="stat-card">
            <div class="stat-number">{num_slides}</div>
            <div class="stat-label">生成スライド枚数</div>
        </div>
    </div>
    """, unsafe_allow_html=True)


def render_category_preview(categories: dict):
    """カテゴリ別プレビューカードを2列グリッドで表示します。"""
    cards_html = ""
    for cat, items in categories.items():
        if not items:
            continue
        icon = CATEGORY_ICONS.get(cat, "📄")
        preview = items[0]["text"].replace("<", "&lt;").replace(">", "&gt;")[:60] + "…"
        cards_html += f"""
        <div class="category-card">
            <div class="category-icon">{icon}</div>
            <div>
                <div class="category-name">{cat}</div>
                <div class="category-count">{len(items)} 件の情報を抽出</div>
                <div class="category-preview">{preview}</div>
            </div>
        </div>
        """
    st.markdown(f'<div class="category-grid">{cards_html}</div>', unsafe_allow_html=True)


def render_log_console(logs: list[str]):
    """処理ログをターミナル風コンソールに表示します。"""
    lines_html = ""
    for log in logs:
        css_class = "warn" if "⏭" in log or "⚠" in log else ("error" if "❌" in log else "")
        lines_html += f'<div class="log-line {css_class}">{log}</div>'
    st.markdown(f'<div class="log-console">{lines_html}</div>', unsafe_allow_html=True)


def render_footer():
    """ページ下部のフッターを表示します。"""
    st.markdown(
        '<div class="footer">'
        'Project Relay &nbsp;|&nbsp; 向平 友治 様専用プロトタイプ &nbsp;|&nbsp; IIJ Brand Guidelines Applied'
        '</div>',
        unsafe_allow_html=True,
    )


# =============================================================================
# 11. メイン（認証済みユーザー専用）
# =============================================================================

def main():
    """
    メインUIを描画します。
    認証後のみ実行されます（st.stop()により未認証時はここに到達しません）。
    """
    st.markdown(GLOBAL_CSS, unsafe_allow_html=True)
    render_hero()
    st.markdown('<div class="main-content">', unsafe_allow_html=True)

    # ── ① アップロードゾーン ──────────────────────────────────────────────────
    st.markdown('<div class="section-label">01 &nbsp; ファイルをアップロード</div>', unsafe_allow_html=True)
    render_format_badges()

    uploaded_files = st.file_uploader(
        "ここに報告資料をドロップしてください　（複数ファイル対応）",
        type=["pptx", "xlsx", "pdf", "txt"],
        accept_multiple_files=True,
    )

    if not uploaded_files:
        st.markdown('</div>', unsafe_allow_html=True)
        render_footer()
        return

    # ── ② 生成ボタン ─────────────────────────────────────────────────────────
    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown('<div class="section-label">02 &nbsp; レポートを生成</div>', unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        generate_btn = st.button("⬡　統合レポートを生成する", use_container_width=True)

    # ── ③ 処理実行 ───────────────────────────────────────────────────────────
    if generate_btn or st.session_state.get("pptx_ready"):
        if generate_btn:
            # プログレスバーのみ表示（中間ログ・ステータス文字列は出さない）
            progress = st.progress(0)

            # Step 1: ファイル読み込み
            file_data_list, _ = process_uploaded_files(uploaded_files)
            progress.progress(35)

            # Step 2: キーワード分類
            categories = classify_text_to_categories(file_data_list)
            progress.progress(65)

            # Step 3: PPTX 生成
            today_str  = datetime.now().strftime("%Y年%m月%d日 %H:%M")
            pptx_bytes = generate_pptx_bytes(categories, today_str)
            progress.progress(100)

            # プログレスバーを消去してノイズをゼロに
            progress.empty()

            # セッションに保存
            st.session_state["pptx_ready"]   = True
            st.session_state["pptx_bytes"]   = pptx_bytes
            st.session_state["categories"]   = categories
            st.session_state["file_count"]   = len(file_data_list)
            st.session_state["show_toast"]   = True   # 初回のみトーストを出すフラグ
            st.rerun()

        # ── ④ 完了後の表示 ─────────────────────────────────────────────────────
        if st.session_state.get("pptx_ready"):

            # ── トースト通知（生成直後の 1 回だけ表示） ──
            if st.session_state.pop("show_toast", False):
                st.toast("✅ レポートの生成が完了しました", icon="✅")

            # ── ダウンロードセクション（アンカー + 強調カード） ──
            st.markdown("""
            <div class="download-section" id="dl-anchor">
                <div class="download-title">✅ IIJブランド・統合レポート完成</div>
                <div class="download-subtitle">
                    すべての解析が正常に終了しました。<br>
                    下のボタンからレポートをダウンロードしてください。
                </div>
            </div>
            """, unsafe_allow_html=True)

            # pulse-button ラッパーで Streamlit DL ボタンを赤く脈動させる
            st.markdown('<div class="pulse-button">', unsafe_allow_html=True)
            col_l, col_c, col_r = st.columns([1, 2, 1])
            with col_c:
                st.download_button(
                    label="⬇　レポートをダウンロード",
                    data=st.session_state["pptx_bytes"],
                    file_name=f"IIJ_Project_Relay_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True,
                )
            st.markdown('</div>', unsafe_allow_html=True)  # .pulse-button 閉じ

            # ── 自動スクロール ──
            # components.html は独立した iframe 内で JS を確実に実行するため
            # Streamlit の再レンダリング後も安定して動作する。
            # parent.document でホスト側の DOM を参照し、#dl-anchor を正中線へスクロール。
            import streamlit.components.v1 as components
            components.html("""
            <script>
                (function () {
                    function scroll() {
                        var el = parent.document.getElementById('dl-anchor');
                        if (el) {
                            el.scrollIntoView({ behavior: 'smooth', block: 'center' });
                        } else {
                            setTimeout(scroll, 100);
                        }
                    }
                    setTimeout(scroll, 300);
                })();
            </script>
            """, height=0)

    st.markdown('</div>', unsafe_allow_html=True)
    render_footer()


# =============================================================================
# エントリーポイント
# =============================================================================
if __name__ == "__main__":
    main()
