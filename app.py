# ==============================================================================
#
#   Project Relay v2  ─  統合報告レポート自動生成システム
#   向平様専用 業務効率化ツール（Streamlit）
#   セキュリティゲート統合版 / UX 極限改善版
#
# ==============================================================================
#
#  【インストール】
#    pip install streamlit python-pptx openpyxl pdfplumber
#
#  【起動】
#    streamlit run app.py
#
#  【認証パスワード】
#    relay2026
#
#  【動作フロー】
#    起動
#     │
#     ├─ CSS 注入（認証画面にもネイビー・ゴールドを適用）
#     │
#     ├─ render_auth_gate()
#     │    ├─ 未認証 → 認証画面を表示して False 返却 → 終了
#     │    └─ 認証済 → True 返却
#     │
#     └─ main()  ← 認証済みのときのみ発火
#          ├─ render_sidebar()           過去レポート一覧
#          ├─ render_hero()              ヒーローヘッダー
#          ├─ ファイルアップロード       st.file_uploader
#          ├─ process_files()            テキスト抽出
#          ├─ classify()                 キーワード分類
#          ├─ make_pptx()               PPTX 生成
#          └─ render_success_banner…()  完了バナー＋DL ボタン
#
# ==============================================================================

import io
import time
from datetime import datetime
from pathlib import Path

import streamlit as st

# ── サードパーティ（pip がなくてもクラッシュしない設計） ──────────────────────
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


# ==============================================================================
# ページ設定  ─  Streamlit への最初の呼び出しでなければならない
# ==============================================================================
st.set_page_config(
    page_title="Project Relay | 統合レポート生成",
    page_icon="⬡",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ==============================================================================
# 定数
# ==============================================================================

HISTORY_DIR = Path("./history")
HISTORY_DIR.mkdir(exist_ok=True)

# ── 認証パスワード（変更する場合はここだけ書き換えてください） ──
CORRECT_PW = "relay2026"

# ── キーワード分類辞書 ────────────────────────────────────────────────────────
CATEGORY_KEYWORDS: dict[str, list[str]] = {
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

CATEGORY_ICONS: dict[str, str] = {
    "今月の成果":       "🏆",
    "数値指標":         "📊",
    "発生した課題":     "⚠️",
    "次月の予定":       "📅",
    "その他・参考情報": "📎",
}

# ── PPTX スライドカラー ────────────────────────────────────────────────────────
if PPTX_AVAILABLE:
    C_DARK   = RGBColor(0x1E, 0x27, 0x61)   # ネイビー  （タイトル背景）
    C_ACCENT = RGBColor(0xCA, 0xDC, 0xFC)   # アイスブルー（アクセント）
    C_WHITE  = RGBColor(0xFF, 0xFF, 0xFF)   # 白
    C_LIGHT  = RGBColor(0xF4, 0xF6, 0xFF)   # 薄ブルー  （コンテンツ背景）
    C_BODY   = RGBColor(0x1E, 0x27, 0x61)   # 本文ネイビー
    C_CITE   = RGBColor(0x99, 0x99, 0xAA)   # 出典グレー


# ==============================================================================
# グローバル CSS
#
#  認証画面・メイン UI 両方のスタイルを一ファイルに集約しています。
#  エントリーポイント冒頭で st.markdown(CSS) を一度だけ呼びます。
#  main() 内での再注入は「単体テスト時の保険」であり、Streamlit は重複を無視します。
# ==============================================================================
CSS = """
<style>

/* ============================================================
   Google Fonts
   ============================================================ */
@import url('https://fonts.googleapis.com/css2?family=Noto+Serif+JP:wght@300;400;600&family=Noto+Sans+JP:wght@300;400;500;700&family=Cormorant+Garamond:ital,wght@0,300;0,600;1,300&display=swap');


/* ============================================================
   CSS 変数（カラーパレット）
   ============================================================ */
:root {
    --navy:      #080F24;
    --navy-2:    #0D1B3E;
    --navy-3:    #1A2848;
    --gold:      #C9A84C;
    --gold-lt:   #E8C97A;
    --off-white: #EEF1F8;
    --muted:     #6B7A9F;
    --border:    rgba(201, 168, 76, 0.20);
    --card:      #111D38;
    --success:   #22c55e;
    --white:     #FFFFFF;
}


/* ============================================================
   Streamlit リセット & ベース
   ============================================================ */
html, body,
[data-testid="stAppViewContainer"],
[data-testid="stMain"] {
    background-color: var(--navy) !important;
    color:            var(--off-white) !important;
    font-family:      'Noto Sans JP', sans-serif;
}

[data-testid="stHeader"]        { background: transparent !important; }
[data-testid="stVerticalBlock"] { gap: 0 !important; }
.block-container                { padding: 0 !important; max-width: 100% !important; }
.appview-container .main .block-container { padding-top: 0 !important; }


/* ============================================================
   サイドバー（メイン UI 用・認証後に表示）
   ============================================================ */
[data-testid="stSidebar"] {
    background:   #04091A !important;
    border-right: 1px solid var(--border) !important;
}
[data-testid="stSidebarContent"] { padding: 0 !important; }


/* ============================================================
   ★ 認証ゲート専用スタイル
      クラス名を auth- でスコープ管理し、メイン UI と完全に分離します。
   ============================================================ */

/* 認証画面中はサイドバーとトグルボタンを完全封鎖 */
.auth-sidebar-kill [data-testid="stSidebar"],
.auth-sidebar-kill [data-testid="collapsedControl"] {
    display:    none   !important;
    visibility: hidden !important;
    width:      0      !important;
    min-width:  0      !important;
}

/* 全画面センタリングのステージ */
.auth-stage {
    min-height:      100vh;
    display:         flex;
    align-items:     center;
    justify-content: center;
    padding:         40px 20px;
    background:
        radial-gradient(ellipse at 65% 25%, rgba(201,168,76,0.065) 0%, transparent 60%),
        linear-gradient(160deg, #03070F 0%, #080F24 55%, #0A1530 100%);
}

/* 認証カード本体 */
.auth-card {
    width:         100%;
    max-width:     400px;
    background:    var(--card);
    border:        1px solid var(--border);
    border-radius: 8px;
    padding:       52px 44px 46px;
    position:      relative;
    overflow:      hidden;
    box-shadow:
        0 32px 80px rgba(0, 0, 0, 0.65),
        0  0   0   1px rgba(201, 168, 76, 0.07),
        inset 0 1px 0   rgba(201, 168, 76, 0.20);
    animation: authCardIn 0.55s cubic-bezier(0.22, 1, 0.36, 1) both;
}

/* カード上辺のゴールドグラデーションライン */
.auth-card::before {
    content:    '';
    position:   absolute;
    top: 0; left: 0; right: 0;
    height:     1px;
    background: linear-gradient(90deg, transparent 0%, var(--gold) 50%, transparent 100%);
}

/* カード右上のアンビエント光 */
.auth-card::after {
    content:       '';
    position:      absolute;
    top: -110px; right: -110px;
    width:         280px;
    height:        280px;
    border-radius: 50%;
    background:    radial-gradient(circle, rgba(201,168,76,0.058) 0%, transparent 70%);
    pointer-events: none;
}

@keyframes authCardIn {
    from { opacity: 0; transform: translateY(28px) scale(0.96); }
    to   { opacity: 1; transform: translateY(0)    scale(1);    }
}

/* ── ブランドロゴ行 ── */
.auth-logo {
    font-family:     'Cormorant Garamond', serif;
    font-size:       10.5px;
    font-weight:     300;
    letter-spacing:  0.50em;
    text-transform:  uppercase;
    color:           var(--gold);
    text-align:      center;
    margin-bottom:   30px;
    display:         flex;
    align-items:     center;
    justify-content: center;
    gap:             12px;
}
.auth-logo::before,
.auth-logo::after {
    content:    '';
    display:    inline-block;
    width:      22px;
    height:     1px;
    background: linear-gradient(90deg, transparent, var(--gold));
}
.auth-logo::after { transform: scaleX(-1); }

/* ── カードタイトル ── */
.auth-title {
    font-family:   'Noto Serif JP', serif;
    font-size:     24px;
    font-weight:   600;
    color:         var(--white);
    text-align:    center;
    line-height:   1.3;
    margin-bottom: 8px;
}

/* ── カードサブテキスト ── */
.auth-sub {
    font-size:     12px;
    color:         var(--muted);
    text-align:    center;
    letter-spacing:0.04em;
    line-height:   1.88;
    margin-bottom: 36px;
}

/* ── 入力欄のカスタムラベル ── */
.auth-label {
    display:        block;
    font-family:    'Cormorant Garamond', serif;
    font-size:      9.5px;
    letter-spacing: 0.42em;
    text-transform: uppercase;
    color:          var(--gold);
    margin-bottom:  8px;
}

/* ── text_input をゴールドテーマに上書き ── */
[data-testid="stTextInput"] input {
    background:      #07101E                          !important;
    border:          1px solid rgba(201,168,76,0.32)  !important;
    border-radius:   3px                              !important;
    color:           var(--off-white)                 !important;
    font-family:     'Noto Sans JP', sans-serif       !important;
    font-size:       13.5px                           !important;
    letter-spacing:  0.05em                           !important;
    padding:         12px 16px                        !important;
    caret-color:     var(--gold)                      !important;
    transition:      border-color 0.25s, box-shadow 0.25s !important;
}
[data-testid="stTextInput"] input:focus {
    border-color:    rgba(201,168,76,0.80)            !important;
    box-shadow:      0 0 0 3px rgba(201,168,76,0.11)  !important;
    outline:         none                             !important;
}
[data-testid="stTextInput"] input::placeholder {
    color:           rgba(107,122,159,0.72)           !important;
    font-size:       12.5px                           !important;
}
/* Streamlit 自動生成ラベルは非表示（.auth-label を使うため） */
[data-testid="stTextInput"] label { display: none !important; }

/* ── 認証ボタン（.auth-wrap 内のボタンにのみ適用） ── */
.auth-wrap .stButton > button {
    background:     linear-gradient(135deg, #B58A28 0%, #E8C97A 50%, #B58A28 100%) !important;
    color:          #040C18                           !important;
    font-family:    'Noto Sans JP', sans-serif        !important;
    font-weight:    700                               !important;
    font-size:      13px                              !important;
    letter-spacing: 0.13em                            !important;
    border:         none                              !important;
    border-radius:  3px                               !important;
    padding:        13px 28px                         !important;
    width:          100%                              !important;
    box-shadow:     0 0 26px rgba(201,168,76,0.40),
                    0 2px 14px rgba(0,0,0,0.42)       !important;
    transition:     all 0.25s ease                    !important;
    margin-top:     6px                               !important;
}
.auth-wrap .stButton > button:hover {
    box-shadow:  0 0 42px rgba(201,168,76,0.68),
                 0 4px 18px rgba(0,0,0,0.52)          !important;
    transform:   translateY(-1px)                     !important;
}

/* ── エラーメッセージ（shake アニメーション付き） ── */
.auth-error {
    background:     rgba(239, 68, 68, 0.09);
    border:         1px solid rgba(239, 68, 68, 0.30);
    border-radius:  3px;
    padding:        10px 14px;
    font-size:      12px;
    color:          #fca5a5;
    letter-spacing: 0.04em;
    text-align:     center;
    margin-top:     14px;
    animation:      authShake 0.38s ease;
}
@keyframes authShake {
    0%, 100% { transform: translateX(0);    }
    20%,  60% { transform: translateX(-6px); }
    40%,  80% { transform: translateX( 6px); }
}

/* ── カード内フッター注釈 ── */
.auth-footer {
    font-size:      9.5px;
    color:          rgba(107,122,159, 0.44);
    text-align:     center;
    margin-top:     30px;
    letter-spacing: 0.15em;
}


/* ============================================================
   ★ メイン UI スタイル（既存・認証後に表示）
   ============================================================ */

/* ── ヒーローヘッダー ── */
.hero {
    background:    linear-gradient(140deg, #04091A 0%, #0D1B3E 55%, #162448 100%);
    border-bottom: 1px solid var(--border);
    padding:       48px 72px 40px;
    position:      relative;
    overflow:      hidden;
}
.hero::before {
    content:       '';
    position:      absolute;
    top: -80px; right: -80px;
    width:         380px;
    height:        380px;
    border-radius: 50%;
    background:    radial-gradient(circle, rgba(201,168,76,0.07) 0%, transparent 68%);
    pointer-events: none;
}
.hero::after {
    content:    '';
    position:   absolute;
    bottom: 0; left: 0; right: 0;
    height:     1px;
    background: linear-gradient(90deg, transparent 0%, var(--gold) 50%, transparent 100%);
}
.hero-eyebrow {
    font-family:    'Cormorant Garamond', serif;
    font-size:      12px;
    font-weight:    300;
    letter-spacing: 0.38em;
    color:          var(--gold);
    text-transform: uppercase;
    margin-bottom:  14px;
}
.hero-title {
    font-family:   'Noto Serif JP', serif;
    font-size:     clamp(28px, 3.5vw, 46px);
    font-weight:   600;
    color:         var(--white);
    line-height:   1.2;
    margin-bottom: 10px;
}
.hero-title span {
    color:       var(--gold-lt);
    font-weight: 300;
    font-style:  italic;
}
.hero-sub {
    font-size:      13px;
    font-weight:    300;
    color:          var(--muted);
    letter-spacing: 0.04em;
    line-height:    1.9;
}

/* ── メインコンテンツ ラッパー ── */
.main-wrap {
    padding:   36px 72px;
    max-width: 1000px;
    margin:    0 auto;
}

/* ── セクションラベル ── */
.sec-label {
    font-family:    'Cormorant Garamond', serif;
    font-size:      10.5px;
    letter-spacing: 0.42em;
    text-transform: uppercase;
    color:          var(--gold);
    margin-bottom:  18px;
    display:        flex;
    align-items:    center;
    gap:            10px;
}
.sec-label::after {
    content:    '';
    flex:       1;
    height:     1px;
    background: var(--border);
}

/* ── フォーマットバッジ ── */
.badges {
    display:   flex;
    gap:       8px;
    flex-wrap: wrap;
    margin:    10px 0 20px;
}
.badge {
    background:     rgba(201,168,76,0.07);
    border:         1px solid rgba(201,168,76,0.28);
    border-radius:  2px;
    padding:        4px 11px;
    font-size:      10.5px;
    letter-spacing: 0.13em;
    color:          var(--gold-lt);
    font-family:    'Courier New', monospace;
    text-transform: uppercase;
}

/* ── ファイルアップローダー ── */
[data-testid="stFileUploader"] {
    background:    var(--card)              !important;
    border:        1px solid var(--border)  !important;
    border-radius: 4px                      !important;
    transition:    border-color 0.3s;
}
[data-testid="stFileUploader"]:hover {
    border-color: rgba(201,168,76,0.52)     !important;
}
[data-testid="stFileUploaderDropzoneInstructions"] {
    color: var(--muted)                     !important;
}
[data-testid="stFileUploaderFile"] {
    background:    rgba(201,168,76,0.07)    !important;
    border:        1px solid var(--border)  !important;
    border-radius: 2px                      !important;
    color:         var(--off-white)         !important;
}

/* ── 生成ボタン（透明・ゴールドボーダー） ── */
.stButton > button {
    background:     transparent              !important;
    color:          var(--gold)              !important;
    font-family:    'Noto Sans JP', sans-serif !important;
    font-size:      13.5px                   !important;
    font-weight:    500                      !important;
    letter-spacing: 0.10em                   !important;
    border:         1px solid var(--gold)    !important;
    border-radius:  2px                      !important;
    padding:        12px 32px               !important;
    width:          100%                     !important;
    transition:     all 0.25s ease           !important;
}
.stButton > button:hover {
    background: rgba(201,168,76,0.10)        !important;
    box-shadow: 0 0 22px rgba(201,168,76,0.22) !important;
}

/* ── プログレスバー ── */
.stProgress > div {
    background:    rgba(201,168,76,0.10)     !important;
    border-radius: 2px                       !important;
    height:        3px                       !important;
}
.stProgress > div > div {
    background:    linear-gradient(90deg, var(--gold), var(--gold-lt)) !important;
    border-radius: 2px                       !important;
}

/* ── 成功バナー ── */
.success-banner {
    background:    linear-gradient(135deg, #071507 0%, #0A2010 100%);
    border:        1.5px solid var(--success);
    border-radius: 6px;
    padding:       26px 36px;
    display:       flex;
    align-items:   center;
    gap:           24px;
    margin-bottom: 24px;
    box-shadow:    0 0 42px rgba(34, 197, 94, 0.14);
    animation:     successPop 0.45s cubic-bezier(0.34, 1.56, 0.64, 1) both;
}
@keyframes successPop {
    from { opacity: 0; transform: scale(0.97) translateY(-8px); }
    to   { opacity: 1; transform: scale(1)    translateY(0);    }
}
.success-left {
    display:     flex;
    align-items: center;
    gap:         16px;
}
.success-check {
    width:           46px;
    height:          46px;
    border-radius:   50%;
    background:      rgba(34, 197, 94, 0.14);
    border:          1.5px solid var(--success);
    display:         flex;
    align-items:     center;
    justify-content: center;
    font-size:       20px;
    flex-shrink:     0;
}
.success-title {
    font-family:   'Noto Serif JP', serif;
    font-size:     19px;
    font-weight:   600;
    color:         var(--white);
    margin-bottom: 4px;
}
.success-meta {
    font-size:      11.5px;
    color:          #86efac;
    letter-spacing: 0.04em;
}

/* ── ダウンロードボタン（ゴールド・パルスグロー） ── */
[data-testid="stDownloadButton"] > button {
    background:      linear-gradient(135deg, #C9A84C 0%, #E8C97A 50%, #C9A84C 100%) !important;
    background-size: 200%                      !important;
    color:           #04101C                   !important;
    font-family:     'Noto Sans JP', sans-serif !important;
    font-weight:     700                       !important;
    font-size:       15px                      !important;
    letter-spacing:  0.08em                    !important;
    border:          none                      !important;
    border-radius:   3px                       !important;
    padding:         16px 52px                 !important;
    width:           100%                      !important;
    box-shadow:      0 0 32px rgba(201,168,76,0.62),
                     0 4px 18px rgba(0,0,0,0.46)  !important;
    transition:      all 0.30s ease            !important;
    animation:       dlPulse 2.6s ease-in-out infinite !important;
}
[data-testid="stDownloadButton"] > button:hover {
    box-shadow:           0 0 54px rgba(201,168,76,0.88),
                          0 8px 28px rgba(0,0,0,0.52)  !important;
    transform:            translateY(-2px)              !important;
    animation-play-state: paused                        !important;
}
@keyframes dlPulse {
    0%, 100% { box-shadow: 0 0 32px rgba(201,168,76,0.62), 0 4px 18px rgba(0,0,0,0.46); }
    50%      { box-shadow: 0 0 54px rgba(201,168,76,0.88), 0 4px 18px rgba(0,0,0,0.46); }
}

/* ── 統計カード（3列） ── */
.stat-row {
    display:               grid;
    grid-template-columns: repeat(3, 1fr);
    gap:                   14px;
    margin:                18px 0 26px;
}
.stat-card {
    background:    var(--card);
    border:        1px solid var(--border);
    border-radius: 4px;
    padding:       22px 26px;
    position:      relative;
    overflow:      hidden;
}
.stat-card::before {
    content:    '';
    position:   absolute;
    top: 0; left: 0;
    width:      3px;
    height:     100%;
    background: linear-gradient(180deg, var(--gold), transparent);
}
.stat-n {
    font-family:   'Cormorant Garamond', serif;
    font-size:     40px;
    font-weight:   600;
    color:         var(--gold-lt);
    line-height:   1;
    margin-bottom: 5px;
}
.stat-l {
    font-size:      10.5px;
    letter-spacing: 0.16em;
    color:          var(--muted);
    text-transform: uppercase;
}

/* ── カテゴリカード（2列グリッド） ── */
.cat-grid {
    display:               grid;
    grid-template-columns: repeat(2, 1fr);
    gap:                   10px;
    margin:                14px 0 22px;
}
.cat-card {
    background:    var(--card);
    border:        1px solid var(--border);
    border-radius: 4px;
    padding:       16px 18px;
    display:       flex;
    align-items:   flex-start;
    gap:           12px;
}
.cat-icon {
    width:           33px;
    height:          33px;
    border-radius:   50%;
    background:      rgba(201,168,76,0.10);
    border:          1px solid var(--border);
    display:         flex;
    align-items:     center;
    justify-content: center;
    font-size:       14px;
    flex-shrink:     0;
    margin-top:      2px;
}
.cat-name {
    font-family:   'Noto Serif JP', serif;
    font-size:     13px;
    font-weight:   600;
    color:         var(--off-white);
    margin-bottom: 3px;
}
.cat-cnt {
    font-size:   11px;
    color:       var(--gold);
    font-weight: 500;
}
.cat-prev {
    font-size:           10.5px;
    color:               var(--muted);
    margin-top:          5px;
    line-height:         1.6;
    display:             -webkit-box;
    -webkit-line-clamp:  2;
    -webkit-box-orient:  vertical;
    overflow:            hidden;
}

/* ── ログコンソール ── */
.log-con {
    background:    #040912;
    border:        1px solid rgba(201,168,76,0.10);
    border-radius: 4px;
    padding:       16px 20px;
    font-family:   'Courier New', monospace;
    font-size:     11.5px;
    color:         #4ade80;
    line-height:   2.1;
    max-height:    190px;
    overflow-y:    auto;
    margin:        12px 0;
}
.ll       { animation: logFadeIn 0.25s ease; }
.ll.warn  { color: #fbbf24; }
.ll.err   { color: #f87171; }
@keyframes logFadeIn {
    from { opacity: 0; transform: translateX(-4px); }
    to   { opacity: 1; transform: translateX(0);    }
}

/* ── 区切り線 ── */
hr {
    border:     none                       !important;
    border-top: 1px solid var(--border)    !important;
    margin:     26px 0                     !important;
}

/* ── サイドバー：ヘッダー ── */
.sb-head {
    background:    linear-gradient(180deg, #030812, #0A1228);
    border-bottom: 1px solid var(--border);
    padding:       26px 18px 18px;
}
.sb-title {
    font-family:   'Noto Serif JP', serif;
    font-size:     14px;
    font-weight:   600;
    color:         var(--white);
    margin-bottom: 3px;
}
.sb-sub {
    font-size:      10.5px;
    color:          var(--muted);
    letter-spacing: 0.05em;
}

/* ── サイドバー：履歴アイテム ── */
.sb-item {
    padding:       12px 18px;
    border-bottom: 1px solid rgba(201,168,76,0.07);
}
.sb-name {
    font-size:     11.5px;
    color:         var(--off-white);
    margin-bottom: 3px;
    word-break:    break-all;
}
.sb-meta {
    font-size: 10px;
    color:     var(--muted);
}
.sb-empty {
    padding:     26px 18px;
    font-size:   11.5px;
    color:       var(--muted);
    text-align:  center;
    line-height: 2.2;
}

/* サイドバー内ダウンロードボタン（コンパクト・グロー無効） */
section[data-testid="stSidebar"] [data-testid="stDownloadButton"] > button {
    background:     rgba(201,168,76,0.08)    !important;
    color:          var(--gold-lt)           !important;
    font-size:      11px                     !important;
    font-weight:    500                      !important;
    padding:        6px 14px                 !important;
    border:         1px solid var(--border)  !important;
    border-radius:  2px                      !important;
    letter-spacing: 0.06em                   !important;
    box-shadow:     none                     !important;
    animation:      none                     !important;
    margin-bottom:  8px                      !important;
}
section[data-testid="stSidebar"] [data-testid="stDownloadButton"] > button:hover {
    background: rgba(201,168,76,0.16)        !important;
    transform:  none                         !important;
}

/* ── フッター ── */
.footer {
    border-top:     1px solid var(--border);
    padding:        18px 72px;
    text-align:     center;
    font-size:      10.5px;
    color:          var(--muted);
    letter-spacing: 0.10em;
}

</style>
"""


# ==============================================================================
# ファイル読み込み関数  ─  BytesIO 対応・ゼロクラッシュ設計
# ==============================================================================

def read_pptx_bytes(data: bytes, name: str) -> str:
    """PPTX ファイル（バイト列）から全スライドのテキストを抽出します。"""
    lines = [f"【出典：{name}】"]
    try:
        prs = Presentation(io.BytesIO(data))
        for i, slide in enumerate(prs.slides, 1):
            texts = [
                para.text.strip()
                for shape in slide.shapes
                if shape.has_text_frame
                for para in shape.text_frame.paragraphs
                if para.text.strip()
            ]
            if texts:
                lines.append(f"--- スライド {i} ---")
                lines.extend(texts)
    except Exception as e:
        lines.append(f"（読み込みエラー: {e}）")
    return "\n".join(lines) + "\n"


def read_xlsx_bytes(data: bytes, name: str) -> str:
    """XLSX ファイル（バイト列）から全シートのセルデータを抽出します。"""
    lines = [f"【出典：{name}】"]
    try:
        wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            lines.append(f"--- シート: {sheet_name} ---")
            for row in ws.iter_rows():
                row_data = [
                    str(cell.value).strip()
                    for cell in row
                    if cell.value is not None
                ]
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
                text = page.extract_text()
                if text and text.strip():
                    lines.append(f"--- ページ {i} ---")
                    lines.append(text.strip())
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


# ==============================================================================
# 処理ロジック
# ==============================================================================

def process_files(uploaded_files) -> tuple[list[dict], list[str]]:
    """
    アップロードされたファイルを読み込んでテキストを抽出します。

    Returns
    -------
    tuple[list[dict], list[str]]
        ファイルデータリスト（filename / text）と処理ログリスト
    """
    READERS: dict = {
        ".pptx": read_pptx_bytes,
        ".xlsx": read_xlsx_bytes,
        ".pdf":  read_pdf_bytes,
        ".txt":  read_txt_bytes,
    }
    results: list[dict] = []
    logs:    list[str]  = []

    for uf in uploaded_files:
        ext = Path(uf.name).suffix.lower()
        if ext not in READERS:
            logs.append(f"⏭  スキップ: {uf.name}（非対応フォーマット）")
            continue
        logs.append(f"📄  {uf.name} を読み込み中...")
        try:
            text = READERS[ext](uf.read(), uf.name)
            results.append({"filename": uf.name, "text": text})
            logs.append(f"✅  {uf.name} 読み込み完了")
        except Exception as e:
            logs.append(f"❌  {uf.name} 失敗: {e}")

    return results, logs


def classify(file_data_list: list[dict]) -> dict:
    """
    キーワードマッチングでテキストを最大 5 カテゴリに分類します。
    要約は行わず、原文の情報をそのまま整理します。

    Returns
    -------
    dict  カテゴリ名 → [{"text": str, "source": str}] のリスト
    """
    cats:  dict       = {k: [] for k in CATEGORY_KEYWORDS}
    other: list[dict] = []

    for fd in file_data_list:
        fn = fd["filename"]
        for line in fd["text"].split("\n"):
            s = line.strip()
            # 空行・ヘッダー行・出典行はスキップ
            if not s or s.startswith("---") or s.startswith("【出典"):
                continue
            matched = False
            for cat, kws in CATEGORY_KEYWORDS.items():
                if any(kw in s for kw in kws):
                    cats[cat].append({"text": s, "source": fn})
                    matched = True
                    break
            if not matched and len(s) > 5:
                other.append({"text": s, "source": fn})

    if other:
        cats["その他・参考情報"] = other

    return cats


# ==============================================================================
# PPTX 生成
# ==============================================================================

def _bg(slide, color: "RGBColor") -> None:
    """スライドの背景色をソリッド塗りで設定します。"""
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def _tb(
    slide,
    text:  str,
    l:     float,
    t:     float,
    w:     float,
    h:     float,
    size:  int,
    bold:  bool             = False,
    color: "RGBColor | None" = None,
    align                   = None,
) -> None:
    """
    スライドにテキストボックスを追加するヘルパー関数。

    Parameters
    ----------
    slide   : slide オブジェクト
    text    : 表示文字列（\n で改行）
    l, t    : 左・上位置（インチ）
    w, h    : 幅・高さ（インチ）
    size    : フォントサイズ（pt）
    bold    : 太字にするか
    color   : RGBColor（None の場合はデフォルト黒）
    align   : PP_ALIGN 定数（None の場合は LEFT）
    """
    if align is None:
        align = PP_ALIGN.LEFT
    tb          = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf          = tb.text_frame
    tf.word_wrap = True
    p           = tf.paragraphs[0]
    p.alignment = align
    run         = p.add_run()
    run.text    = text
    run.font.size = Pt(size)
    run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color


def _build_title_slide(prs: "Presentation", today_str: str) -> None:
    """タイトルスライド（1 枚目）を生成します。"""
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(sl, C_DARK)

    # 左端アクセントバー（アイスブルー縦線）
    bar = sl.shapes.add_shape(1, Inches(0), Inches(0), Inches(0.15), Inches(7.5))
    bar.fill.solid()
    bar.fill.fore_color.rgb = C_ACCENT
    bar.line.fill.background()

    # メインタイトル
    _tb(
        sl,
        "【自動生成】\nチーム進捗報告\n統合レポート",
        l=0.5, t=1.5, w=9.0, h=3.0,
        size=40, bold=True, color=C_WHITE,
        align=PP_ALIGN.LEFT,
    )

    # 向平様へのご挨拶テキスト
    _tb(
        sl,
        "向平様\n\n"
        "お忙しい中、ご確認いただきありがとうございます。\n"
        "本レポートは各部門からの報告資料を自動統合・整理したものです。\n"
        "情報の正確性を最優先し、原文を整理して掲載しております。",
        l=0.5, t=4.5, w=8.5, h=2.5,
        size=14, color=C_ACCENT,
        align=PP_ALIGN.LEFT,
    )

    # 右下の生成日時
    _tb(
        sl,
        f"生成日時: {today_str}　Project Relay 自動生成",
        l=0.5, t=7.0, w=9.0, h=0.4,
        size=10, color=C_CITE,
        align=PP_ALIGN.RIGHT,
    )


def _build_index_slide(prs: "Presentation", cats: dict) -> None:
    """目次スライド（2 枚目）を生成
