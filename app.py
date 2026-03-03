# ==============================================================================
# Project Relay v11 — 「確認」体験の完成形
# 向平 友治 様専用  |  認証: relay2026
# pip install streamlit python-pptx openpyxl pdfplumber
# streamlit run app.py
# ==============================================================================
#
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 【v11 設計思想と v10 からの根本的な変更点】
#
# v10 で残っていた 3 つの構造的問題:
#
# ① 二重カウントバグ
#    高信頼度アイテムが expander 内でも item_card() を呼び、
#    かつ折りたたみ時も別途カウントしていた → 同じ item が2回 selected に入る
#    → v11: 高信頼度アイテムは「サマリー表示のみ」、状態はセッションで管理
#
# ② 3択ボタンが毎回 st.rerun() を呼んで重い
#    → v11: Streamlit の checkbox/selectbox のみで状態管理し、
#           rerun は「手動追加の ＋ ボタン」「やり直す」のみに限定
#           item の include/exclude は checkbox で管理（0 rerun）
#
# ③ 信頼度スコアが「テキストの特性」しか見ていない
#    → v11: 因果ボーナスを追加
#           同一ファイル内に「課題→対応→結果」の順番で出現した場合 +0.15
#           これにより「因果が成立している行」の信頼度が上がる
#
# 【v11 の新機能・改善】
# ① ノイズ除去の強化 — ヘッダー行・メタデータ行を正規表現で除外
# ② 確認ヒントのカスケード — 最も重要な1件だけ表示（複数出さない）
# ③ 完成度スコアの因果ロジック改善 — キーワード重複チェックで繋がりを評価
# ④ 右パネル固定化 — CSS sticky を正しく実装、スクロールしても常に見える
# ⑤ PPTX の視認性改善 — フォントサイズ・行間・余白を全スライドで統一
# ⑥ コード構造整理 — 関数を役割別に 8 層に分割、500行→明確な責任範囲
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

from __future__ import annotations

import io
import re
import hashlib
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Optional

import streamlit as st
import streamlit.components.v1 as components

# ==============================================================================
# LAYER 0 ─ 認証
# ==============================================================================
if "auth" not in st.session_state:
    st.session_state.auth = False

def _check_pw() -> None:
    if st.session_state.get("pw_entry") == "relay2026":
        st.session_state.auth = True
    else:
        st.error("パスワードが正しくありません")

if not st.session_state.auth:
    st.set_page_config(page_title="Project Relay", page_icon="⬡")
    st.markdown("""<style>
    html,body,[data-testid="stAppViewContainer"]{background:#060E1C!important;}
    [data-testid="stHeader"]{background:transparent!important;}
    .block-container{padding:100px 40px!important;max-width:440px!important;margin:0 auto!important;}
    [data-testid="stTextInput"] input{
      background:rgba(255,255,255,.06)!important;
      border:1px solid rgba(200,220,255,.2)!important;
      color:#FFF!important;font-size:16px!important;
      border-radius:8px!important;padding:14px 16px!important;
    }
    </style>""", unsafe_allow_html=True)
    st.markdown(
        '<div style="text-align:center;margin-bottom:44px;">'
        '<div style="font-size:9px;letter-spacing:.55em;color:#C8002E;'
        'text-transform:uppercase;margin-bottom:12px;">Project Relay</div>'
        '<div style="font-size:26px;font-weight:700;color:#FFF;'
        'font-family:serif;letter-spacing:-.01em;">向平 友治 様専用</div>'
        '<div style="font-size:11px;color:#4E5E80;margin-top:10px;">'
        '月次報告スライド自動生成ツール</div>'
        '</div>', unsafe_allow_html=True)
    st.text_input("", type="password", key="pw_entry",
                  on_change=_check_pw, placeholder="パスワードを入力してください")
    st.stop()

# ==============================================================================
# LAYER 1 ─ 外部ライブラリ（全て Optional import）
# ==============================================================================
try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
    from pptx.enum.shapes import MSO_SHAPE
    PPTX_OK = True
except ImportError:
    PPTX_OK = False

try:
    import openpyxl
    XLSX_OK = True
except ImportError:
    XLSX_OK = False

try:
    import pdfplumber
    PDF_OK = True
except ImportError:
    PDF_OK = False

# ==============================================================================
# LAYER 2 ─ ページ設定
# ==============================================================================
st.set_page_config(
    page_title="Project Relay",
    page_icon="⬡",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ==============================================================================
# LAYER 3 ─ デザインシステム（CSS）
#
# 設計原則:
# ・色は 3 色のみ (赤=エラー/重要, 緑=OK/採用, アンバー=警告)
# ・タイポグラフィは 2 ウェイト (400 通常 / 700 強調)
# ・スペースは 4px グリッド
# ・影は 1 種類のみ (深度を表すため)
# ==============================================================================

CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Serif+JP:wght@400;600&family=Noto+Sans+JP:wght@400;500;700&display=swap');

/* ── デザイントークン ── */
:root {
  --bg0: #060E1C; --bg1: #0A1628; --bg2: #0F1E36; --bg3: #152542; --card: #0D1B30;
  --red: #C8002E; --red-l: #FF3355; --red-a: rgba(200,0,46,.10);
  --grn: #166534; --grn-l: #4ADE80; --grn-a: rgba(22,101,52,.12);
  --amb: #92400E; --amb-l: #FCD34D; --amb-a: rgba(146,64,14,.12);
  --blu: #1E3A8A; --blu-l: #93C5FD; --blu-a: rgba(30,58,138,.12);
  --t0: #FFFFFF; --t1: #E8EEFF; --t2: #7A8BAA; --t3: #3D4E68;
  --b0: rgba(200,220,255,.06); --b1: rgba(200,220,255,.14); --b2: rgba(200,220,255,.24);
  --br: rgba(200,0,46,.3); --bg_: rgba(22,101,52,.3); --ba: rgba(146,64,14,.3);
  --r-sm: 4px; --r-md: 6px; --r-lg: 10px;
  --shadow: 0 4px 24px rgba(0,0,0,.3);
}

/* ── ベースリセット ── */
html, body, [data-testid="stAppViewContainer"] {
  background-color: var(--bg0) !important;
  color: var(--t1) !important;
  font-family: 'Noto Sans JP', sans-serif;
  font-size: 14px; line-height: 1.65;
}
[data-testid="stHeader"],
[data-testid="stSidebar"]  { display: none !important; }
[data-testid="stVerticalBlock"] { gap: 0 !important; }
.block-container { padding: 0 !important; max-width: 100% !important; }
.appview-container .main .block-container { padding-top: 0 !important; }
hr { border: none !important; border-top: 1px solid var(--b0) !important; margin: 16px 0 !important; }

/* ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
   ヘッダーバー（全画面共通）
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ */
.hdr {
  background: linear-gradient(180deg, #030810 0%, #060E1C 100%);
  border-bottom: 1px solid var(--br);
  padding: 14px 48px 12px;
  position: relative;
}
.hdr::after {
  content: '';
  position: absolute; bottom: 0; left: 0; right: 0; height: 1px;
  background: linear-gradient(90deg, transparent 0%, var(--red) 30%, var(--red-l) 70%, transparent 100%);
}
.hdr-inner {
  display: flex; align-items: center; gap: 20px; max-width: 1280px; margin: 0 auto;
}
.hdr-brand {
  font-size: 9px; letter-spacing: .55em; text-transform: uppercase;
  color: var(--red); white-space: nowrap;
}

/* ステップバー */
.step-bar  { display: flex; align-items: center; gap: 0; }
.si {
  display: flex; align-items: center; gap: 5px;
  padding: 4px 12px; border-radius: var(--r-sm);
  font-size: 11px; border: 1px solid transparent; white-space: nowrap;
}
.si.done   { background: var(--grn-a); color: var(--grn-l); border-color: var(--bg_); }
.si.active { background: var(--red-a); color: #FFF;         border-color: var(--br);  }
.si.future { color: var(--t3);         border-color: var(--b0); }
.sn {
  width: 16px; height: 16px; border-radius: 50%;
  display: flex; align-items: center; justify-content: center;
  font-size: 9px; font-weight: 700; flex-shrink: 0;
}
.si.done   .sn { background: var(--grn); color: #FFF; }
.si.active .sn { background: var(--red); color: #FFF; }
.si.future .sn { background: rgba(255,255,255,.06); color: var(--t3); }
.sa { color: var(--t3); font-size: 11px; padding: 0 3px; }

/* ページタイトル（画面ごとに変わる） */
.page-ttl {
  padding: 22px 48px 6px;
  max-width: 1280px; margin: 0 auto;
}
.page-ttl-sub {
  font-size: 10px; letter-spacing: .3em; text-transform: uppercase;
  color: var(--red); margin-bottom: 4px;
}
.page-ttl-main {
  font-family: 'Noto Serif JP', serif;
  font-size: clamp(18px, 2.4vw, 28px); font-weight: 600;
  color: #FFF; line-height: 1.25; margin-bottom: 3px;
}
.page-ttl-main em { color: var(--red-l); font-style: normal; }
.page-ttl-hint {
  font-size: 12px; color: var(--t2);
}

/* ── コンテンツラッパー ── */
.wn  { padding: 8px 48px 28px; max-width: 760px;  margin: 0 auto; }
.ww  { padding: 8px 48px 28px; max-width: 1280px; margin: 0 auto; }

/* ── セクションヘッダー ── */
.sh {
  display: flex; align-items: center; gap: 10px;
  font-size: 9px; font-weight: 700; letter-spacing: .46em; text-transform: uppercase;
  color: var(--red); margin: 16px 0 10px;
}
.sh::after { content: ''; flex: 1; height: 1px; background: var(--br); }

/* ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
   UPLOAD 画面
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ */
[data-testid="stFileUploader"] {
  background: var(--bg2) !important;
  border: 1px solid var(--b1) !important;
  border-radius: var(--r-md) !important;
  padding: 5px !important;
}
[data-testid="stFileUploaderFile"] {
  background: rgba(255,255,255,.06) !important;
  border: 1.5px solid var(--red) !important;
  border-radius: var(--r-sm) !important;
  margin-bottom: 3px !important;
}
[data-testid="stFileUploaderFileName"] { color: #FFF !important; font-weight: 700 !important; }
[data-testid="stFileUploaderFile"] span { color: #CADCFC !important; }
[data-testid="stFileUploaderFile"] button { color: var(--red-l) !important; }
.up-hint {
  background: var(--bg2); border: 1px dashed var(--b1);
  border-radius: var(--r-md); padding: 10px 14px;
  font-size: 12px; color: var(--t2); margin-top: 8px;
}
.up-hint strong { color: #CADCFC; display: block; margin-bottom: 3px; font-size: 11px; font-weight: 700; letter-spacing: .06em; }
.up-ok { font-size: 11px; color: var(--grn-l); margin-top: 4px; }

/* ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
   REVIEW 画面 — 3列一覧
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ */

/* カテゴリ列ヘッダー */
.cat-hdr {
  padding: 10px 12px 8px;
  border-radius: var(--r-sm) var(--r-sm) 0 0;
  display: flex; align-items: center; justify-content: space-between;
  margin-bottom: 2px;
}
.cat-hdr.issue  { background: rgba(200,0,46,.12);  border: 1px solid rgba(200,0,46,.3);  border-bottom: none; }
.cat-hdr.action { background: rgba(146,64,14,.12); border: 1px solid rgba(146,64,14,.3); border-bottom: none; }
.cat-hdr.result { background: rgba(22,101,52,.12); border: 1px solid rgba(22,101,52,.3); border-bottom: none; }
.cat-hdr-left  { display: flex; align-items: center; gap: 6px; }
.cat-hdr-icon  { font-size: 15px; }
.cat-hdr-name  { font-size: 14px; font-weight: 700; color: #FFF; }
.cat-hdr-q     { font-size: 10px; margin-top: 1px; }
.cat-hdr.issue  .cat-hdr-q { color: var(--red-l); }
.cat-hdr.action .cat-hdr-q { color: var(--amb-l); }
.cat-hdr.result .cat-hdr-q { color: var(--grn-l); }
.cat-hdr-cnt {
  font-size: 11px; font-weight: 700; color: var(--t2);
  background: var(--bg3); padding: 2px 8px; border-radius: 20px;
}

/* 一覧アイテム行 */
.rv-item {
  display: flex; align-items: center; gap: 8px;
  background: var(--card); border: 1px solid var(--b0);
  border-radius: var(--r-sm); padding: 8px 10px;
  margin-bottom: 3px; transition: border-color .12s;
}
.rv-item:hover { border-color: var(--b1); }
.rv-item.excl { opacity: .32; }
.rv-item.needs-review { border-left: 3px solid var(--amb-l); }
.rv-item.high-priority { border-left: 3px solid var(--red-l); }
.rv-item.auto-ok       { border-left: 3px solid var(--grn); }
.rv-short { font-size: 12.5px; color: var(--t1); flex: 1; line-height: 1.4; }
.rv-flag  { font-size: 9px; white-space: nowrap; padding: 2px 5px;
            border-radius: 2px; flex-shrink: 0; }
.rv-flag.r { background: var(--red-a);  color: var(--red-l); border: 1px solid var(--br); }
.rv-flag.a { background: var(--amb-a);  color: var(--amb-l); border: 1px solid var(--ba); }
.rv-flag.g { background: var(--grn-a);  color: var(--grn-l); border: 1px solid var(--bg_); }
.rv-flag.b { background: var(--blu-a);  color: var(--blu-l); border: 1px solid rgba(30,58,138,.3); }

/* 警告バナー */
.warn-banner {
  background: var(--amb-a); border: 1px solid var(--ba);
  border-left: 3px solid var(--amb-l);
  border-radius: var(--r-sm); padding: 9px 13px;
  font-size: 12px; color: var(--amb-l);
  margin-bottom: 12px;
}

/* ライブカウンター */
.lctr {
  background: var(--bg2); border: 1px solid var(--br);
  border-radius: var(--r-sm); padding: 9px 14px;
  display: flex; align-items: center; gap: 14px; flex-wrap: wrap;
  margin-bottom: 12px;
  position: sticky; top: 0; z-index: 200;
  backdrop-filter: blur(8px);
}
.lc-lbl  { font-size: 10px; color: var(--t2); white-space: nowrap; }
.lc-tot  { font-size: 21px; font-weight: 700; color: #FFF; line-height: 1; }
.lc-cat  { display: flex; align-items: center; gap: 4px; font-size: 12px; }
.lc-dot  { width: 6px; height: 6px; border-radius: 50%; flex-shrink: 0; }
.lc-warn { font-size: 10px; color: var(--amb-l); margin-left: auto; }

/* ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
   EDIT 画面 — 1件ずつ詳細編集
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ */

/* 進行状況バー */
.edit-progress {
  background: var(--bg2); border-radius: var(--r-sm);
  padding: 8px 12px; margin-bottom: 14px;
  display: flex; align-items: center; gap: 12px;
}
.edit-prog-bar {
  flex: 1; height: 4px; background: rgba(255,255,255,.08);
  border-radius: 2px; overflow: hidden;
}
.edit-prog-fill {
  height: 100%; border-radius: 2px;
  background: linear-gradient(90deg, var(--red), var(--red-l));
  transition: width .3s ease;
}
.edit-prog-txt { font-size: 11px; color: var(--t2); white-space: nowrap; }

/* 編集カード本体 */
.edit-card {
  background: var(--card); border: 1px solid var(--b1);
  border-radius: var(--r-lg); padding: 20px 22px;
  margin-bottom: 12px;
}
.edit-card-top {
  display: flex; align-items: center; gap: 8px;
  margin-bottom: 14px;
}
.edit-cat-badge {
  font-size: 10px; font-weight: 700; padding: 3px 9px;
  border-radius: 20px; letter-spacing: .04em;
}
.edit-cat-badge.issue  { background: var(--red-a);  color: var(--red-l);  border: 1px solid var(--br); }
.edit-cat-badge.action { background: var(--amb-a);  color: var(--amb-l);  border: 1px solid var(--ba); }
.edit-cat-badge.result { background: var(--grn-a);  color: var(--grn-l);  border: 1px solid var(--bg_); }
.edit-src {
  font-size: 10px; color: var(--t3); margin-left: auto;
}

/* 元テキストボックス */
.orig-box {
  background: rgba(255,255,255,.03);
  border: 1px dashed rgba(200,220,255,.14);
  border-radius: var(--r-sm);
  padding: 9px 12px; margin-bottom: 12px;
  font-size: 12px; color: var(--t2);
  line-height: 1.65; word-break: break-all;
}
.orig-box-lbl {
  font-size: 9px; font-weight: 700; letter-spacing: .1em;
  text-transform: uppercase; color: var(--t3); margin-bottom: 5px;
}

/* AI補助パネル */
.ai-panel {
  background: var(--blu-a); border: 1px solid rgba(30,58,138,.22);
  border-left: 3px solid var(--blu-l);
  border-radius: var(--r-sm); padding: 9px 12px;
  margin-bottom: 12px;
}
.ai-panel-ttl { font-size: 10px; font-weight: 700; color: var(--blu-l); margin-bottom: 6px; }
.ai-row { font-size: 11px; color: var(--t2); padding: 2px 0; line-height: 1.6; }
.ai-row strong { color: var(--t1); }

/* 短文編集入力 */
.edit-input-lbl {
  font-size: 10px; font-weight: 700; letter-spacing: .08em;
  color: var(--t3); text-transform: uppercase; margin-bottom: 4px;
}

/* ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
   CONFIRM 画面 — プレビュー + スコア + DL
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ */

/* スライドプレビュー */
.sp { background: #F5F7FC; border: 1px solid #DDE3F0; border-radius: var(--r-md); overflow: hidden; box-shadow: var(--shadow); }
.sp-top  { background: #192848; padding: 9px 14px; display: flex; align-items: center; justify-content: space-between; }
.sp-ttl  { font-size: 11px; font-weight: 700; color: #FFF; }
.sp-date { font-size: 10px; color: rgba(255,255,255,.38); }
.sp-body { padding: 11px 13px; }
.sp-sec  { margin-bottom: 9px; }
.sp-sn   { font-size: 8px; font-weight: 700; letter-spacing: .14em; text-transform: uppercase; margin-bottom: 4px; padding-bottom: 2px; border-bottom: 1px solid; }
.sp-sn.i { color: #B8001E; border-color: rgba(184,0,30,.25); }
.sp-sn.a { color: #92400E; border-color: rgba(146,64,14,.25); }
.sp-sn.r { color: #166534; border-color: rgba(22,101,52,.25); }
.sp-it   { font-size: 11.5px; color: #1A1A2A; padding: 2px 0 2px 8px; line-height: 1.5; }
.sp-em   { font-size: 11px; color: #B0B4C0; font-style: italic; padding-left: 8px; }
.sp-ft   { padding: 6px 13px; border-top: 1px solid #E4E8F2; background: #EDF0F8; font-size: 10px; color: #888; display: flex; justify-content: space-between; }

/* 品質スコア */
.score-wrap   { background: var(--bg2); border: 1px solid var(--b1); border-radius: var(--r-md); padding: 14px 16px; }
.score-ttl    { font-size: 10px; font-weight: 700; color: var(--t2); letter-spacing: .08em; text-transform: uppercase; margin-bottom: 12px; }
.score-axis   { display: flex; align-items: center; gap: 10px; margin-bottom: 7px; }
.score-lbl    { font-size: 11px; color: var(--t2); width: 84px; flex-shrink: 0; }
.score-bg     { flex: 1; background: rgba(255,255,255,.06); border-radius: 3px; height: 5px; overflow: hidden; }
.score-fill   { height: 100%; border-radius: 3px; transition: width .5s ease; }
.sf-r { background: linear-gradient(90deg, var(--red), var(--red-l)); }
.sf-a { background: linear-gradient(90deg, var(--amb), var(--amb-l)); }
.sf-g { background: linear-gradient(90deg, var(--grn), var(--grn-l)); }
.score-pct  { font-size: 11px; font-weight: 700; width: 30px; text-align: right; flex-shrink: 0; }
.sp-r { color: var(--red-l); }
.sp-a { color: var(--amb-l); }
.sp-g { color: var(--grn-l); }
.score-verdict { font-size: 12px; margin-top: 9px; padding: 7px 10px; border-radius: var(--r-sm); text-align: center; font-weight: 700; }
.sv-ok   { background: var(--grn-a); color: var(--grn-l); border: 1px solid var(--bg_); }
.sv-warn { background: var(--amb-a); color: var(--amb-l); border: 1px solid var(--ba);  }
.sv-ng   { background: var(--red-a); color: var(--red-l); border: 1px solid var(--br);  }

/* ダウンロードカード */
.dl-card {
  background: linear-gradient(140deg, #040A14, #0A1628);
  border: 2px solid var(--red); border-radius: var(--r-lg);
  padding: 24px 36px; text-align: center; margin: 14px 0;
  box-shadow: 0 0 40px rgba(200,0,46,.08);
}
.dl-ttl { font-family: 'Noto Serif JP', serif; font-size: 17px; font-weight: 600; color: #FFF; margin-bottom: 4px; }
.dl-sub { font-size: 12px; color: #CADCFC; line-height: 1.8; }
.pulse  { animation: pa 1.8s ease-in-out infinite; border-radius: var(--r-sm); display: block; }
@keyframes pa {
  0%, 100% { box-shadow: 0 0 8px rgba(200,0,46,.3); }
  50%       { box-shadow: 0 0 28px rgba(200,0,46,.75), 0 0 50px rgba(200,0,46,.3); }
}

/* ── フッター ── */
.footer { border-top: 1px solid var(--b0); padding: 8px 48px; text-align: center; font-size: 10px; color: var(--t3); letter-spacing: .08em; }

/* ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
   Streamlit ウィジェット上書き
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ */
.stButton > button {
  background: transparent !important; color: var(--red) !important;
  font-family: 'Noto Sans JP', sans-serif !important; font-size: 13px !important;
  font-weight: 500 !important; letter-spacing: .05em !important;
  border: 1px solid var(--red) !important; border-radius: var(--r-sm) !important;
  padding: 8px 20px !important; width: 100% !important; transition: all .16s !important;
}
.stButton > button:hover { background: rgba(200,0,46,.07) !important; box-shadow: 0 0 12px rgba(200,0,46,.16) !important; }
[data-testid="stDownloadButton"] > button {
  background: linear-gradient(135deg, var(--red), var(--red-l)) !important;
  color: #FFF !important; font-weight: 700 !important; font-size: 15px !important;
  border: none !important; border-radius: var(--r-sm) !important;
  padding: 13px 40px !important; width: 100% !important;
}
[data-testid="stDownloadButton"] > button:hover {
  transform: translateY(-2px) scale(1.008) !important;
  box-shadow: 0 10px 28px rgba(200,0,46,.45) !important;
}
.stProgress > div { background: rgba(200,0,46,.08) !important; height: 3px !important; border-radius: 2px !important; }
.stProgress > div > div { background: linear-gradient(90deg, var(--red), var(--red-l)) !important; }
[data-testid="stCheckbox"] label { color: var(--t1) !important; font-size: 13px !important; }
[data-testid="stCheckbox"] { padding: 0 !important; }
[data-testid="stTextInput"] input {
  background: rgba(255,255,255,.04) !important; border: 1px solid var(--b1) !important;
  color: var(--t1) !important; font-size: 13px !important; border-radius: var(--r-sm) !important;
}
[data-testid="stTextInput"] input:focus {
  border-color: rgba(200,0,46,.4) !important;
  box-shadow: 0 0 0 2px rgba(200,0,46,.09) !important;
}
[data-testid="stSelectbox"] > div > div {
  background: rgba(255,255,255,.04) !important; border: 1px solid var(--b1) !important;
  color: var(--t1) !important; border-radius: var(--r-sm) !important;
}
[data-baseweb="select"] li { background: var(--bg2) !important; color: var(--t1) !important; }
.stAlert { background: var(--bg2) !important; border-radius: var(--r-sm) !important; }
.stRadio > div { flex-direction: row !important; flex-wrap: wrap !important; gap: 8px !important; }
.stRadio label { color: var(--t1) !important; font-size: 12px !important; }
[data-testid="stTextArea"] textarea {
  background: rgba(255,255,255,.04) !important; border: 1px solid var(--b1) !important;
  color: var(--t1) !important; font-size: 13px !important; border-radius: var(--r-sm) !important;
}
</style>
"""


# ==============================================================================
# LAYER 8 ─ UI コンポーネント (v14 完全再設計)
#
# v14 設計思想:
# 「画面遷移」で認知コストを解消する。
# session_state["ui_phase"] が "upload" / "review" / "edit" / "confirm" の
# 4値を取り、それぞれの画面に完全に分離する。
#
# 画面遷移フロー:
#   upload ──[解析ボタン]──► review ──[詳細を編集]──► edit ──[確認画面へ]──►
#                                    ──[確認画面へ]──────────────────────────►
#   confirm ◄──────────────────────────────────────────────────────────────
#   confirm ──[修正に戻る]──► review
#
# 各画面の目的:
#   upload  : ファイルをドロップするだけ。1アクション。
#   review  : 3列一覧で採用/除外を決める。短文のみ表示。詳細情報ゼロ。
#   edit    : 1件ずつ詳細確認・編集。元テキスト・AI補助表示。
#   confirm : プレビュー＋スコア＋ダウンロード。決定フェーズ。
# ==============================================================================


# ==============================================================================
# LAYER 4 ─ ドメインモデル (データクラス)
# ==============================================================================

@dataclass
class CatDef:
    """カテゴリの定義。UIの表示・分類ロジック・スライド生成すべてに使用する。"""
    name:      str
    icon:      str
    css:       str          # col-hdr に適用する CSS クラス
    question:  str          # カラムヘッダーに表示する問い
    examples:  str
    keywords:  list[str]
    weight:    int          # キーワードマッチスコアの重み
    color_key: str          # PPTX 生成で使う PC 色キー

CAT_DEFS: list[CatDef] = [
    CatDef(
        name="課題", icon="⚠️", css="issue", question="何が問題？", examples="遅延・クレーム・未達",
        weight=2, color_key="red",
        keywords=["課題","問題","障害","遅延","遅れ","バグ","エラー","リスク","懸念","未達","不足",
                  "困難","ペンディング","ボトルネック","停滞","失注","クレーム","不具合","超過",
                  "トラブル","対応中","調査中","原因","要因","影響","未解決","未対応","滞留",
                  "低下","悪化","急増","過多","オーバー","遅延","欠如","欠陥"],
    ),
    CatDef(
        name="対応", icon="🛠", css="action", question="何をやった？", examples="施策・改善・変更",
        weight=2, color_key="amb",
        keywords=["対応","対策","施策","改善","実施","導入","変更","修正","強化","推進","検討",
                  "整備","構築","展開","移行","再設計","見直し","追加","採用","開始","完了",
                  "協議","調整","確認","共有","報告","連携","作成","配布","展開","説明","周知",
                  "依頼","提案","承認","決定","実行"],
    ),
    CatDef(
        name="結果", icon="📊", css="result", question="どうなった？（数値で）", examples="120%・30%削減・0件",
        weight=3, color_key="grn",
        keywords=["売上","収益","利益","コスト","費用","予算","KPI","達成率","前月比","前年比",
                  "%","万円","億円","千件","PV","CVR","ROI","件数","稼働率","削減","短縮",
                  "増加","達成","成果","効果","改善幅","クリック率","回収","黒字","흑자"],
    ),
]

CATS: list[str]           = [c.name for c in CAT_DEFS]
CAT:  dict[str, CatDef]   = {c.name: c for c in CAT_DEFS}
HIGH_KW: list[str]        = [
    "売上","遅延","未達","障害","クレーム","失注","緊急","重大","停止",
    "不具合","赤字","損失","急増","炎上","訴訟","撤退","キャンセル",
]

# PPTX カラーパレット（PPTX_OK 時のみ使用）
if PPTX_OK:
    PC: dict[str, RGBColor] = {
        "navy":    RGBColor(0x06, 0x0E, 0x1C),
        "navy_lt": RGBColor(0x19, 0x28, 0x48),
        "white":   RGBColor(0xFF, 0xFF, 0xFF),
        "body":    RGBColor(0x1A, 0x1A, 0x2A),
        "muted":   RGBColor(0x88, 0x88, 0x99),
        "light":   RGBColor(0xF4, 0xF6, 0xFB),
        "red":     RGBColor(0xC8, 0x00, 0x2E),
        "grn":     RGBColor(0x16, 0x65, 0x34),
        "amb":     RGBColor(0x92, 0x40, 0x0E),
        "blu":     RGBColor(0x1E, 0x3A, 0x8A),
    }

# ==============================================================================
# LAYER 5 ─ テキスト処理エンジン
#
# ノイズ除去の設計:
# ① 文字種ノイズ: 記号・罫線文字のみの行を除外
# ② 長さノイズ:  5文字以下、または400文字超は除外
# ③ メタデータノイズ: 日付・ページ番号・ヘッダーパターンを除外
# ④ 構造ノイズ: 箇条書き記号のみの行を除外
# ⑤ 数値保護: 短縮時に数値含む表現を切り落とさない
# ==============================================================================

# 記号のみ行を判定する正規表現
_NOISE_CHARS = re.compile(
    r"^[\s\u3000\-=_■□◆◇▲▼●○★☆①-⑩〇\|/\\～〜＝─━…・。、　]+$"
)

# 除去する語尾パターン（短縮時）
_NOISE_WORDS = [
    r"という(こと|状況|状態|ため)",
    r"て(いただ|おり|いる|き)ます?",
    r"させていただ[きく]",
    r"ということで",
    r"に(関し|つき|ついて)(ましては?|は|も)?",
    r"なお[、 　,]",
    r"また[、 　,]",
    r"ご(報告|連絡|案内)(いたし)?ます",
    r"(お知らせ|ご確認)します",
    r"以上(となります|です)?",
    r"よろしく(お願い)?いたします",
]

# 数値パターン（保護対象）
_NUM_PAT: list[str] = [
    r'\d+[%％]',
    r'\d+\.?\d*\s*[万億千百]?円',
    r'\d+\s*件',
    r'\d+\s*[時間分秒]',
    r'前(月|年|期)比\s*\d+',
    r'[A-Z]{2,}\s*\d+',
    r'\d+\s*[倍割]',
    r'\d+\s*[人名台台個]',
    r'(?:▲|△|\+|-)\s*\d+',
]

# メタデータ行パターン（ノイズ除去対象）
_META_PAT = re.compile(
    r'^(第?\d+[ページ頁回期章節]|[Pp]\.?\s*\d+|slide\s*\d+|【.{1,8}】|《.{1,8}》|〔.{1,8}〕|\d{4}年\d{1,2}月.{0,4}$)',
    re.IGNORECASE
)

def has_num(text: str) -> bool:
    """数値パターンが含まれるか判定する。"""
    return any(re.search(p, text) for p in _NUM_PAT) or bool(re.search(r'\d+', text))

def is_high_priority(text: str) -> bool:
    """重要キーワードが含まれるか判定する。"""
    return any(kw in text for kw in HIGH_KW)

def is_noise(text: str) -> bool:
    """
    ノイズ行かどうかを判定する。
    v11 強化: メタデータパターンを追加、長文ノイズも除去。
    """
    t = text.strip()
    if len(t) <= 5:             return True
    if len(t) > 400:            return True  # v11 追加: 超長文は文章ではなくデータ行
    if _NOISE_CHARS.match(t):   return True
    if _META_PAT.match(t):      return True  # v11 追加
    # 日付のみ行
    if re.match(r'^\d{1,4}[年/\-]\d{1,2}[月/\-]\d{1,2}[日]?\s*$', t): return True
    # 見出し行（記号+短テキスト）
    if re.match(r'^[■□▲●◆★☆①-⑩【〔《]', t) and len(t) <= 15:         return True
    return False

def shorten(raw: str, max_chars: int = 48) -> str:
    """
    テキストを報告書向けの短文に変換する。
    数値を保護しながらノイズ語尾を除去し、自然な位置で切断する。
    """
    t = raw.strip()
    # 数値を保護リストに保存
    saved_nums = [m for p in _NUM_PAT for m in re.findall(p, t)]
    # ノイズ語尾を除去
    for pat in _NOISE_WORDS:
        t = re.sub(pat, "", t)
    t = re.sub(r"[\s\u3000]+", " ", t).strip()
    # 保護した数値が消えていたら復元
    for term in saved_nums:
        if term not in t:
            idx = raw.find(term)
            if idx >= 0:
                ctx = raw[max(0, idx - 6): idx + len(term) + 18].strip()
                t = f"{ctx} {t}".strip()
                break
    if len(t) <= max_chars:
        return t
    # 自然な区切り位置で切断
    cut = t[:max_chars]
    for sep in ["。", "、", "）", "】", "』", "」"]:
        idx = cut.rfind(sep)
        if idx > max_chars // 2:
            return cut[: idx + 1]
    return cut + "…"

def classify(text: str) -> str:
    """
    キーワードスコアリングによるカテゴリ分類。
    各カテゴリの weight × マッチ数 で最高スコアのカテゴリを返す。
    """
    best, best_score = CAT_DEFS[0].name, 0
    for cd in CAT_DEFS:
        score = sum(cd.weight for kw in cd.keywords if kw in text)
        if score > best_score:
            best_score, best = score, cd.name
    return best

def item_key(item: dict, prefix: str) -> str:
    """アイテムの一意キーを生成する（source + original のハッシュ）。"""
    h = hashlib.md5(
        f"{item['source']}||{item['original']}".encode()
    ).hexdigest()[:10]
    return f"{prefix}_{h}"

def calc_confidence(item: dict) -> float:
    """
    信頼度スコア算出ロジック (0.0 〜 1.0)

    スコアの構成:
      [A] テキスト長が適切 (15〜100 文字)          → +0.30
      [B] テキスト長が許容範囲 (100〜180 文字)     → +0.15
      [C] キーワード 2 件以上マッチ                → +0.30
      [D] キーワード 1 件マッチ                    → +0.15
      [E] 数値含有 (結果カテゴリ)                  → +0.22
      [F] 数値含有 (他カテゴリ)                    → +0.08
      [G] 重要キーワード含有                       → +0.10
      [H] 短文化後にノイズでない                   → +0.10
      [I] v11 新規: 因果ボーナス                   → +0.15
          同ソースファイルの前後に対応カテゴリの
          キーワードが存在する場合に加点
          (因果の繋がりを持つ行は信頼度を上げる)

    最大合計: [A]+[C]+[E]+[G]+[H]+[I] = 0.30+0.30+0.22+0.10+0.10+0.15 = 1.17
    → min(score, 1.0) でクリップ

    しきい値 CONF_TH = 0.52:
      0.52 未満 → 要確認 (ユーザーが見るべき行)
      0.52 以上 → 自動採用候補 (折りたたみ表示)
    """
    orig = item.get("original", "")
    cat  = item.get("category", "課題")
    cd   = CAT.get(cat, CAT_DEFS[0])
    score = 0.0

    # [A][B] テキスト長
    l = len(orig)
    if 15 <= l <= 100:   score += 0.30
    elif 100 < l <= 180: score += 0.15

    # [C][D] キーワードマッチ
    mc = sum(1 for kw in cd.keywords if kw in orig)
    if mc >= 2:   score += 0.30
    elif mc == 1: score += 0.15

    # [E][F] 数値含有
    if has_num(orig):
        score += 0.22 if cat == "結果" else 0.08

    # [G] 重要キーワード
    if is_high_priority(orig): score += 0.10

    # [H] 短文ノイズチェック
    if not is_noise(item.get("short", "")): score += 0.10

    # [I] 因果ボーナス (v11 新規)
    if item.get("causal_bonus", False): score += 0.15

    return min(score, 1.0)

def get_flags(item: dict) -> list[tuple[str, str]]:
    """
    問題フラグを返す。最初にマッチした 1 件のみを返す優先順位制。
    「複数フラグが重なって混乱する」問題を解消するため、
    最も重要な1つだけを表示する。(v11 変更)
    """
    orig = item.get("original", "")
    cat  = item.get("category", "")
    conf = item.get("confidence", 0)

    if is_high_priority(orig):              return [("high",    "🔴 重要案件")]
    if cat == "結果" and not has_num(orig): return [("nonnum",  "💡 数値なし")]
    if conf < 0.35:                         return [("lowconf", "❓ 分類不確実")]
    if has_num(orig):                       return [("num",     "📊 数値あり")]
    return [("ok", "")]

def get_review_hint(item: dict) -> str:
    """
    確認理由を 1 行で返す。v11 変更: 最も緊急度の高い 1 件のみ。
    「何をすればいいか」を迷わせない。
    """
    cat  = item.get("category", "")
    conf = item.get("confidence", 0)
    orig = item.get("original", "")
    if cat == "結果" and not has_num(orig):
        return "💡 「結果」に数値がありません。数値を追加するかカテゴリを変えてください。"
    if conf < 0.35:
        return "❓ 分類キーワードが不足しています。カテゴリが正しいか確認してください。"
    if len(orig) > 160:
        return "✂️ 元テキストが長く短文化の精度が下がる可能性があります。短文を確認してください。"
    return ""

CONF_TH = 0.52
VOL = {"少（上位8件）": 8, "中（上位15件）": 15, "多（全件）": 999}

# ==============================================================================
# LAYER 6 ─ ファイル読み込みエンジン
# ==============================================================================

def _rd_pptx(fb: bytes, nm: str) -> list[dict]:
    items = []
    try:
        prs = Presentation(io.BytesIO(fb))
        for i, sl in enumerate(prs.slides, 1):
            for sh in sl.shapes:
                if not sh.has_text_frame:
                    continue
                for pa in sh.text_frame.paragraphs:
                    t = pa.text.strip()
                    if t and len(t) > 4:
                        items.append({"original": t, "source": f"{nm} スライド{i}"})
    except Exception as e:
        items.append({"original": f"読み込みエラー: {e}", "source": nm})
    return items

def _rd_xlsx(fb: bytes, nm: str) -> list[dict]:
    items = []
    try:
        wb = openpyxl.load_workbook(io.BytesIO(fb), data_only=True)
        for sn in wb.sheetnames:
            for row in wb[sn].iter_rows():
                cells = [str(c.value).strip() for c in row if c.value is not None]
                if cells:
                    t = " | ".join(cells)
                    if len(t) > 4:
                        items.append({"original": t, "source": f"{nm} {sn}"})
    except Exception as e:
        items.append({"original": f"読み込みエラー: {e}", "source": nm})
    return items

def _rd_pdf(fb: bytes, nm: str) -> list[dict]:
    items = []
    try:
        with pdfplumber.open(io.BytesIO(fb)) as pdf:
            for i, pg in enumerate(pdf.pages, 1):
                for line in (pg.extract_text() or "").split("\n"):
                    t = line.strip()
                    if t and len(t) > 4:
                        items.append({"original": t, "source": f"{nm} p.{i}"})
    except Exception as e:
        items.append({"original": f"読み込みエラー: {e}", "source": nm})
    return items

def _rd_txt(fb: bytes, nm: str) -> list[dict]:
    items = []
    for enc in ["utf-8", "shift-jis", "cp932", "utf-16", "latin-1"]:
        try:
            for line in fb.decode(enc).split("\n"):
                t = line.strip()
                if t and len(t) > 4:
                    items.append({"original": t, "source": nm})
            return items
        except (UnicodeDecodeError, LookupError):
            continue
    items.append({"original": "文字コードを特定できませんでした", "source": nm})
    return items

def _attach_causal_bonus(items: list[dict]) -> None:
    """
    因果ボーナスの付与 (v11 新規)

    同一ソース内で前後20行以内に「課題→対応」「対応→結果」「課題→結果」
    の順番でキーワードが出現している行に causal_bonus=True を付与する。
    このフラグを受け取った calc_confidence() が +0.15 を加点する。

    これにより「因果の繋がりを持つ行」の信頼度が上がり、
    因果の繋がりが薄い行は要確認に分類される。
    """
    by_source: dict[str, list[dict]] = {}
    for it in items:
        src = it.get("source", "")
        by_source.setdefault(src, []).append(it)

    for src_items in by_source.values():
        for i, it in enumerate(src_items):
            cat = it.get("category", "")
            window = src_items[max(0, i - 20): i + 20]
            window_cats = [w.get("category", "") for w in window]
            if cat == "課題" and ("対応" in window_cats or "結果" in window_cats):
                it["causal_bonus"] = True
            elif cat == "対応" and ("課題" in window_cats or "結果" in window_cats):
                it["causal_bonus"] = True
            elif cat == "結果" and ("課題" in window_cats or "対応" in window_cats):
                it["causal_bonus"] = True

# ==============================================================================
# LAYER 7 ─ PPTX 生成エンジン
#
# v11 変更点:
# ・_tb() に line_spacing パラメータ追加 (行間を統一)
# ・全スライドの余白を統一 (0.5 Inches)
# ・数値スライドのグリッドを 2×4 から 2×3 に変更 (情報密度の適正化)
# ==============================================================================

def _bg(sl, c):
    f = sl.background.fill; f.solid(); f.fore_color.rgb = c

def _logo(sl):
    tb = sl.shapes.add_textbox(Inches(8.55), Inches(0.18), Inches(1.0), Inches(0.42))
    r = tb.text_frame.paragraphs[0].add_run()
    r.text = "IIJ"; r.font.size = Pt(20); r.font.bold = True
    r.font.color.rgb = PC["body"]
    d = sl.shapes.add_shape(
        MSO_SHAPE.OVAL, Inches(9.32), Inches(0.44), Inches(0.10), Inches(0.10)
    )
    d.fill.solid(); d.fill.fore_color.rgb = PC["red"]; d.line.fill.background()

def _tb(sl, text, l, t, w, h, size,
        bold=False, italic=False, color=None, align=None, spacing=1.2):
    """テキストボックスを追加する。spacing は行間倍率。"""
    from pptx.util import Pt as _Pt, Emu
    if align is None:
        align = PP_ALIGN.LEFT
    tb = sl.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = True
    p  = tf.paragraphs[0]; p.alignment = align
    # 行間設定
    from pptx.oxml.ns import qn
    from lxml import etree as _et
    pPr = p._p.get_or_add_pPr()
    lnSpc = _et.SubElement(pPr, qn('a:lnSpc'))
    spcPct = _et.SubElement(lnSpc, qn('a:spcPct'))
    spcPct.set('val', str(int(spacing * 100000)))
    r = p.add_run(); r.text = text
    r.font.size = _Pt(size); r.font.bold = bold; r.font.italic = italic
    if color:
        r.font.color.rgb = color

def _rc(sl, l, t, w, h, fill, line=None, lw=None):
    s = sl.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(l), Inches(t), Inches(w), Inches(h)
    )
    s.fill.solid(); s.fill.fore_color.rgb = fill
    if line:
        s.line.color.rgb = line; s.line.width = Pt(lw or 1)
    else:
        s.line.fill.background()

def _sl_title(prs, today: str) -> None:
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(sl, PC["navy"]); _rc(sl, 0, 0, 0.20, 7.5, PC["red"])
    _tb(sl, "Project Relay", 0.52, 1.8, 9.0, 1.0, 42, bold=True, color=PC["white"])
    _tb(sl, "月次報告レポート", 0.52, 2.9, 9.0, 0.6, 19, color=PC["muted"])
    sep = sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.52), Inches(3.6), Inches(8.8), Inches(0.018))
    sep.fill.solid(); sep.fill.fore_color.rgb = RGBColor(0x2A, 0x3C, 0x60)
    sep.line.fill.background()
    _tb(sl,
        "向平 友治 様\n"
        "確認・承認済みの情報を「課題 → 対応 → 結果」の構造で整理しました。",
        0.52, 4.0, 8.8, 1.7, 14, color=PC["white"], spacing=1.6)
    _tb(sl, f"生成日時: {today}", 0.52, 6.85, 9.0, 0.38, 10, color=PC["muted"])

def _sl_summary(prs, sel: dict) -> None:
    sl = prs.slides.add_slide(prs.slide_layouts[6]); _bg(sl, PC["light"]); _logo(sl)
    _tb(sl, "月次報告 サマリー", 0.50, 0.20, 9.0, 0.55, 24, bold=True, color=PC["body"])
    _tb(sl, "— 課題 / 対応 / 結果 —", 0.50, 0.75, 9.0, 0.30, 11, italic=True, color=PC["muted"])
    sep = sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.50), Inches(1.10), Inches(9.0), Inches(0.022))
    sep.fill.solid(); sep.fill.fore_color.rgb = PC["red"]; sep.line.fill.background()
    CW = 2.80
    for i, (lbl, cat, col) in enumerate([
        ("⚠️ 課題（問題点）", "課題", PC["red"]),
        ("🛠 対応（施策）",   "対応", PC["amb"]),
        ("📊 結果（数値成果）","結果", PC["grn"]),
    ]):
        x = 0.50 + i * (CW + 0.15)
        _rc(sl, x, 1.33, CW, 5.90, PC["light"], line=col, lw=0.8)
        _rc(sl, x, 1.33, CW, 0.38, col)
        _tb(sl, lbl, x + 0.10, 1.37, CW - 0.20, 0.30, 10, bold=True, color=PC["white"])
        items = sel.get(cat, [])
        content = "\n".join(f"・{it['short']}" for it in items[:6]) or "（データなし）"
        _tb(sl, content, x + 0.10, 1.82, CW - 0.20, 5.30, 11, color=PC["body"], spacing=1.5)
    _tb(sl, f"生成: {datetime.now().strftime('%Y/%m/%d')}",
        0.50, 7.12, 9.0, 0.28, 9, color=PC["muted"], align=PP_ALIGN.RIGHT)

def _sl_flow(prs, sel: dict) -> None:
    sl = prs.slides.add_slide(prs.slide_layouts[6]); _bg(sl, PC["white"]); _logo(sl)
    _tb(sl, "課題 → 対応 → 結果", 0.50, 0.20, 9.0, 0.55, 22, bold=True, color=PC["body"])
    _tb(sl, "— 問題の構造と対応結果 —", 0.50, 0.75, 9.0, 0.30, 11, italic=True, color=PC["muted"])
    sep = sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.50), Inches(1.10), Inches(9.0), Inches(0.022))
    sep.fill.solid(); sep.fill.fore_color.rgb = PC["red"]; sep.line.fill.background()
    CW = 2.52
    for i, (lbl, cat, hc, bg) in enumerate([
        ("⚠️ 課題", "課題", PC["red"], RGBColor(0xFF, 0xF0, 0xF3)),
        ("🛠 対応", "対応", PC["amb"], RGBColor(0xFF, 0xF7, 0xE8)),
        ("📊 結果", "結果", PC["grn"], RGBColor(0xEA, 0xFB, 0xEF)),
    ]):
        x = 0.50 + i * (CW + 0.32)
        _rc(sl, x, 1.35, CW, 5.82, bg, line=hc, lw=0.8)
        _rc(sl, x, 1.35, CW, 0.42, hc)
        _tb(sl, lbl, x + 0.10, 1.39, CW - 0.20, 0.34, 11, bold=True, color=PC["white"])
        items  = sel.get(cat, [])
        lines  = [f"  {j + 1}. {it['short']}" for j, it in enumerate(items[:5])] or ["（データなし）"]
        _tb(sl, "\n".join(lines), x + 0.08, 1.88, CW - 0.16, 5.20, 11, color=PC["body"], spacing=1.55)
        if i < 2:
            arr = sl.shapes.add_shape(
                MSO_SHAPE.RIGHT_ARROW,
                Inches(x + CW + 0.05), Inches(3.88), Inches(0.24), Inches(0.24)
            )
            arr.fill.solid(); arr.fill.fore_color.rgb = RGBColor(0xAA, 0xAA, 0xBB)
            arr.line.fill.background()
    _tb(sl, f"生成: {datetime.now().strftime('%Y/%m/%d')}",
        0.50, 7.12, 9.0, 0.28, 9, color=PC["muted"], align=PP_ALIGN.RIGHT)

def _sl_numbers(prs, items: list) -> None:
    ni = [it for it in items if has_num(it.get("short", "")) or has_num(it.get("original", ""))]
    if not ni:
        return
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(sl, PC["navy"]); _rc(sl, 0, 0, 0.20, 7.5, PC["grn"])
    tb2 = sl.shapes.add_textbox(Inches(8.55), Inches(0.18), Inches(1.0), Inches(0.42))
    r2  = tb2.text_frame.paragraphs[0].add_run()
    r2.text = "IIJ"; r2.font.size = Pt(20); r2.font.bold = True; r2.font.color.rgb = PC["white"]
    _tb(sl, "📊 数値サマリー", 0.52, 0.20, 9.0, 0.55, 24, bold=True, color=PC["white"])
    _tb(sl, "— 今月の定量的な成果 —", 0.52, 0.75, 9.0, 0.30, 11, italic=True, color=PC["muted"])
    sep = sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.52), Inches(1.10), Inches(9.0), Inches(0.020))
    sep.fill.solid(); sep.fill.fore_color.rgb = PC["grn"]; sep.line.fill.background()
    CW, RH, MX = 4.08, 1.40, 6  # v11: MX を 8→6 に変更 (2×3 グリッド)
    for idx, it in enumerate(ni[:MX]):
        col, row = idx % 2, idx // 2
        x = 0.52 + col * (CW + 0.36)
        y = 1.38 + row * (RH + 0.12)
        bg = sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(CW), Inches(RH))
        bg.fill.solid(); bg.fill.fore_color.rgb = RGBColor(0x12, 0x1E, 0x3A)
        bg.line.fill.background()
        bc = PC["red"] if is_high_priority(it["short"]) else PC["grn"]
        _rc(sl, x, y, 0.05, RH, bc)
        _tb(sl, it["short"], x + 0.12, y + 0.17, CW - 0.22, RH - 0.32, 12,
            color=PC["white"], spacing=1.45)
        _tb(sl, it["source"], x + 0.12, y + RH - 0.28, CW - 0.22, 0.24, 8, color=PC["muted"])
    _tb(sl, f"生成: {datetime.now().strftime('%Y/%m/%d')}",
        0.50, 7.12, 9.0, 0.28, 9, color=PC["muted"], align=PP_ALIGN.RIGHT)

def _sl_detail(prs, cd: CatDef, items: list) -> None:
    if not items:
        return
    sl = prs.slides.add_slide(prs.slide_layouts[6]); _bg(sl, PC["white"]); _logo(sl)
    col = PC[cd.color_key]
    _tb(sl, f"{cd.icon}  {cd.name}の詳細", 0.50, 0.20, 9.0, 0.58, 22, bold=True, color=PC["body"])
    _tb(sl, cd.question, 0.50, 0.78, 9.0, 0.30, 11, italic=True, color=PC["muted"])
    sep = sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.50), Inches(1.12), Inches(9.0), Inches(0.022))
    sep.fill.solid(); sep.fill.fore_color.rgb = col; sep.line.fill.background()
    MX    = 8
    lines = []
    for it in items[:MX]:
        marker = "🔴" if is_high_priority(it["short"]) else "■"
        lines.append(f"{marker} {it['short']}")
        lines.append(f"      📁 {it['source']}")
    if len(items) > MX:
        lines.append(f"\n（他 {len(items) - MX} 件省略）")
    _tb(sl, "\n".join(lines), 0.50, 1.20, 9.0, 5.72, 10, color=PC["body"], spacing=1.55)
    _tb(sl, f"全 {len(items)} 件 | 生成: {datetime.now().strftime('%Y/%m/%d')}",
        0.50, 7.12, 9.0, 0.28, 9, color=PC["muted"], align=PP_ALIGN.RIGHT)

def build_pptx(sel: dict, today: str) -> bytes:
    """選択済みアイテムから PPTX を生成して bytes で返す。"""
    prs = Presentation()
    prs.slide_width  = Inches(10)
    prs.slide_height = Inches(7.5)
    all_items = [it for items in sel.values() for it in items]
    _sl_title(prs, today)
    _sl_summary(prs, sel)
    _sl_flow(prs, sel)
    _sl_numbers(prs, all_items)
    for cd in CAT_DEFS:
        _sl_detail(prs, cd, sel.get(cd.name, []))
    buf = io.BytesIO(); prs.save(buf)
    return buf.getvalue()



# UI フェーズの定数
PHASE_UPLOAD  = "upload"
PHASE_REVIEW  = "review"
PHASE_EDIT    = "edit"
PHASE_CONFIRM = "confirm"

# 「確認が必要」なアイテムの信頼度しきい値
# CONF_TH は LAYER 5 で定義済み（0.52）

VOL = {"少（上位8件）": 8, "中（上位15件）": 15, "多（全件）": 999}


# ------------------------------------------------------------------------------
# 共通コンポーネント
# ------------------------------------------------------------------------------

def render_header(phase: str) -> None:
    """
    全画面共通のヘッダーバー。
    ロゴ + ステップバー（現在地を常に明示）。
    """
    phase_to_step = {
        PHASE_UPLOAD:  1,
        PHASE_REVIEW:  2,
        PHASE_EDIT:    2,   # EDITはREVIEWのサブ画面なのでSTEP2
        PHASE_CONFIRM: 3,
    }
    step = phase_to_step.get(phase, 1)
    steps = [(1, "読み込み"), (2, "確認・編集"), (3, "確認・出力")]

    bar = '<div class="step-bar">'
    for i, (n, lbl) in enumerate(steps):
        if n < step:    css, ic = "done",   "✓"
        elif n == step: css, ic = "active",  str(n)
        else:           css, ic = "future",  str(n)
        bar += (
            f'<div class="si {css}">'
            f'<span class="sn">{ic}</span>'
            f'<span style="font-weight:700;margin-right:3px;">STEP {n}</span>'
            f'<span style="opacity:.7;font-size:10px;">{lbl}</span>'
            f'</div>'
        )
        if i < 2:
            bar += '<span class="sa">›</span>'
    bar += '</div>'

    st.markdown(
        f'<div class="hdr">'
        f'<div class="hdr-inner">'
        f'<span class="hdr-brand">⬡ Project Relay &nbsp;|&nbsp; 向平 友治 様専用</span>'
        f'{bar}'
        f'</div></div>',
        unsafe_allow_html=True,
    )


def render_page_title(sub: str, main: str, hint: str = "") -> None:
    """ページタイトルブロック。画面遷移ごとに「今やること」を1行で示す。"""
    h = f'<div class="page-ttl"><div class="page-ttl-sub">{sub}</div>'
    h += f'<div class="page-ttl-main">{main}</div>'
    if hint:
        h += f'<div class="page-ttl-hint">{hint}</div>'
    h += '</div>'
    st.markdown(h, unsafe_allow_html=True)


def render_footer() -> None:
    st.markdown(
        '<div class="footer">Project Relay v14 &nbsp;|&nbsp; 向平 友治 様専用</div>',
        unsafe_allow_html=True,
    )


# ------------------------------------------------------------------------------
# UPLOAD 画面
# ------------------------------------------------------------------------------

def extract(uploaded_files) -> list[dict]:
    """
    複数ファイルからアイテムを抽出し、スコアリング・ソートして返す。
    ノイズ除去 → 短縮 → 分類 → 因果ボーナス → 信頼度計算 の順。
    ※ この関数が依存する _rd_pptx / shorten / classify / has_num /
      is_noise / calc_confidence / _attach_causal_bonus / HIGH_KW は
      すべてこのファイルの下方（LAYER 5〜6）で定義されているが、
      Pythonは関数本体を「呼び出し時」に評価するため問題ない。
      ただし呼び出し元の render_upload_screen より前にこの def が
      必要なため、ここに配置する。
    """
    # READERS辞書を使わず都度呼び出す（定義順序問題を回避するため）
    SUPPORTED = {".pptx", ".xlsx", ".pdf", ".txt"}
    all_items: list[dict] = []

    for uf in uploaded_files:
        ext = Path(uf.name).suffix.lower()
        if ext not in SUPPORTED:
            continue
        fb = uf.read()
        if   ext == ".pptx": raw = _rd_pptx(fb, uf.name)
        elif ext == ".xlsx": raw = _rd_xlsx(fb, uf.name)
        elif ext == ".pdf":  raw = _rd_pdf(fb, uf.name)
        elif ext == ".txt":  raw = _rd_txt(fb, uf.name)
        else:                continue
        for it in raw:
            orig = it["original"]
            it["short"]    = shorten(orig)
            it["category"] = classify(orig)
            it["score"]    = (
                sum(10 for kw in HIGH_KW if kw in orig)
                + (5 if has_num(orig) else 0)
                + (2 if len(orig) < 80 else 0)
            )
            it["causal_bonus"] = False  # 後で _attach_causal_bonus が設定
        raw = [it for it in raw if not is_noise(it["short"])]
        all_items.extend(raw)

    # 因果ボーナスを付与してから信頼度を計算
    _attach_causal_bonus(all_items)
    for it in all_items:
        it["confidence"] = calc_confidence(it)

    return sorted(all_items, key=lambda x: x["score"], reverse=True)


def render_upload_screen() -> None:
    """
    【目的】ファイルをドロップして解析ボタンを押す。それだけ。
    【削除した情報】
    ・拡張子バッジ（抽象的すぎる）
    ・抽出量ラジオ（解析前に見せても意味不明）
    ・長い説明文（読まれない）
    → 追加したもの: 解析後に抽出量を確認できるようにした（review画面に移動）
    """
    render_header(PHASE_UPLOAD)
    render_page_title(
        "STEP 1  ファイルを読み込む",
        "月次報告ファイルを<em>ここにドロップ</em>してください",
        "会議メモ・日報・スプレッドシート・PDF — 複数ファイル同時にOK",
    )

    st.markdown('<div class="wn">', unsafe_allow_html=True)

    uploaded = st.file_uploader(
        "月次報告ファイルをドロップ",
        type=["pptx", "xlsx", "pdf", "txt"],
        accept_multiple_files=True,
        label_visibility="collapsed",
    )

    # ファイルが変わったらセッション初期化
    new_names = sorted(f.name for f in uploaded) if uploaded else []
    if st.session_state.get("_fn") != new_names:
        st.session_state["_fn"] = new_names
        for k in ["raw", "ui_phase", "edit_idx", "edit_cat", "edit_cat_filter"]:
            st.session_state.pop(k, None)
        for k in list(st.session_state.keys()):
            if k.startswith(("chk_", "sht_", "cat_")):
                del st.session_state[k]

    if not uploaded:
        st.markdown(
            '<div class="up-hint">'
            '<strong>💡 読み込めるファイル</strong>'
            'Excel・PowerPoint・PDF・テキストファイル（会議メモ・日報など）'
            '<div class="up-ok">✅ 箇条書きでOK &nbsp;・&nbsp; ✅ 数値がなくてもOK &nbsp;・&nbsp; ✅ 複数ファイルまとめてOK</div>'
            '</div>',
            unsafe_allow_html=True,
        )
        st.markdown('</div>', unsafe_allow_html=True)
        render_footer()
        return

    # ファイルがある → ラジオ + 解析ボタン
    st.markdown('<hr>', unsafe_allow_html=True)

    fnames = "・".join(f.name for f in uploaded[:2])
    if len(uploaded) > 2:
        fnames += f"…他{len(uploaded)-2}件"

    cv, ch = st.columns([3, 2])
    with cv:
        vol_choice = st.radio(
            "読み込む件数", list(VOL.keys()), index=0,
            horizontal=True, label_visibility="visible",
        )
    with ch:
        st.markdown(
            '<div style="padding-top:26px;font-size:11px;color:#7A8BAA;">'
            '初めての場合は「少」がおすすめ</div>',
            unsafe_allow_html=True,
        )

    st.markdown('<hr>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        if st.button(f"⬡　{fnames} を解析する", use_container_width=True):
            prog = st.progress(0)
            with st.spinner("解析中…"):
                items = extract(uploaded)
            prog.progress(100); prog.empty()
            st.session_state["raw"]       = items
            st.session_state["vol_choice"] = vol_choice
            st.session_state["ui_phase"]  = PHASE_REVIEW
            st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)
    render_footer()


# ------------------------------------------------------------------------------
# REVIEW 画面 — 3列一覧確認
# ------------------------------------------------------------------------------

def _get_review_flag(item: dict) -> tuple[str, str]:
    """
    一覧表示用のフラグ（1件のみ、最優先のものを返す）。
    returns: (css_class, label_text)
    """
    orig = item.get("original", "")
    cat  = item.get("category", "")
    conf = item.get("confidence", 0)
    if is_high_priority(orig):              return ("r", "🔴 重要")
    if cat == "結果" and not has_num(orig): return ("a", "💡 数値なし")
    if conf < CONF_TH:                      return ("b", "❓ 要確認")
    return ("g", "")


def render_review_screen() -> None:
    """
    【目的】採用/除外を決める。それだけ。
    【表示するもの】
    ・短文テキスト（スライドに出る文）
    ・フラグ（重要/数値なし/要確認）のみ
    【非表示にしたもの】
    ・元テキスト（→ EDIT画面へ）
    ・信頼度スコア数値
    ・カテゴリ変更UI
    ・編集フォーム
    ・ガイドボックス（複数のルールを一度に説明しない）
    """
    render_header(PHASE_REVIEW)

    raw: list[dict] = st.session_state.get("raw", [])
    vol_limit = VOL.get(
        st.session_state.get("vol_choice", "少（上位8件）"), 8
    )

    by_cat: dict[str, list] = {c: [] for c in CATS}
    for it in raw:
        c = it.get("category", "課題")
        if c in by_cat:
            by_cat[c].append(it)

    # 各カテゴリをvol_limitで絞る
    by_cat_limited = {c: items[:vol_limit] for c, items in by_cat.items()}

    # 警告計算（ページタイトルに反映）
    review_needed = sum(
        1 for items in by_cat_limited.values()
        for it in items if it.get("confidence", 0) < CONF_TH
    )
    total_items = sum(len(v) for v in by_cat_limited.values())

    if review_needed > 0:
        hint = f"黄色の行が<strong>{review_needed}件</strong>あります。詳細編集で確認してください"
    else:
        hint = f"全{total_items}件が自動採用候補です。そのまま確認画面へ進めます"

    render_page_title(
        "STEP 2  採用する項目を確認する",
        "チェックを<em>外すと除外</em>されます",
        "",
    )
    # ページタイトル下に件数ヒント
    st.markdown(
        f'<div class="wn" style="padding-bottom:0;">'
        f'<div class="warn-banner">{hint}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

    # ── sticky カウンター ──
    dots = {"課題": "#C8002E", "対応": "#92400E", "結果": "#166534"}
    def _render_counter(by_cat_lim: dict) -> None:
        cat_parts = "".join(
            f'<span class="lc-cat">'
            f'<span class="lc-dot" style="background:{dots[c]}"></span>'
            f'{CAT[c].icon}{c}:<strong id="cnt_{c}">?</strong>'
            f'</span>'
            for c in CATS
        )
        # カウントはセッションから算出
        cat_counts = {
            c: sum(1 for it in items if st.session_state.get(item_key(it, f"chk_{c}"), True))
            for c, items in by_cat_lim.items()
        }
        total = sum(cat_counts.values())
        cat_html = "".join(
            f'<span class="lc-cat">'
            f'<span class="lc-dot" style="background:{dots[c]}"></span>'
            f'{CAT[c].icon}{c}:<strong>{cat_counts[c]}</strong>'
            f'</span>'
            for c in CATS
        )
        warn = (
            '<span class="lc-warn">⚠️ 結果が0件</span>'
            if cat_counts.get("課題", 0) > 0 and cat_counts.get("結果", 0) == 0
            else ""
        )
        st.markdown(
            f'<div class="wn" style="padding-top:4px;padding-bottom:4px;">'
            f'<div class="lctr">'
            f'<span class="lc-lbl">採用合計</span>'
            f'<span class="lc-tot">{total}</span>'
            f'<span class="lc-lbl">件</span>'
            f'{cat_html}{warn}'
            f'</div></div>',
            unsafe_allow_html=True,
        )

    _render_counter(by_cat_limited)

    # ── 3列一覧 ──
    cat_css = {"課題": "issue", "対応": "action", "結果": "result"}
    st.markdown('<div class="ww">', unsafe_allow_html=True)
    col_i, col_a, col_r = st.columns([1, 1, 1])
    col_map = {"課題": col_i, "対応": col_a, "結果": col_r}

    for cat, col in col_map.items():
        items_cat = by_cat_limited[cat]
        cd = CAT[cat]

        # 採用件数（セッションから）
        sel_cnt = sum(
            1 for it in items_cat
            if st.session_state.get(item_key(it, f"chk_{cat}"), True)
        )

        with col:
            # カテゴリヘッダー
            st.markdown(
                f'<div class="cat-hdr {cat_css[cat]}">'
                f'<div class="cat-hdr-left">'
                f'<span class="cat-hdr-icon">{cd.icon}</span>'
                f'<div>'
                f'<div class="cat-hdr-name">{cat}</div>'
                f'<div class="cat-hdr-q">{cd.question}</div>'
                f'</div></div>'
                f'<span class="cat-hdr-cnt">{sel_cnt} / {len(items_cat)}</span>'
                f'</div>',
                unsafe_allow_html=True,
            )

            if not items_cat:
                st.markdown(
                    '<div style="padding:16px;text-align:center;font-size:12px;'
                    'color:#3D4E68;background:var(--card);border:1px solid var(--b0);'
                    'border-radius:0 0 var(--r-sm) var(--r-sm);">項目なし</div>',
                    unsafe_allow_html=True,
                )
                continue

            # 各アイテムを行で表示
            for it in items_cat:
                ck = item_key(it, f"chk_{cat}")
                if ck not in st.session_state:
                    st.session_state[ck] = True

                is_checked = st.session_state[ck]
                flag_cls_name, flag_lbl = _get_review_flag(it)

                # アイテム行のCSSクラスを決定
                if not is_checked:
                    row_cls = "rv-item excl"
                elif flag_cls_name == "r":
                    row_cls = "rv-item high-priority"
                elif flag_cls_name in ("a", "b"):
                    row_cls = "rv-item needs-review"
                else:
                    row_cls = "rv-item auto-ok"

                # フラグHTML
                flag_html = (
                    f'<span class="rv-flag {flag_cls_name}">{flag_lbl}</span>'
                    if flag_lbl else ""
                )

                # 行コンテナ（チェックボックスと短文を横並び）
                col_chk, col_content = st.columns([0.4, 5.6])
                with col_chk:
                    st.checkbox("", key=ck, label_visibility="collapsed")
                with col_content:
                    short_txt = st.session_state.get(
                        item_key(it, f"sht_{cat}"), it["short"]
                    )
                    st.markdown(
                        f'<div class="{row_cls}">'
                        f'<span class="rv-short">{short_txt}</span>'
                        f'{flag_html}'
                        f'</div>',
                        unsafe_allow_html=True,
                    )

    st.markdown('</div>', unsafe_allow_html=True)  # ww

    # ── ボタン行 ──
    st.markdown('<div class="wn" style="padding-top:16px;">', unsafe_allow_html=True)
    st.markdown('<hr>', unsafe_allow_html=True)

    # 要確認アイテムがある場合に「詳細編集」ボタンを目立たせる
    all_items_flat = [
        (cat, it)
        for cat, items in by_cat_limited.items()
        for it in items
    ]
    review_items = [
        (i, cat, it)
        for i, (cat, it) in enumerate(all_items_flat)
        if it.get("confidence", 0) < CONF_TH or is_high_priority(it.get("original", ""))
    ]

    if review_items:
        c1, c2, c3 = st.columns([1, 2, 1])
        with c2:
            if st.button(
                f"✎　要確認 {len(review_items)} 件を詳細編集する",
                use_container_width=True,
            ):
                # 最初の要確認アイテムにジャンプ
                first_idx, first_cat, _ = review_items[0]
                st.session_state["ui_phase"]       = PHASE_EDIT
                st.session_state["edit_idx"]       = first_idx
                st.session_state["edit_items_flat"] = all_items_flat
                # フィルター: 要確認のみ or 全件
                st.session_state["edit_filter"]    = "要確認のみ"
                st.session_state["edit_review_idxs"] = [i for i, _, _ in review_items]
                st.rerun()

    # 全件編集ボタン（サブ）
    ca, cb = st.columns([1, 1])
    with ca:
        if st.button("📋　全件を詳細確認する", use_container_width=True):
            all_items_flat_recomputed = [
                (cat, it)
                for cat, items in by_cat_limited.items()
                for it in items
            ]
            st.session_state["ui_phase"]        = PHASE_EDIT
            st.session_state["edit_idx"]        = 0
            st.session_state["edit_items_flat"] = all_items_flat_recomputed
            st.session_state["edit_filter"]     = "全件"
            st.session_state["edit_review_idxs"] = list(range(len(all_items_flat_recomputed)))
            st.rerun()

    with cb:
        # 確認画面へ進む（編集スキップ）
        total_sel = sum(
            1 for cat, items in by_cat_limited.items()
            for it in items
            if st.session_state.get(item_key(it, f"chk_{cat}"), True)
        )
        if total_sel == 0:
            st.warning("⚠️ 1件以上チェックを入れてください")
        else:
            if st.button(
                f"▶　このまま確認画面へ（{total_sel}件）",
                use_container_width=True,
            ):
                # selected を計算してセッションに保存
                selected = _build_selected_from_review(by_cat_limited)
                st.session_state["selected"] = selected
                st.session_state["ui_phase"] = PHASE_CONFIRM
                st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)
    render_footer()


def _build_selected_from_review(by_cat_limited: dict) -> dict:
    """
    一覧確認画面のチェックボックス状態から selected dict を構築する。
    カテゴリ変更があった場合も正しいカテゴリへ振り替える。
    """
    selected: dict[str, list]    = {c: [] for c in CATS}
    moved_cross: dict[str, list] = {c: [] for c in CATS}

    for cat, items in by_cat_limited.items():
        for it in items:
            ck = item_key(it, f"chk_{cat}")
            if not st.session_state.get(ck, True):
                continue  # 除外
            sk     = item_key(it, f"sht_{cat}")
            gk     = item_key(it, f"cat_{cat}")
            short  = st.session_state.get(sk, it["short"])
            to_cat = st.session_state.get(gk, cat)
            r = {**it, "short": short, "category": to_cat}
            if to_cat == cat:
                selected[cat].append(r)
            else:
                moved_cross.setdefault(to_cat, []).append(r)

    for to_cat, ml in moved_cross.items():
        if to_cat in selected:
            selected[to_cat].extend(ml)

    return selected


# ------------------------------------------------------------------------------
# EDIT 画面 — 1件ずつ詳細編集
# ------------------------------------------------------------------------------

def _ai_assist(item: dict) -> dict:
    """
    AI補助情報を生成する（ロジックベース、外部API不使用）。

    returns:
      {
        "reason":    str  — なぜこのカテゴリに分類されたか
        "hint":      str  — 改善提案（数値不足・長文など）
        "keywords":  list — マッチしたカテゴリキーワード
      }
    """
    orig = item.get("original", "")
    cat  = item.get("category", "課題")
    conf = item.get("confidence", 0)
    cd   = CAT.get(cat, CAT_DEFS[0])

    matched_kw = [kw for kw in cd.keywords if kw in orig]
    has_causal  = item.get("causal_bonus", False)

    # 分類理由
    if matched_kw:
        reason = f"「{'」「'.join(matched_kw[:3])}」が含まれているため「{cat}」に分類"
    elif has_causal:
        reason = f"前後の文脈から「{cat}」との関連が検出されました"
    else:
        reason = f"キーワードのマッチが弱いため分類が不確かです（信頼度 {int(conf*100)}%）"

    # 改善提案
    hints = []
    if cat == "結果" and not has_num(orig):
        hints.append("💡 「結果」には数値（%・円・件数）があると説得力が上がります")
    if conf < 0.35:
        hints.append(f"❓ このカテゴリで正しいか確認してください（信頼度 {int(conf*100)}%）")
    if len(orig) > 160:
        hints.append("✂️ 元テキストが長いため短文化の精度が低下している可能性があります")
    if is_high_priority(orig):
        hints.append("🔴 重要キーワードが含まれています。スライドで目立つ位置に配置されます")

    return {
        "reason":   reason,
        "hint":     " / ".join(hints) if hints else "✅ 問題なし",
        "keywords": matched_kw[:5],
    }


def render_edit_screen() -> None:
    """
    【目的】1件ずつ元テキストを確認し、短文と分類を修正する。
    【表示するもの】
    ・進行状況バー（何件中何件目か）
    ・元テキスト（全文）
    ・AI補助パネル（分類理由・改善提案・キーワード）
    ・短文編集テキストボックス
    ・カテゴリ変更セレクト
    ・前へ/次へ ナビゲーション
    【非表示にしたもの】
    ・他のアイテム（1件に集中させる）
    ・スライドプレビュー
    ・品質スコア
    """
    render_header(PHASE_EDIT)

    items_flat: list[tuple[str, dict]] = st.session_state.get("edit_items_flat", [])
    review_idxs: list[int] = st.session_state.get("edit_review_idxs", list(range(len(items_flat))))
    edit_filter: str = st.session_state.get("edit_filter", "全件")

    # 現在インデックス（review_idxs の中での位置）
    # edit_idx = items_flat の絶対インデックス
    edit_idx = st.session_state.get("edit_idx", review_idxs[0] if review_idxs else 0)

    # edit_idx が review_idxs にない場合は先頭に修正
    if edit_idx not in review_idxs:
        edit_idx = review_idxs[0] if review_idxs else 0
        st.session_state["edit_idx"] = edit_idx

    pos_in_queue = review_idxs.index(edit_idx) if edit_idx in review_idxs else 0
    total_in_queue = len(review_idxs)

    cat, item = items_flat[edit_idx]
    cd = CAT[cat]
    cat_css_map = {"課題": "issue", "対応": "action", "結果": "result"}

    # セッションキー
    ck = item_key(item, f"chk_{cat}")
    sk = item_key(item, f"sht_{cat}")
    gk = item_key(item, f"cat_{cat}")
    if ck not in st.session_state: st.session_state[ck] = True
    if sk not in st.session_state: st.session_state[sk] = item["short"]
    if gk not in st.session_state: st.session_state[gk] = cat

    current_cat = st.session_state[gk]
    ai = _ai_assist({**item, "category": current_cat})

    # ページタイトル
    render_page_title(
        f"STEP 2  詳細編集 {pos_in_queue + 1} / {total_in_queue}",
        "元テキストを確認して<em>短文を修正</em>してください",
        "カテゴリが間違っている場合は変更できます",
    )

    st.markdown('<div class="wn">', unsafe_allow_html=True)

    # ── 進行状況バー ──
    pct = int((pos_in_queue + 1) / total_in_queue * 100) if total_in_queue > 0 else 100
    st.markdown(
        f'<div class="edit-progress">'
        f'<div class="edit-prog-bar"><div class="edit-prog-fill" style="width:{pct}%;"></div></div>'
        f'<span class="edit-prog-txt">{pos_in_queue + 1} / {total_in_queue} 件</span>'
        f'</div>',
        unsafe_allow_html=True,
    )

    # ── 編集カード ──
    st.markdown(
        f'<div class="edit-card">'
        f'<div class="edit-card-top">'
        f'<span class="edit-cat-badge {cat_css_map.get(current_cat, "issue")}">'
        f'{CAT[current_cat].icon} {current_cat}</span>'
        f'<span class="edit-src">📁 {item["source"]}</span>'
        f'</div>',
        unsafe_allow_html=True,
    )

    # 元テキスト
    st.markdown(
        f'<div class="orig-box-lbl">元テキスト（ファイルからの抽出原文）</div>'
        f'<div class="orig-box">{item["original"]}</div>',
        unsafe_allow_html=True,
    )

    st.markdown('</div>', unsafe_allow_html=True)  # edit-card

    # ── AI補助パネル ──
    kw_html = (
        "　".join(f'<code style="font-size:10px;background:rgba(147,197,253,.1);'
                  f'color:var(--blu-l);padding:1px 4px;border-radius:2px;">{kw}</code>'
                  for kw in ai["keywords"])
        if ai["keywords"] else '<span style="color:var(--t3);font-size:10px;">なし</span>'
    )
    st.markdown(
        f'<div class="ai-panel">'
        f'<div class="ai-panel-ttl">🤖 AI分析</div>'
        f'<div class="ai-row"><strong>分類理由：</strong>{ai["reason"]}</div>'
        f'<div class="ai-row"><strong>改善提案：</strong>{ai["hint"]}</div>'
        f'<div class="ai-row"><strong>検出キーワード：</strong>{kw_html}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

    # ── 短文編集 ──
    st.markdown('<div class="edit-input-lbl">スライドに表示するテキスト（編集可）</div>', unsafe_allow_html=True)
    st.text_input(
        "短文", key=sk, label_visibility="collapsed",
        placeholder="スライドに表示するテキスト",
    )

    # ── カテゴリ変更 ──
    col_a, col_b = st.columns([1, 1])
    with col_a:
        st.markdown('<div class="edit-input-lbl" style="margin-top:8px;">カテゴリ</div>', unsafe_allow_html=True)
        st.selectbox(
            "カテゴリ", CATS,
            index=CATS.index(st.session_state[gk]),
            key=gk, label_visibility="collapsed",
        )
    with col_b:
        st.markdown('<div class="edit-input-lbl" style="margin-top:8px;">この項目を採用する</div>', unsafe_allow_html=True)
        st.checkbox("採用", key=ck, label_visibility="collapsed")

    # ── ナビゲーション ──
    st.markdown('<hr>', unsafe_allow_html=True)
    nav1, nav2, nav3 = st.columns([1, 1, 1])

    with nav1:
        if pos_in_queue > 0:
            if st.button("← 前の項目", use_container_width=True):
                st.session_state["edit_idx"] = review_idxs[pos_in_queue - 1]
                st.rerun()

    with nav2:
        if st.button("← 一覧に戻る", use_container_width=True):
            st.session_state["ui_phase"] = PHASE_REVIEW
            st.rerun()

    with nav3:
        if pos_in_queue < total_in_queue - 1:
            if st.button("次の項目 →", use_container_width=True):
                st.session_state["edit_idx"] = review_idxs[pos_in_queue + 1]
                st.rerun()
        else:
            # 最後の項目 → 確認画面へ進む
            if st.button("▶　確認画面へ", use_container_width=True):
                raw       = st.session_state.get("raw", [])
                vol_limit = VOL.get(st.session_state.get("vol_choice", "少（上位8件）"), 8)
                by_cat_lim = {}
                by_cat_tmp: dict[str, list] = {c: [] for c in CATS}
                for it in raw:
                    c = it.get("category", "課題")
                    if c in by_cat_tmp:
                        by_cat_tmp[c].append(it)
                by_cat_lim = {c: items[:vol_limit] for c, items in by_cat_tmp.items()}
                selected = _build_selected_from_review(by_cat_lim)
                st.session_state["selected"]  = selected
                st.session_state["ui_phase"]  = PHASE_CONFIRM
                st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)  # wn
    render_footer()


# ------------------------------------------------------------------------------
# CONFIRM 画面 — プレビュー・スコア・ダウンロード
# ------------------------------------------------------------------------------

def render_completeness_score(sel: dict) -> None:
    """
    報告品質スコアパネル（3軸評価）。
    CONFIRM 画面専用。
    """
    issue  = len(sel.get("課題", []))
    action = len(sel.get("対応", []))
    result = len(sel.get("結果", []))
    total  = issue + action + result
    if total == 0:
        return

    res_items = sel.get("結果", [])
    all_items = [it for its in sel.values() for it in its]
    n_res_num = sum(1 for it in res_items if has_num(it.get("short", "")) or has_num(it.get("original", "")))
    n_all_num = sum(1 for it in all_items if has_num(it.get("original", "")))
    score_num = min(100, int(n_res_num / max(1, len(res_items)) * 70 + n_all_num / max(1, total) * 30))

    all_three  = issue > 0 and action > 0 and result > 0
    cause_base = 100 if all_three else (60 if sum([issue > 0, action > 0, result > 0]) == 2 else 20)
    issue_words = set()
    for it in sel.get("課題", []):
        issue_words |= set(it.get("short", "").split())
    kw_overlap  = sum(
        1 for it in sel.get("対応", []) + sel.get("結果", [])
        if any(w in it.get("short", "") for w in issue_words if len(w) >= 2)
    )
    cause_score = min(100, cause_base + min(20, kw_overlap * 5))

    if total > 0:
        ratios  = [issue / total, action / total, result / total]
        balance = max(0, 100 - int((max(ratios) - min(ratios)) * 200))
    else:
        balance = 0

    overall = int((score_num + cause_score + balance) / 3)

    def _cls(v: int) -> str:
        return "sf-g" if v >= 70 else ("sf-a" if v >= 40 else "sf-r")
    def _pcls(v: int) -> str:
        return "sp-g" if v >= 70 else ("sp-a" if v >= 40 else "sp-r")

    if overall >= 70:   vrd_cls, vrd_msg = "sv-ok",   "✅ このまま提出できます"
    elif overall >= 45: vrd_cls, vrd_msg = "sv-warn", "⚠️ 数値か因果関係を補強してください"
    else:               vrd_cls, vrd_msg = "sv-ng",   "❌ 内容を確認してください"

    axes = [("数値の充実度", score_num), ("因果の完全性", cause_score), ("バランス", balance)]
    axes_html = "".join(
        f'<div class="score-axis">'
        f'<span class="score-lbl">{lbl}</span>'
        f'<div class="score-bg"><div class="score-fill {_cls(v)}" style="width:{v}%;"></div></div>'
        f'<span class="score-pct {_pcls(v)}">{v}%</span>'
        f'</div>'
        for lbl, v in axes
    )
    st.markdown(
        f'<div class="score-wrap">'
        f'<div class="score-ttl">📋 報告品質スコア</div>'
        f'{axes_html}'
        f'<div class="score-verdict {vrd_cls}">{vrd_msg}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )


def render_slide_preview(sel: dict) -> None:
    """スライドのライトモードプレビュー。"""
    cls_map = {"課題": "i", "対応": "a", "結果": "r"}
    secs = ""
    for cat in CATS:
        items = sel.get(cat, []); cd = CAT[cat]; cl = cls_map[cat]
        rows  = "".join(f'<div class="sp-it">・{it["short"]}</div>' for it in items[:5])
        rows  = rows or '<div class="sp-em">（未入力）</div>'
        if len(items) > 5:
            rows += f'<div class="sp-em">… 他 {len(items)-5} 件</div>'
        secs += (
            f'<div class="sp-sec">'
            f'<div class="sp-sn {cl}">{cd.icon} {cat}</div>'
            f'{rows}</div>'
        )
    total = sum(len(v) for v in sel.values())
    st.markdown(
        f'<div class="sp">'
        f'<div class="sp-top"><span class="sp-ttl">📄 月次報告レポート</span>'
        f'<span class="sp-date">完成イメージ</span></div>'
        f'<div class="sp-body">{secs}</div>'
        f'<div class="sp-ft">'
        f'<span>課題 {len(sel.get("課題",[]))} / 対応 {len(sel.get("対応",[]))} / 結果 {len(sel.get("結果",[]))}</span>'
        f'<span>計 {total} 件</span></div>'
        f'</div>',
        unsafe_allow_html=True,
    )


def render_confirm_screen() -> None:
    """
    【目的】ここだけ見ればOK、の最終確認画面。
    【表示するもの — 3セクションのみ】
    1. KPI（数値）   : selected["結果"] から数値含むものを最大5件、編集可能
    2. 課題（要約）  : selected["課題"] を最大3件、短文のみ
    3. 抜け漏れ入力欄: テキスト入力1つ、生成時に結果として追加
    【削除したもの】
    ・スライドプレビュー全体
    ・品質スコアパネル
    ・全抽出データ一覧 / 詳細テキスト / ファイル情報
    """
    render_header(PHASE_CONFIRM)

    selected: dict = st.session_state.get("selected", {c: [] for c in CATS})

    # ── 生成前の確認画面 ──────────────────────────────────────────────
    if not st.session_state.get("pptx_ready"):
        render_page_title(
            "STEP 3  最終確認",
            "3つの項目だけ確認してください",
            "問題なければそのまま生成できます",
        )
        st.markdown('<div class="wn">', unsafe_allow_html=True)

        # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        # セクション 1 — KPI（数値）
        # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        st.markdown(
            '<div class="sh">📊 KPI・数値実績</div>',
            unsafe_allow_html=True,
        )

        # 結果カテゴリから数値含むものを優先、なければ全件、最大5件
        result_items = selected.get("結果", [])
        kpi_items = [it for it in result_items if has_num(it.get("short", "")) or has_num(it.get("original", ""))]
        if not kpi_items:
            kpi_items = result_items  # 数値なくても全件フォールバック
        kpi_items = kpi_items[:5]

        if not kpi_items:
            st.markdown(
                '<div style="padding:12px;background:var(--bg2);border:1px dashed var(--b1);'
                'border-radius:var(--r-sm);font-size:12px;color:var(--t3);">'
                '結果カテゴリに項目がありません。下の入力欄から追加できます。</div>',
                unsafe_allow_html=True,
            )
        else:
            for i, it in enumerate(kpi_items):
                sk = f"confirm_kpi_{i}"
                if sk not in st.session_state:
                    st.session_state[sk] = it.get("short", "")
                st.text_input(
                    f"KPI {i + 1}",
                    key=sk,
                    label_visibility="collapsed",
                )
            # セッションの編集値を selected に反映
            for i, it in enumerate(kpi_items):
                sk = f"confirm_kpi_{i}"
                it["short"] = st.session_state.get(sk, it.get("short", ""))

        # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        # セクション 2 — 課題（要約）
        # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        st.markdown(
            '<div class="sh" style="margin-top:20px;">⚠️ 課題（最大3件）</div>',
            unsafe_allow_html=True,
        )

        issue_items = selected.get("課題", [])[:3]

        if not issue_items:
            st.markdown(
                '<div style="padding:12px;background:var(--bg2);border:1px dashed var(--b1);'
                'border-radius:var(--r-sm);font-size:12px;color:var(--t3);">'
                '課題カテゴリに項目がありません。</div>',
                unsafe_allow_html=True,
            )
        else:
            for it in issue_items:
                st.markdown(
                    f'<div style="background:var(--card);border:1px solid var(--b0);'
                    f'border-left:3px solid var(--red-l);border-radius:var(--r-sm);'
                    f'padding:10px 13px;margin-bottom:6px;font-size:13px;color:var(--t1);'
                    f'line-height:1.5;">{it.get("short", "")}</div>',
                    unsafe_allow_html=True,
                )

        # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        # セクション 3 — 抜け漏れ入力欄
        # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        st.markdown(
            '<div class="sh" style="margin-top:20px;">✏️ 抜け漏れがあれば追加</div>',
            unsafe_allow_html=True,
        )
        extra = st.text_input(
            "追加情報",
            key="confirm_extra",
            label_visibility="collapsed",
            placeholder="例：売上前月比 120%　／　クレーム件数 0件",
        )

        # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        # 生成ボタン
        # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        st.markdown('<hr>', unsafe_allow_html=True)

        total = sum(len(v) for v in selected.values())
        if total == 0 and not extra.strip():
            st.warning("⚠️ 採用された項目が0件です。一覧に戻って確認してください。")
        else:
            # 抜け漏れ入力があれば selected["結果"] に追加してから生成
            def _build_final(sel: dict, extra_text: str) -> dict:
                import copy
                final = copy.deepcopy(sel)
                if extra_text.strip():
                    final.setdefault("結果", []).append({
                        "original": extra_text.strip(),
                        "short":    extra_text.strip(),
                        "source":   "手動入力（確認画面）",
                        "category": "結果",
                        "score":    20,
                        "confidence": 1.0,
                    })
                return final

            c1, c2, c3 = st.columns([1, 2, 1])
            with c2:
                if st.button("⬡　スライドを生成する", use_container_width=True):
                    if not PPTX_OK:
                        st.error("❌ python-pptx が必要: pip install python-pptx")
                    else:
                        final_sel = _build_final(selected, extra)
                        pr = st.progress(0); pr.progress(30)
                        today = datetime.now().strftime("%Y年%m月%d日 %H:%M")
                        pb    = build_pptx(final_sel, today)
                        pr.progress(100); pr.empty()
                        st.session_state.update({
                            "pptx_bytes": pb,
                            "pptx_ready": True,
                            "show_toast": True,
                            "last_sel":   {k: list(v) for k, v in final_sel.items()},
                        })
                        st.rerun()

        # 戻るボタン（小さく）
        st.markdown('<div style="margin-top:12px;"></div>', unsafe_allow_html=True)
        c1, c2, c3 = st.columns([2, 1, 2])
        with c2:
            if st.button("← 一覧に戻る", use_container_width=True):
                st.session_state.pop("pptx_ready", None)
                st.session_state.pop("pptx_bytes", None)
                st.session_state["ui_phase"] = PHASE_REVIEW
                st.rerun()

        st.markdown('</div>', unsafe_allow_html=True)

    # ── 生成完了画面 ─────────────────────────────────────────────────
    if st.session_state.get("pptx_ready"):
        render_page_title(
            "STEP 3  生成完了",
            "スライドの生成が<em>完了しました</em>",
            "下のボタンを押してPowerPointファイルを保存してください",
        )
        if st.session_state.pop("show_toast", False):
            st.toast("✅ スライドの生成が完了しました", icon="✅")

        _s = st.session_state.get("last_sel", {})
        total_dl = sum(len(v) for v in _s.values()) if _s else 0

        st.markdown('<div class="wn">', unsafe_allow_html=True)
        st.markdown(
            '<div class="dl-card" id="dla">'
            '<div class="dl-ttl">✅ 生成が完了しました</div>'
            f'<div class="dl-sub">計 {total_dl} 件を整理しました。</div>'
            '</div>',
            unsafe_allow_html=True,
        )

        st.markdown('<div class="pulse">', unsafe_allow_html=True)
        c1, c2, c3 = st.columns([1, 2, 1])
        with c2:
            st.download_button(
                label="⬇　IIJ月次報告レポートをダウンロード",
                data=st.session_state["pptx_bytes"],
                file_name=f"IIJ_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.pptx",
                mime=(
                    "application/vnd.openxmlformats-officedocument"
                    ".presentationml.presentation"
                ),
                use_container_width=True,
            )
        st.markdown('</div>', unsafe_allow_html=True)  # pulse

        st.markdown('<div style="margin-top:14px;"></div>', unsafe_allow_html=True)
        c1, c2, c3 = st.columns([2, 1, 2])
        with c2:
            if st.button("🔄 内容を修正する", use_container_width=True):
                st.session_state.pop("pptx_ready", None)
                st.session_state.pop("pptx_bytes", None)
                st.session_state["ui_phase"] = PHASE_REVIEW
                st.rerun()

        st.markdown('</div>', unsafe_allow_html=True)  # wn

        components.html(
            """<script>
            (function(){
              function go(){
                var e=parent.document.getElementById('dla');
                if(e) e.scrollIntoView({behavior:'smooth',block:'center'});
                else  setTimeout(go,100);
              }
              setTimeout(go,300);
            })();
            </script>""",
            height=0,
        )

    render_footer()


# ==============================================================================
# LAYER 9 ─ メインアプリ (v14 4画面ルーター)
#
# セッション状態 ui_phase の値で画面をルーティングする。
# ロジック・データ処理はLAYER 0〜7が担う。
# この関数はルーターとしての役割のみ持つ。
# ==============================================================================

def main() -> None:
    st.markdown(CSS, unsafe_allow_html=True)

    # ui_phase が未設定 → upload へ
    if "ui_phase" not in st.session_state:
        if st.session_state.get("raw"):
            st.session_state["ui_phase"] = PHASE_REVIEW
        else:
            st.session_state["ui_phase"] = PHASE_UPLOAD

    phase = st.session_state["ui_phase"]

    if phase == PHASE_UPLOAD:
        render_upload_screen()
    elif phase == PHASE_REVIEW:
        render_review_screen()
    elif phase == PHASE_EDIT:
        render_edit_screen()
    elif phase == PHASE_CONFIRM:
        render_confirm_screen()
    else:
        st.session_state["ui_phase"] = PHASE_UPLOAD
        st.rerun()


if __name__ == "__main__":
    main()
