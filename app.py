#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SMS分析レポート Web UI
=======================
起動方法:
  streamlit run app.py
"""

import os
import glob
import tempfile
import streamlit as st
from generate_report import generate_report_core, load_campaign_targets

# ── ページ設定 ────────────────────────────────
st.set_page_config(
    page_title='SMS分析レポート 生成ツール',
    page_icon='📊',
    layout='centered',
)

# ── パスワード認証 ────────────────────────────
def check_password():
    correct = st.secrets.get('app_password', '')
    if not correct:
        st.error('管理者設定が必要です（Secrets未設定）')
        st.stop()
    if st.session_state.get('authenticated'):
        return
    st.title('🔐 ログイン')
    pw = st.text_input('パスワード', type='password')
    if st.button('ログイン'):
        if pw == correct:
            st.session_state['authenticated'] = True
            st.rerun()
        else:
            st.error('パスワードが違います')
    st.stop()

check_password()

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# SMSターゲットファイルをglobで自動検出（ファイル名の揺れに対応）
def find_target_file():
    # 「SMS」を含む xlsx を探す
    hits = glob.glob(os.path.join(SCRIPT_DIR, 'SMS*.xlsx'))
    if hits:
        return hits[0]
    # フォールバック：候補名リスト
    for name in ['SMSターゲット.xlsx', 'SMS対象.xlsx', 'sms_targets.xlsx']:
        p = os.path.join(SCRIPT_DIR, name)
        if os.path.exists(p):
            return p
    return None

TARGET_FILE = os.path.join(SCRIPT_DIR, 'sms_targets.xlsx') if os.path.exists(os.path.join(SCRIPT_DIR, 'sms_targets.xlsx')) else find_target_file()

# ── スタイル ──────────────────────────────────
st.markdown("""
<style>
  .block-container { max-width: 740px; padding-top: 2rem; }
  .stButton > button { width: 100%; height: 3rem; font-size: 1rem; font-weight: 700; }
  .section-title { font-size: 0.85rem; font-weight: 700; color: #64748b;
                   text-transform: uppercase; letter-spacing: .05em; margin: 1.5rem 0 .5rem; }
</style>
""", unsafe_allow_html=True)

# ── タイトル ──────────────────────────────────
st.title('📊 SMS分析レポート 生成ツール')
st.caption('必要なファイルをアップロードして「レポート生成」を押してください。')
st.divider()

# ── キャンペーンタイプ一覧を取得 ───────────────
campaign_targets = []
if os.path.exists(TARGET_FILE):
    campaign_targets = load_campaign_targets(TARGET_FILE)
campaign_names = [ct['name'] for ct in campaign_targets]

# ════════════════════════════════════════════
# STEP 1 : 必須ファイル
# ════════════════════════════════════════════
st.markdown('<div class="section-title">STEP 1 ─ 必須ファイル</div>', unsafe_allow_html=True)

xlsx_file = st.file_uploader(
    'KO XLSX（送信結果）',
    type=['xlsx'],
    help='20260609_1200_45889_KO.xlsx のような KO フォーマットのファイル',
)

# ════════════════════════════════════════════
# STEP 2 : 任意ファイル
# ════════════════════════════════════════════
st.markdown('<div class="section-title">STEP 2 ─ 任意ファイル（あれば精度が上がります）</div>', unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    scroll_file = st.file_uploader(
        'Scroll CSV（到達率）',
        type=['csv'],
        help='Clarity の Scroll 深度 CSV',
    )
with col2:
    attention_file = st.file_uploader(
        'Attention CSV（注目割合）',
        type=['csv'],
        help='Clarity の Attention CSV',
    )

lp_image_file = st.file_uploader(
    'LP画像（スクリーンショット）',
    type=['jpg', 'jpeg', 'png', 'webp'],
    help='スマホで撮影した LP のスクリーンショット等',
)

# ════════════════════════════════════════════
# STEP 3 : 入力項目
# ════════════════════════════════════════════
st.markdown('<div class="section-title">STEP 3 ─ 店舗情報</div>', unsafe_allow_html=True)

col_m, col_c = st.columns(2)
with col_m:
    machines = st.number_input(
        '総台数',
        min_value=1,
        max_value=9999,
        value=500,
        step=1,
        help='店舗の設置台数（パチンコ＋スロット合計）',
    )
with col_c:
    campaign_type = st.selectbox(
        'キャンペーンタイプ',
        options=campaign_names if campaign_names else ['（SMSターゲット.xlsx が見つかりません）'],
        help='SMSターゲット.xlsx に定義されているキャンペーン種別',
    )

# ════════════════════════════════════════════
# STEP 4 : SMS本文チェック（手動選択）
# ════════════════════════════════════════════
st.markdown('<div class="section-title">STEP 4 ─ SMS本文チェック（手動選択）</div>', unsafe_allow_html=True)
st.caption('KOレポートから自動判定が難しい項目を選択してください。「自動判定」にすると本文テキストから推定します。')

col_s1, col_s2 = st.columns(2)
with col_s1:
    _store_sel = st.radio('店名の記載', ['自動判定','有','無'], horizontal=True,
                          help='SMS本文に店名が記載されているか')
with col_s2:
    _customer_sel = st.radio('お客様名の記載', ['自動判定','有','無'], horizontal=True,
                             help='SMS本文にお客様の個人名が差し込まれているか')

col_s3, col_s4 = st.columns(2)
with col_s3:
    _warmth_sel = st.radio('お店の思い・温かみ', ['自動判定','有','無'], horizontal=True,
                           help='感謝・期待感など温かみのある表現があるか')
with col_s4:
    _generic_sel = st.radio('汎用フレーズのみ', ['自動判定','有','無'], horizontal=True,
                            help='有＝汎用フレーズのみで具体性がない（要改善）')

col_s5, _ = st.columns(2)
with col_s5:
    _hook_sel = st.radio('興味喚起フック', ['自動判定','有','無'], horizontal=True,
                         help='数字・限定・固有ワードなどフックとなる表現があるか')

store_name_status    = None if _store_sel    == '自動判定' else _store_sel
customer_name_status = None if _customer_sel == '自動判定' else _customer_sel
warmth_status        = None if _warmth_sel   == '自動判定' else _warmth_sel
generic_status       = None if _generic_sel  == '自動判定' else _generic_sel
hook_status          = None if _hook_sel     == '自動判定' else _hook_sel

# ════════════════════════════════════════════
# 生成ボタン
# ════════════════════════════════════════════
st.divider()
generate_btn = st.button('🚀 レポートを生成する', type='primary', disabled=(xlsx_file is None))

if xlsx_file is None:
    st.info('KO XLSX をアップロードするとレポートを生成できます。')

# ── 生成処理 ──────────────────────────────────
if generate_btn and xlsx_file is not None:
    with st.spinner('レポートを生成中...'):
        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                # ファイルを一時ディレクトリに保存
                def save_upload(uploaded, suffix):
                    if uploaded is None:
                        return None
                    path = os.path.join(tmpdir, f'upload{suffix}')
                    with open(path, 'wb') as f:
                        f.write(uploaded.read())
                    uploaded.seek(0)  # 再読み込み可能にリセット
                    return path

                xlsx_path      = save_upload(xlsx_file,      '.xlsx')
                scroll_path    = save_upload(scroll_file,    '_scroll.csv')
                attention_path = save_upload(attention_file, '_attention.csv')
                image_path     = save_upload(lp_image_file,  os.path.splitext(lp_image_file.name)[1] if lp_image_file else '.jpg')

                html = generate_report_core(
                    xlsx_path             = xlsx_path,
                    scroll_csv_path       = scroll_path,
                    attention_csv_path    = attention_path,
                    image_path            = image_path,
                    machines              = int(machines),
                    campaign_type         = campaign_type if campaign_names else None,
                    script_dir            = SCRIPT_DIR,
                    store_name_status     = store_name_status,
                    customer_name_status  = customer_name_status,
                    warmth_status         = warmth_status,
                    generic_status        = generic_status,
                    hook_status           = hook_status,
                )

            # 出力ファイル名を 店名_ID.html 形式で決定
            import re
            send_id_match = re.search(r'sendId:\s*[\'"](\d+)[\'"]', html)
            store_match   = re.search(r'store:\s*[\'"]([^\'"]+)[\'"]', html)
            send_id   = send_id_match.group(1) if send_id_match else 'output'
            store_raw = store_match.group(1) if store_match else ''
            # ファイル名に使えない文字を除去
            store_safe = re.sub(r'[\\/:*?"<>|\s]', '_', store_raw)
            filename = f'{store_safe}_{send_id}.html' if store_safe else f'report_{send_id}.html'

            st.success(f'✅ レポート生成完了！（{len(html) // 1024} KB）')

            st.download_button(
                label     = f'📥 {filename} をダウンロード',
                data      = html.encode('utf-8'),
                file_name = filename,
                mime      = 'text/html',
            )

        except Exception as e:
            st.error(f'エラーが発生しました：{e}')
            import traceback
            st.code(traceback.format_exc())

# ════════════════════════════════════════════
# フッター
# ════════════════════════════════════════════
st.divider()
st.caption('© Pix Inc. All Rights Reserved.')
