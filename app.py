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
from generate_report import generate_report_core, load_campaign_targets, parse_ko_xlsx

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
    hits = glob.glob(os.path.join(SCRIPT_DIR, 'SMS*.xlsx'))
    if hits:
        return hits[0]
    for name in ['SMSターゲット.xlsx', 'SMS対象.xlsx', 'sms_targets.xlsx']:
        p = os.path.join(SCRIPT_DIR, name)
        if os.path.exists(p):
            return p
    return None

TARGET_FILE = os.path.join(SCRIPT_DIR, 'sms_targets.xlsx') if os.path.exists(os.path.join(SCRIPT_DIR, 'sms_targets.xlsx')) else find_target_file()

# ── フォームリセット用カウンター初期化 ──────────
if 'form_key' not in st.session_state:
    st.session_state['form_key'] = 0

def reset_form():
    st.session_state['form_key'] += 1

fk = st.session_state['form_key']  # ウィジェットキーのサフィックス

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
if TARGET_FILE and os.path.exists(TARGET_FILE):
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
    key=f'xlsx_{fk}',
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
        key=f'scroll_{fk}',
    )
with col2:
    attention_file = st.file_uploader(
        'Attention CSV（注目割合）',
        type=['csv'],
        help='Clarity の Attention CSV',
        key=f'attention_{fk}',
    )

lp_image_file = st.file_uploader(
    'LP画像（スクリーンショット）',
    type=['jpg', 'jpeg', 'png', 'webp'],
    help='スマホで撮影した LP のスクリーンショット等',
    key=f'image_{fk}',
)

# ── 事前解析（XLSX解析：STEP 3 で共通利用）──
import hashlib as _hashlib
import re as _re
_auto_hash = 'none'
_meta_auto: dict = {}
if xlsx_file is not None:
    try:
        _xlsx_bytes = xlsx_file.read()
        xlsx_file.seek(0)
        _auto_hash = _hashlib.md5(_xlsx_bytes).hexdigest()[:8]
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as _tf:
            _tf.write(_xlsx_bytes)
            _tf.flush()
            _tmp_path = _tf.name
        _meta_auto, _, _ = parse_ko_xlsx(_tmp_path)
        os.unlink(_tmp_path)
    except Exception:
        pass

_cb_key = f'{fk}_{_auto_hash}'

# ════════════════════════════════════════════
# STEP 3 : 入力項目
# ════════════════════════════════════════════
st.markdown('<div class="section-title">STEP 3 ─ 店舗情報</div>', unsafe_allow_html=True)

col_m, col_c = st.columns(2)
with col_m:
    machines_input = st.text_input(
        '総台数 ＊必須',
        value='',
        placeholder='例：578（半角数字）',
        help='店舗の設置台数（パチンコ＋スロット合計）。半角数字で入力してください',
        key=f'machines_{fk}',
    )
    # 半角数字のみ受け付ける（全角・文字列はNG）
    import re as _re
    _machines_valid = bool(_re.fullmatch(r'[0-9]+', machines_input)) and 1 <= int(machines_input) <= 9999 if machines_input else False
    machines = int(machines_input) if _machines_valid else None
    if machines_input and not _machines_valid:
        st.caption('⚠️ 半角数字（1〜9999）で入力してください')
with col_c:
    campaign_type = st.selectbox(
        'キャンペーンタイプ ＊必須',
        options=campaign_names if campaign_names else ['（SMSターゲット.xlsx が見つかりません）'],
        index=None,
        placeholder='── 選択してください ──',
        help='SMSターゲット.xlsx に定義されているキャンペーン種別（必ず選択してください）',
        key=f'campaign_{fk}',
    )

# 想定効果測定日数：自動読み取り or 手動入力
_auto_mdays = st.checkbox(
    '想定効果測定日数を自動読み取りする',
    value=True,
    key=f'auto_mdays_{fk}',
)
st.caption('チェックを外すと手動で変更できます。')
_auto_mdays_val = _meta_auto.get('measurementDays', 0)
_mdays_display  = str(_auto_mdays_val) if _auto_mdays_val > 0 else ''
if _auto_mdays:
    st.text_input(
        '想定効果測定日数（日）',
        value=_mdays_display,
        placeholder='ファイルをアップロードすると自動計算されます',
        disabled=True,
        key=f'mdays_auto_{_cb_key}',
    )
    measurement_days = None
else:
    _mdays_input = st.text_input(
        '想定効果測定日数（日）',
        value=_mdays_display,
        placeholder='例：7（半角数字）',
        key=f'mdays_manual_{_cb_key}',
    )
    _mdays_valid = bool(_re.fullmatch(r'[0-9]+', _mdays_input)) and 1 <= int(_mdays_input) <= 365 if _mdays_input else False
    measurement_days = int(_mdays_input) if _mdays_valid else None
    if _mdays_input and not _mdays_valid:
        st.caption('⚠️ 半角数字（1〜365）で入力してください')

# ════════════════════════════════════════════
# STEP 4 : SMS本文チェック
# ════════════════════════════════════════════
st.markdown('<div class="section-title">STEP 4 ─ SMS本文チェック</div>', unsafe_allow_html=True)
st.caption('各項目の判定方法を選択してください。「自動判別」はレポート生成時に自動で判定します。')

_SMS_ITEMS = [
    '店名の記載',
    'お客様名の記載',
    'チラ見せ度',
    '緊急性・限定感',
    'CTAの明確さ',
    'お店の思い・温かみ',
    '文章の使い回し感',
]

_sms_vals = {}
for _label in _SMS_ITEMS:
    _c1, _c2 = st.columns([2, 3])
    with _c1:
        st.markdown(f'**{_label}**')
    with _c2:
        _sms_vals[_label] = st.radio(
            _label,
            options=['自動判別', '有', '無'],
            index=0,
            horizontal=True,
            label_visibility='collapsed',
            key=f'sms_{_cb_key}_{_label}',
        )

# 自動判別→None、有/無はそのまま渡す（generate_report_core がNone=自動検出として処理）
store_name_status    = None if _sms_vals['店名の記載']         == '自動判別' else _sms_vals['店名の記載']
customer_name_status = None if _sms_vals['お客様名の記載']     == '自動判別' else _sms_vals['お客様名の記載']
warmth_status        = None if _sms_vals['お店の思い・温かみ'] == '自動判別' else _sms_vals['お店の思い・温かみ']
tease_status         = None if _sms_vals['チラ見せ度']         == '自動判別' else _sms_vals['チラ見せ度']
urgency_status       = None if _sms_vals['緊急性・限定感']     == '自動判別' else _sms_vals['緊急性・限定感']
cta_status           = None if _sms_vals['CTAの明確さ']        == '自動判別' else _sms_vals['CTAの明確さ']
reuse_status         = None if _sms_vals['文章の使い回し感']   == '自動判別' else _sms_vals['文章の使い回し感']

# ════════════════════════════════════════════
# ボタン行（生成 ＋ リセット）
# ════════════════════════════════════════════
st.divider()
btn_col1, btn_col2 = st.columns([3, 1])
with btn_col1:
    _btn_disabled = (
        (xlsx_file is None) or
        (campaign_type is None and bool(campaign_names)) or
        (machines is None)
    )
    generate_btn = st.button('🚀 レポートを生成する', type='primary', disabled=_btn_disabled)
with btn_col2:
    st.button('🔄 リセット', on_click=reset_form)

if xlsx_file is None:
    st.info('KO XLSX をアップロードするとレポートを生成できます。')
elif campaign_type is None and bool(campaign_names):
    st.warning('キャンペーンタイプを選択してください。')

# ── 生成処理 ──────────────────────────────────
if generate_btn and xlsx_file is not None:
    with st.spinner('レポートを生成中...'):
        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                def save_upload(uploaded, suffix):
                    if uploaded is None:
                        return None
                    path = os.path.join(tmpdir, f'upload{suffix}')
                    with open(path, 'wb') as f:
                        f.write(uploaded.read())
                    uploaded.seek(0)
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
                    tease_status          = tease_status,
                    urgency_status        = urgency_status,
                    cta_status            = cta_status,
                    reuse_status          = reuse_status,
                    measurement_days      = measurement_days,
                )

            import re
            send_id_match = re.search(r'sendId:\s*[\'"](\d+)[\'"]', html)
            store_match   = re.search(r'store:\s*[\'"]([^\'"]+)[\'"]', html)
            send_id   = send_id_match.group(1) if send_id_match else 'output'
            store_raw = store_match.group(1) if store_match else ''
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
