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
from generate_report import generate_report_core, load_campaign_targets, analyze_sms, parse_ko_xlsx

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

# ════════════════════════════════════════════
# STEP 4 : SMS本文チェック（手動選択）
# ════════════════════════════════════════════
st.markdown('<div class="section-title">STEP 4 ─ SMS本文チェック（手動選択）</div>', unsafe_allow_html=True)
st.caption('KOレポートから自動判定が難しい項目を選択してください。「自動判定」にすると本文テキストから推定します。')

col_s1, col_s2 = st.columns(2)
with col_s1:
    _store_sel = st.radio('店名の記載', ['自動判定','有','無'], horizontal=True,
                          help='SMS本文に店名が記載されているか', key=f'store_{fk}')
with col_s2:
    _customer_sel = st.radio('お客様名の記載', ['自動判定','有','無'], horizontal=True,
                             help='SMS本文にお客様の個人名が差し込まれているか', key=f'customer_{fk}')

col_s3, col_s4 = st.columns(2)
with col_s3:
    _warmth_sel = st.radio('お店の思い・温かみ', ['自動判定','有','無'], horizontal=True,
                           help='感謝・期待感など温かみのある表現があるか', key=f'warmth_{fk}')
with col_s4:
    _generic_sel = st.radio('汎用フレーズのみ', ['自動判定','改善不要','要改善'], horizontal=True,
                            help='要改善＝汎用フレーズのみで具体性がない / 改善不要＝具体的な内容が含まれている', key=f'generic_{fk}')

col_s5, _ = st.columns(2)
with col_s5:
    _hook_sel = st.radio('興味喚起フック', ['自動判定','有','無'], horizontal=True,
                         help='数字・限定・固有ワードなどフックとなる表現があるか', key=f'hook_{fk}')

store_name_status    = None if _store_sel    == '自動判定' else _store_sel
customer_name_status = None if _customer_sel == '自動判定' else _customer_sel
warmth_status        = None if _warmth_sel   == '自動判定' else _warmth_sel
# 「要改善」→内部値「有」（汎用フレーズのみ＝問題あり）、「改善不要」→「無」
generic_status       = None if _generic_sel  == '自動判定' else ('有' if _generic_sel == '要改善' else '無')
hook_status          = None if _hook_sel     == '自動判定' else _hook_sel

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
                    generic_status        = generic_status,
                    hook_status           = hook_status,
                )

            import re
            send_id_match = re.search(r'sendId:\s*[\'"](\d+)[\'"]', html)
            store_match   = re.search(r'store:\s*[\'"]([^\'"]+)[\'"]', html)
            send_id   = send_id_match.group(1) if send_id_match else 'output'
            store_raw = store_match.group(1) if store_match else ''
            store_safe = re.sub(r'[\\/:*?"<>|\s]', '_', store_raw)
            filename = f'{store_safe}_{send_id}.html' if store_safe else f'report_{send_id}.html'

            # ── 手動設定と自動判定の矛盾チェック ──
            try:
                _, _, meta_q = parse_ko_xlsx(xlsx_path)
                sms_text_q = meta_q.get('smsText', '')
                store_q    = meta_q.get('store', '')
                if sms_text_q:
                    auto = {c['label']: c for c in analyze_sms(sms_text_q, store_name=store_q)['checks']}
                    # label → (手動値, 手動値がwarnになる内部status, 手動値がokになる内部status)
                    checks_map = [
                        ('店名の記載',        store_name_status,    '有', 'ok',   '無', 'warn'),
                        ('お客様名の記載',    customer_name_status, '有', 'ok',   '無', 'na'),
                        ('お店の思い・温かみ', warmth_status,       '有', 'ok',   '無', 'warn'),
                        ('汎用フレーズのみ',  generic_status,       '無', 'ok',   '有', 'warn'),
                        ('興味喚起フック',    hook_status,          '有', 'ok',   '無', 'warn'),
                    ]
                    for label, manual_val, ok_val, ok_st, ng_val, ng_st in checks_map:
                        if manual_val is None or label not in auto:
                            continue
                        auto_status = auto[label]['status']
                        auto_detail = auto[label]['detail']
                        manual_status = ok_st if manual_val == ok_val else ng_st
                        if manual_status != auto_status:
                            ui_val = ('要改善' if manual_val == '有' else '改善不要') if label == '汎用フレーズのみ' else manual_val
                            st.warning(
                                f'⚠️ **{label}**：手動で「{ui_val}」に設定されていますが、'
                                f'自動判定では異なる結果でした。\n'
                                f'自動判定の根拠：{auto_detail}'
                            )
            except Exception:
                pass  # 矛盾チェック失敗はサイレントに無視

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
