#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SMSåæã¬ãã¼ã Web UI
=======================
èµ·åæ¹æ³:
  streamlit run app.py
"""

import os
import glob
import tempfile
import streamlit as st
from generate_report import generate_report_core, load_campaign_targets, analyze_sms, parse_ko_xlsx

# ââ ãã¼ã¸è¨­å® ââââââââââââââââââââââââââââââââ
st.set_page_config(
    page_title='SMSåæã¬ãã¼ã çæãã¼ã«',
    page_icon='ð',
    layout='centered',
)

# ââ ãã¹ã¯ã¼ãèªè¨¼ ââââââââââââââââââââââââââââ
def check_password():
    correct = st.secrets.get('app_password', '')
    if not correct:
        st.error('ç®¡çèè¨­å®ãå¿è¦ã§ãï¼Secretsæªè¨­å®ï¼')
        st.stop()
    if st.session_state.get('authenticated'):
        return
    st.title('ð ã­ã°ã¤ã³')
    pw = st.text_input('ãã¹ã¯ã¼ã', type='password')
    if st.button('ã­ã°ã¤ã³'):
        if pw == correct:
            st.session_state['authenticated'] = True
            st.rerun()
        else:
            st.error('ãã¹ã¯ã¼ããéãã¾ã')
    st.stop()

check_password()

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# SMSã¿ã¼ã²ãããã¡ã¤ã«ãglobã§èªåæ¤åºï¼ãã¡ã¤ã«åã®æºãã«å¯¾å¿ï¼
def find_target_file():
    hits = glob.glob(os.path.join(SCRIPT_DIR, 'SMS*.xlsx'))
    if hits:
        return hits[0]
    for name in ['SMSã¿ã¼ã²ãã.xlsx', 'SMSå¯¾è±¡.xlsx', 'sms_targets.xlsx']:
        p = os.path.join(SCRIPT_DIR, name)
        if os.path.exists(p):
            return p
    return None

TARGET_FILE = os.path.join(SCRIPT_DIR, 'sms_targets.xlsx') if os.path.exists(os.path.join(SCRIPT_DIR, 'sms_targets.xlsx')) else find_target_file()

# ââ ãã©ã¼ã ãªã»ããç¨ã«ã¦ã³ã¿ã¼åæå ââââââââââ
if 'form_key' not in st.session_state:
    st.session_state['form_key'] = 0

def reset_form():
    st.session_state['form_key'] += 1

fk = st.session_state['form_key']  # ã¦ã£ã¸ã§ããã­ã¼ã®ãµãã£ãã¯ã¹

# ââ ã¹ã¿ã¤ã« ââââââââââââââââââââââââââââââââââ
st.markdown("""
<style>
  .block-container { max-width: 740px; padding-top: 2rem; }
  .stButton > button { width: 100%; height: 3rem; font-size: 1rem; font-weight: 700; }
  .section-title { font-size: 0.85rem; font-weight: 700; color: #64748b;
                   text-transform: uppercase; letter-spacing: .05em; margin: 1.5rem 0 .5rem; }
</style>
""", unsafe_allow_html=True)

# ââ ã¿ã¤ãã« ââââââââââââââââââââââââââââââââââ
st.title('ð SMSåæã¬ãã¼ã çæãã¼ã«')
st.caption('å¿è¦ãªãã¡ã¤ã«ãã¢ããã­ã¼ããã¦ãã¬ãã¼ãçæããæ¼ãã¦ãã ããã')
st.divider()

# ââ ã­ã£ã³ãã¼ã³ã¿ã¤ãä¸è¦§ãåå¾ âââââââââââââââ
campaign_targets = []
if TARGET_FILE and os.path.exists(TARGET_FILE):
    campaign_targets = load_campaign_targets(TARGET_FILE)
campaign_names = [ct['name'] for ct in campaign_targets]

# ââââââââââââââââââââââââââââââââââââââââââââ
# STEP 1 : å¿é ãã¡ã¤ã«
# ââââââââââââââââââââââââââââââââââââââââââââ
st.markdown('<div class="section-title">STEP 1 â å¿é ãã¡ã¤ã«</div>', unsafe_allow_html=True)

xlsx_file = st.file_uploader(
    'KO XLSXï¼éä¿¡çµæï¼',
    type=['xlsx'],
    help='20260609_1200_45889_KO.xlsx ã®ãã KO ãã©ã¼ãããã®ãã¡ã¤ã«',
    key=f'xlsx_{fk}',
)

# ââââââââââââââââââââââââââââââââââââââââââââ
# STEP 2 : ä»»æãã¡ã¤ã«
# ââââââââââââââââââââââââââââââââââââââââââââ
st.markdown('<div class="section-title">STEP 2 â ä»»æãã¡ã¤ã«ï¼ããã°ç²¾åº¦ãä¸ããã¾ãï¼</div>', unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    scroll_file = st.file_uploader(
        'Scroll CSVï¼å°éçï¼',
        type=['csv'],
        help='Clarity ã® Scroll æ·±åº¦ CSV',
        key=f'scroll_{fk}',
    )
with col2:
    attention_file = st.file_uploader(
        'Attention CSVï¼æ³¨ç®å²åï¼',
        type=['csv'],
        help='Clarity ã® Attention CSV',
        key=f'attention_{fk}',
    )

lp_image_file = st.file_uploader(
    'LPç»åï¼ã¹ã¯ãªã¼ã³ã·ã§ããï¼',
    type=['jpg', 'jpeg', 'png', 'webp'],
    help='ã¹ããã§æ®å½±ãã LP ã®ã¹ã¯ãªã¼ã³ã·ã§ããç­',
    key=f'image_{fk}',
)

# ââââââââââââââââââââââââââââââââââââââââââââ
# STEP 3 : å¥åé ç®
# ââââââââââââââââââââââââââââââââââââââââââââ
st.markdown('<div class="section-title">STEP 3 â åºèæå ±</div>', unsafe_allow_html=True)

col_m, col_c = st.columns(2)
with col_m:
    machines_input = st.text_input(
        'ç·å°æ° ï¼å¿é ',
        value='',
        placeholder='ä¾ï¼578ï¼åè§æ°å­ï¼',
        help='åºèã®è¨­ç½®å°æ°ï¼ããã³ã³ï¼ã¹ã­ããåè¨ï¼ãåè§æ°å­ã§å¥åãã¦ãã ãã',
        key=f'machines_{fk}',
    )
    # åè§æ°å­ã®ã¿åãä»ããï¼å¨è§ã»æå­åã¯NGï¼
    import re as _re
    _machines_valid = bool(_re.fullmatch(r'[0-9]+', machines_input)) and 1 <= int(machines_input) <= 9999 if machines_input else False
    machines = int(machines_input) if _machines_valid else None
    if machines_input and not _machines_valid:
        st.caption('â ï¸ åè§æ°å­ï¼1ã9999ï¼ã§å¥åãã¦ãã ãã')
with col_c:
    campaign_type = st.selectbox(
        'ã­ã£ã³ãã¼ã³ã¿ã¤ã ï¼å¿é ',
        options=campaign_names if campaign_names else ['ï¼SMSã¿ã¼ã²ãã.xlsx ãè¦ã¤ããã¾ããï¼'],
        index=None,
        placeholder='ââ é¸æãã¦ãã ãã ââ',
        help='SMSã¿ã¼ã²ãã.xlsx ã«å®ç¾©ããã¦ããã­ã£ã³ãã¼ã³ç¨®å¥ï¼å¿ãé¸æãã¦ãã ããï¼',
        key=f'campaign_{fk}',
    )

# ââââââââââââââââââââââââââââââââââââââââââââ
# STEP 4 : SMSæ¬æãã§ãã¯ï¼æåé¸æï¼
# ââââââââââââââââââââââââââââââââââââââââââââ
st.markdown('<div class="section-title">STEP 4 â SMSæ¬æãã§ãã¯ï¼æåé¸æï¼</div>', unsafe_allow_html=True)
st.caption('KOã¬ãã¼ãããèªåå¤å®ãé£ããé ç®ãé¸æãã¦ãã ããããèªåå¤å®ãã«ããã¨æ¬æãã­ã¹ãããæ¨å®ãã¾ãã')

col_s1, col_s2 = st.columns(2)
with col_s1:
    _store_sel = st.radio('åºåã®è¨è¼', ['èªåå¤å®','æ','ç¡'], horizontal=True,
                          help='SMSæ¬æã«åºåãè¨è¼ããã¦ããã', key=f'store_{fk}')
with col_s2:
    _customer_sel = st.radio('ãå®¢æ§åã®è¨è¼', ['èªåå¤å®','æ','ç¡'], horizontal=True,
                             help='SMSæ¬æã«ãå®¢æ§ã®åäººåãå·®ãè¾¼ã¾ãã¦ããã', key=f'customer_{fk}')

col_s3, col_s4 = st.columns(2)
with col_s3:
    _warmth_sel = st.radio('ãåºã®æãã»æ¸©ãã¿', ['èªåå¤å®','æ','ç¡'], horizontal=True,
                           help='æè¬ã»æå¾æãªã©æ¸©ãã¿ã®ããè¡¨ç¾ãããã', key=f'warmth_{fk}')
with col_s4:
    _generic_sel = st.radio('æ±ç¨ãã¬ã¼ãºã®ã¿', ['èªåå¤å®','æ¹åä¸è¦','è¦æ¹å'], horizontal=True,
                            help='è¦æ¹åï¼æ±ç¨ãã¬ã¼ãºã®ã¿ã§å·ä½æ§ããªã / æ¹åä¸è¦ï¼å·ä½çãªåå®¹ãå«ã¾ãã¦ãã', key=f'generic_{fk}')

col_s5, _ = st.columns(2)
with col_s5:
    _hook_sel = st.radio('èå³åèµ·ããã¯', ['èªåå¤å®','æ','ç¡'], horizontal=True,
                         help='æ°å­ã»éå®ã»åºæã¯ã¼ããªã©ããã¯ã¨ãªãè¡¨ç¾ãããã', key=f'hook_{fk}')

store_name_status    = None if _store_sel    == 'èªåå¤å®' else _store_sel
customer_name_status = None if _customer_sel == 'èªåå¤å®' else _customer_sel
warmth_status        = None if _warmth_sel   == 'èªåå¤å®' else _warmth_sel
# ãè¦æ¹åãâåé¨å¤ãæãï¼æ±ç¨ãã¬ã¼ãºã®ã¿ï¼åé¡ããï¼ããæ¹åä¸è¦ãâãç¡ã
generic_status       = None if _generic_sel  == 'èªåå¤å®' else ('æ' if _generic_sel == 'è¦æ¹å' else 'ç¡')
hook_status          = None if _hook_sel     == 'èªåå¤å®' else _hook_sel

# ââââââââââââââââââââââââââââââââââââââââââââ
# ãã¿ã³è¡ï¼çæ ï¼ ãªã»ããï¼
# ââââââââââââââââââââââââââââââââââââââââââââ
st.divider()
btn_col1, btn_col2 = st.columns([3, 1])
with btn_col1:
    _btn_disabled = (
        (xlsx_file is None) or
        (campaign_type is None and bool(campaign_names)) or
        (machines is None)
    )
    generate_btn = st.button('ð ã¬ãã¼ããçæãã', type='primary', disabled=_btn_disabled)
with btn_col2:
    st.button('ð ãªã»ãã', on_click=reset_form)

if xlsx_file is None:
    st.info('KO XLSX ãã¢ããã­ã¼ãããã¨ã¬ãã¼ããçæã§ãã¾ãã')
elif campaign_type is None and bool(campaign_names):
    st.warning('ã­ã£ã³ãã¼ã³ã¿ã¤ããé¸æãã¦ãã ããã')

# ââ çæå¦ç ââââââââââââââââââââââââââââââââââ
if generate_btn and xlsx_file is not None:
    with st.spinner('ã¬ãã¼ããçæä¸­...'):
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

            # ââ æåè¨­å®ã¨èªåå¤å®ã®çç¾ãã§ãã¯ ââ
            try:
                _, _, meta_q = parse_ko_xlsx(xlsx_path)
                sms_text_q = meta_q.get('smsText', '')
                store_q    = meta_q.get('store', '')
                if sms_text_q:
                    auto = {c['label']: c for c in analyze_sms(sms_text_q, store_name=store_q)['checks']}
                    # label â (æåå¤, æåå¤ãwarnã«ãªãåé¨status, æåå¤ãokã«ãªãåé¨status)
                    checks_map = [
                        ('åºåã®è¨è¼',        store_name_status,    'æ', 'ok',   'ç¡', 'warn'),
                        ('ãå®¢æ§åã®è¨è¼',    customer_name_status, 'æ', 'ok',   'ç¡', 'na'),
                        ('ãåºã®æãã»æ¸©ãã¿', warmth_status,       'æ', 'ok',   'ç¡', 'warn'),
                        ('æ±ç¨ãã¬ã¼ãºã®ã¿',  generic_status,       'ç¡', 'ok',   'æ', 'warn'),
                        ('èå³åèµ·ããã¯',    hook_status,          'æ', 'ok',   'ç¡', 'warn'),
                    ]
                    for label, manual_val, ok_val, ok_st, ng_val, ng_st in checks_map:
                        if manual_val is None or label not in auto:
                            continue
                        auto_status = auto[label]['status']
                        auto_detail = auto[label]['detail']
                        manual_status = ok_st if manual_val == ok_val else ng_st
                        if manual_status != auto_status:
                            ui_val = ('è¦æ¹å' if manual_val == 'æ' else 'æ¹åä¸è¦') if label == 'æ±ç¨ãã¬ã¼ãºã®ã¿' else manual_val
                            st.warning(
                                f'â ï¸ **{label}**ï¼æåã§ã{ui_val}ãã«è¨­å®ããã¦ãã¾ããã'
                                f'èªåå¤å®ã§ã¯ç°ãªãçµæã§ããã\n'
                                f'èªåå¤å®ã®æ ¹æ ï¼{auto_detail}'
                            )
            except Exception:
                pass  # çç¾ãã§ãã¯å¤±æã¯ãµã¤ã¬ã³ãã«ç¡è¦

            st.success(f'â ã¬ãã¼ãçæå®äºï¼ï¼{len(html) // 1024} KBï¼')
            st.download_button(
                label     = f'ð¥ {filename} ããã¦ã³ã­ã¼ã',
                data      = html.encode('utf-8'),
                file_name = filename,
                mime      = 'text/html',
            )

        except Exception as e:
            st.error(f'ã¨ã©ã¼ãçºçãã¾ããï¼{e}')
            import traceback
            st.code(traceback.format_exc())

# ââââââââââââââââââââââââââââââââââââââââââââ
# ããã¿ã¼
# ââââââââââââââââââââââââââââââââââââââââââââ
st.divider()
st.caption('Â© Pix Inc. All Rights Reserved.')
