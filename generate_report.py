#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SMS分析レポート 自動生成スクリプト
====================================
使い方（KO形式）:
  python3 generate_report.py --xlsx 43345_KO.xlsx --csv 43345.csv --output report_43345.html

使い方（Clarityのみ）:
  python3 generate_report.py --csv 44402.csv --output report_44402.html

オプション:
  --xlsx      KOフォーマットのXLSXファイル（ファイル名自由）
  --csv       ClarityのCSVファイル（ファイル名自由）
  --image     LPページ画像ファイル（省略時はXLSX/CSVと同名を自動検索）
  --output    出力HTMLファイル名（デフォルト: report_output.html）
  --template  テンプレートHTMLパス（省略時は同フォルダのsms_report_template.html）
"""

import argparse
import csv
import json
import base64
import os
import re
import sys
from collections import defaultdict
import openpyxl

def find_data_file(script_dir, candidates):
    """複数の候補ファイル名から存在するものを返す（日本語ファイル名の揺れ対応）"""
    for name in candidates:
        p = os.path.join(script_dir, name)
        if os.path.exists(p):
            return p
    return None

# ── 離反期間バケット ──────────────────────────
SEG_ORDER = ['1ヶ月未満','1〜2ヶ月','2〜3ヶ月','3〜6ヶ月','半年〜1年','1〜3年','3年以上']

# カード使用率プレースホルダー（後で実績値に更新）
CARD_RATES = {
    '1ヶ月未満': 0.85,
    '1〜2ヶ月':  0.60,
    '2〜3ヶ月':  0.50,
    '3〜6ヶ月':  0.35,
    '半年〜1年': 0.20,
    '1〜3年':    0.10,
    '3年以上':   0.05,
}

# 来店結果報告の列マッピング（0始まり）
VISIT_RATE_COLS = {
    '1ヶ月未満': 2,
    '1〜2ヶ月':  3,
    '2〜3ヶ月':  4,
    '3〜6ヶ月':  5,
    '半年〜1年': 6,
    '1〜3年':    7,
    '3年以上':   8,
}

def load_visit_rates(xlsx_path):
    """来店結果報告_平滑化版.xlsxから効果測定日数別・離反期間別来店率を読み込む"""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb.active
    rates = {}  # {day: {segment_label: rate}}
    for row in ws.iter_rows(min_row=3, values_only=True):
        if row[0] is None: continue
        try:
            day = int(row[0])
        except (TypeError, ValueError):
            continue
        day_rates = {}
        for label, col_idx in VISIT_RATE_COLS.items():
            val = row[col_idx]
            if val is not None:
                day_rates[label] = float(val)
        rates[day] = day_rates
    return rates

def get_segment(days):
    if days is None: return None
    if days < 30:    return '1ヶ月未満'
    if days < 60:    return '1〜2ヶ月'
    if days < 90:    return '2〜3ヶ月'
    if days < 180:   return '3〜6ヶ月'
    if days < 365:   return '半年〜1年'
    if days < 1095:  return '1〜3年'
    return '3年以上'

# ── SMSターゲット種別（最大離反期間・バースデー特別処理）────────
def load_campaign_targets(xlsx_path):
    """SMSターゲット.xlsx を読み込み、種別リストを返す
    戻り値: [{'name': str, 'max_seg': str, 'birthday': bool}, ...]
    """
    if not os.path.exists(xlsx_path):
        return []
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    result = []
    for row in rows[1:]:  # ヘッダー行スキップ
        if not row[0]: continue
        name    = str(row[0]).strip()
        max_seg = str(row[1]).strip() if row[1] else '1〜3年'
        note    = str(row[2]).strip() if row[2] else ''
        birthday = '12分の1' in note or 'バースデー' in name
        result.append({'name': name, 'max_seg': max_seg, 'birthday': birthday})
    return result

# ── 会員データ（台数別推定）────────────────────
MEMBER_DATA_SEG_MAP = {
    '1ヶ月未満':   '1ヶ月未満',
    '1〜2ヶ月未満': '1〜2ヶ月',
    '2〜3ヶ月未満': '2〜3ヶ月',
    '3～6ヶ月未満': '3〜6ヶ月',
    '半年～1年未満': '半年〜1年',
    '1年～3年未満': '1〜3年',
}

def load_member_estimates(xlsx_path, machines):
    """設置台数から各離反期間の推定会員数（排他的人数）を線形補間で算出。

    テーブルの人数列は累積値（その期間以内の全会員数）のため、
    排他的人数 = 当行の累積値 - 直前行の累積値 として計算する。
    例: 1〜2ヶ月のみの人数 = (2ヶ月未満累積) - (1ヶ月未満累積)
    """
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))

    lower = (machines // 100) * 100
    upper = lower + 100
    frac = (machines - lower) / 100

    # 台数→列インデックス: X台 → col = X/100 * 2
    lo_col = (lower // 100) * 2
    hi_col = (upper // 100) * 2

    # まず累積値を SEG_ORDER 順に収集
    cumulative = {}
    for row in rows[2:]:
        seg_key = row[1] if len(row) > 1 else None
        if seg_key not in MEMBER_DATA_SEG_MAP: continue
        label = MEMBER_DATA_SEG_MAP[seg_key]
        if label in cumulative: continue  # 最初の有効な行のみ（後半の#REF!行をスキップ）
        lo_val = row[lo_col] if len(row) > lo_col and row[lo_col] is not None else 0
        hi_val = row[hi_col] if len(row) > hi_col and row[hi_col] is not None else lo_val
        if not isinstance(lo_val, (int, float)): continue
        if not isinstance(hi_val, (int, float)): hi_val = lo_val
        cumulative[label] = round(lo_val + frac * (hi_val - lo_val))

    # 累積値→排他的人数（差分）に変換
    result = {}
    prev = 0
    for label in SEG_ORDER:
        if label == '3年以上': continue
        cum = cumulative.get(label, prev)
        result[label] = max(0, cum - prev)
        prev = cum
    return result

# ── 年代バケット ──────────────────────────────
AGE_ORDER = ['20代まで','30代','40代','50代','60代以上']

def get_age_segment(age):
    if age is None: return None
    if age <= 29: return '20代まで'
    if age <= 39: return '30代'
    if age <= 49: return '40代'
    if age <= 59: return '50代'
    return '60代以上'


# ── KO XLSX パーサー ─────────────────────────
def parse_ko_xlsx(filepath, visit_rate_data=None):
    """KOフォーマットのXLSXを読み込んでmeta・segments・KPIを返す"""
    try:
        import openpyxl
    except ImportError:
        print('openpyxlが必要です: pip install openpyxl --break-system-packages')
        sys.exit(1)

    wb = openpyxl.load_workbook(filepath)
    meta = {}

    # ── メタ情報（会員支持率詳細情報シート）
    ws1 = wb['会員支持率詳細情報']
    for row in ws1.iter_rows():
        for cell in row:
            v = cell.value
            if v is None: continue
            if v == '店舗名：':
                meta['store'] = ws1.cell(row=cell.row, column=cell.column+5).value
            elif v == '送信日時：':
                d = ws1.cell(row=cell.row, column=cell.column+6).value
                t = ws1.cell(row=cell.row, column=cell.column+16).value
                if d and t:
                    meta['datetime'] = f"{str(d)[:10]}T{str(t)}"
            elif v == '送信ID：':
                meta['sendId'] = str(int(ws1.cell(row=cell.row, column=cell.column+5).value))
            elif v == 'タイトル：':
                meta['campaign'] = ws1.cell(row=cell.row, column=cell.column+6).value
            elif v == '送信予約数':
                # 同じ列を下に探す（他列の値を誤取得しないよう列固定）
                col = cell.column
                for r in range(cell.row + 1, cell.row + 10):
                    val = ws1.cell(row=r, column=col).value
                    if isinstance(val, (int, float)) and val > 100:
                        meta['reserved'] = int(val)
                        break

    # ── SMS本文（K8 = 行8, 列11）
    sms_text = ws1.cell(row=8, column=11).value or ''
    meta['smsText'] = str(sms_text).strip()

    # ── 会員別データ集計（各会員様別送信詳細シート）
    ws2 = wb['各会員様別送信詳細']
    buckets = defaultdict(lambda: {'sent':0,'lp':0,'multi':0,'visits':0})
    total_sent = total_lp = total_multi = total_visits = 0
    max_last_visit = None  # 効果測定日数算出用

    for row in ws2.iter_rows(min_row=7, values_only=True):
        if row[0] is None: continue
        result   = row[3]
        lp_views = row[5] or 0
        days     = row[11]
        visited  = row[18]

        if result != '送達': continue
        total_sent += 1
        if lp_views >= 1: total_lp    += 1
        if lp_views >= 2: total_multi += 1
        if visited == '〇':
            total_visits += 1
            # col21（最終来店日 after sending）= row[20]
            last_visit_val = row[20]
            if last_visit_val:
                from datetime import date as _date
                if hasattr(last_visit_val, 'date'):
                    lv = last_visit_val.date()
                elif isinstance(last_visit_val, str):
                    try:
                        lv = _date.fromisoformat(last_visit_val[:10].replace('/', '-'))
                    except Exception:
                        lv = None
                else:
                    lv = None
                if lv and (max_last_visit is None or lv > max_last_visit):
                    max_last_visit = lv

        seg = get_segment(days)
        if seg is None: continue
        buckets[seg]['sent']   += 1
        if lp_views >= 1: buckets[seg]['lp']     += 1
        if lp_views >= 2: buckets[seg]['multi']  += 1
        if visited == '〇': buckets[seg]['visits'] += 1

    meta['sent']       = total_sent
    meta['lpViews']    = total_lp
    meta['multiViews'] = total_multi
    meta['visits']     = total_visits

    # 効果測定日数 = 送信日〜最終来店日（最大1〜）
    if max_last_visit and meta.get('datetime'):
        from datetime import date as _date
        send_dt = _date.fromisoformat(meta['datetime'][:10])
        measurement_days = max(1, (max_last_visit - send_dt).days)
    else:
        measurement_days = 0
    meta['measurementDays'] = measurement_days
    meta['pageViews']  = sum((row[5] or 0) for row in wb['各会員様別送信詳細'].iter_rows(min_row=7, values_only=True) if row[0] and row[3] == '送達')

    # ── 平均LP閲覧率・来店転換率（象限分類用）
    avg_lp_rate    = total_lp    / total_sent * 100 if total_sent else 0
    avg_visit_rate = total_visits / total_sent * 100 if total_sent else 0

    # ── 離反期間別セグメント配列生成
    segments = []
    for label in SEG_ORDER:
        d = buckets[label]
        if d['sent'] == 0: continue
        lp_rate    = round(d['lp']     / d['sent'] * 100, 1)
        visit_rate = round(d['visits'] / d['sent'] * 100, 1)
        # 来店結果報告データがあれば効果測定日数ベースの実績値を使用、なければフォールバック
        if visit_rate_data and meta.get('measurementDays'):
            ref_days = min(meta['measurementDays'], 30)
            card_rate = visit_rate_data.get(ref_days, {}).get(label, CARD_RATES.get(label, 0.5))
        else:
            card_rate = CARD_RATES.get(label, 0.5)

        # 象限分類（X=来店転換率、Y=LP閲覧率）
        high_visit = visit_rate >= avg_visit_rate
        high_lp    = lp_rate    >= avg_lp_rate
        if   high_visit and high_lp:    tag = 'q1'  # 優良反応層
        elif high_visit and not high_lp: tag = 'q2'  # 高来店層
        elif not high_visit and high_lp: tag = 'q3'  # 関心限定層
        else:                            tag = 'q4'  # 低反応層

        segments.append({
            'label':     label,
            'sent':      d['sent'],
            'lp':        d['lp'],
            'multi':     d['multi'],
            'visits':    d['visits'],
            'lpRate':    lp_rate,
            'visitRate': visit_rate,
            'cardRate':  card_rate,
            'tag':       tag,
        })

    # ── 年代別セグメント配列生成（4象限用）
    age_buckets = defaultdict(lambda: {'sent':0,'lp':0,'multi':0,'visits':0})
    wb2 = openpyxl.load_workbook(filepath)  # 再読み込み不要だが再利用
    ws2b = wb2['各会員様別送信詳細']
    for row in ws2b.iter_rows(min_row=7, values_only=True):
        if row[0] is None: continue
        if row[3] != '送達': continue
        age      = row[9]   # 年齢
        lp_views = row[5] or 0
        visited  = row[18]
        age_seg  = get_age_segment(age)
        if age_seg is None: continue
        age_buckets[age_seg]['sent']   += 1
        if lp_views >= 1: age_buckets[age_seg]['lp']     += 1
        if lp_views >= 2: age_buckets[age_seg]['multi']  += 1
        if visited == '〇': age_buckets[age_seg]['visits'] += 1

    age_segments = []
    for label in AGE_ORDER:
        d = age_buckets[label]
        if d['sent'] == 0: continue
        lp_rate    = round(d['lp']     / d['sent'] * 100, 1)
        visit_rate = round(d['visits'] / d['sent'] * 100, 1)
        high_visit = visit_rate >= avg_visit_rate
        high_lp    = lp_rate    >= avg_lp_rate
        if   high_visit and high_lp:     tag = 'q1'
        elif high_visit and not high_lp: tag = 'q2'
        elif not high_visit and high_lp: tag = 'q3'
        else:                            tag = 'q4'
        age_segments.append({
            'label':     label,
            'sent':      d['sent'],
            'lp':        d['lp'],
            'multi':     d['multi'],
            'visits':    d['visits'],
            'lpRate':    lp_rate,
            'visitRate': visit_rate,
            'tag':       tag,
        })

    return meta, segments, age_segments


# ── Clarity CSV パーサー ──────────────────────
def classify(pct):
    if pct >= 7.0:  return 'q1', 'keep'
    elif pct >= 3.0: return 'q2', 'fix'
    elif pct >= 1.5: return 'q3', 'fix'
    else:            return 'q4', 'del'

def parse_time(time_str):
    parts = time_str.strip().split(':')
    if len(parts) == 3:
        total = int(parts[0])*3600 + int(parts[1])*60 + int(parts[2])
        return f'{total//60}m{total%60}s' if total >= 60 else f'{total}s'
    return time_str

def parse_clarity_csv(filepath):
    scroll_depths = []
    meta = {}
    with open(filepath, encoding='utf-8-sig') as f:
        rows = list(csv.reader(f))
    for row in rows:
        if len(row) >= 2:
            if row[0] == 'プロジェクト名':    meta['project']   = row[1]
            elif row[0] == '閲覧済み URL を含む': meta['url']  = row[1]
            elif row[0] == 'ページ ビュー':   meta['pageViews'] = int(row[1])
    for row in rows:
        if len(row) >= 3 and row[0].strip().lstrip('"').isdigit():
            depth   = int(row[0].strip().strip('"'))
            time_raw = row[1].strip().strip('"')
            pct     = float(row[2].strip().strip('"').replace('%',''))
            tag, judge = classify(pct)
            scroll_depths.append({'depth':depth,'time':parse_time(time_raw),'pct':pct,'tag':tag,'judge':judge})
    return meta, scroll_depths

def detect_csv_type(filepath):
    fname = os.path.basename(filepath).lower()
    with open(filepath, encoding='utf-8-sig') as f:
        head = f.read(500)
    if 'メトリック' in head and 'Scroll' in head and 'ドロップ' in head:
        return 'scroll'
    if 'メトリック' in head and 'Attention' in head:
        return 'attention'
    if 'プロジェクト名' in head and 'スクロールの奥行き' in head:
        return 'attention'
    return 'unknown'

def parse_scroll_csv(filepath):
    """到達率CSV（Scroll metric）: スクロールの奥行き, 訪問者数, % ドロップ オフ"""
    reach_map = {}
    with open(filepath, encoding='utf-8-sig') as f:
        rows = list(csv.reader(f))
    max_visitors = None
    for row in rows:
        if len(row) >= 2 and row[0].strip().lstrip('"').isdigit():
            depth = int(row[0].strip().strip('"'))
            try:
                visitors = float(row[1].strip().strip('"'))
            except:
                continue
            if max_visitors is None:
                max_visitors = visitors
            reach = round(visitors / max_visitors * 100) if max_visitors else 100
            reach_map[depth] = reach
    return reach_map


# ── 画像変換 ──────────────────────────────────
def image_to_base64(filepath):
    ext  = os.path.splitext(filepath)[1].lower().lstrip('.')
    mime = {'jpg':'image/jpeg','jpeg':'image/jpeg','png':'image/png','webp':'image/webp'}.get(ext,'image/jpeg')
    with open(filepath,'rb') as f:
        return f'data:{mime};base64,{base64.b64encode(f.read()).decode()}'


# ── HTML注入 ──────────────────────────────────
def inject_scroll_depths(html, scroll_depths):
    new_str = 'scrollDepths: ' + json.dumps(scroll_depths, ensure_ascii=False, indent=4) + ','
    pattern = r'// LP スクロール深度データ.*?scrollDepths: \[.*?\],\n'
    return re.sub(pattern, '  // LP スクロール深度データ（自動生成）\n  ' + new_str + '\n', html, flags=re.DOTALL)

def analyze_sms(text, store_name=''):
    """SMS本文の品質を簡易チェックして評価結果を返す"""
    if not text:
        return {'score': 'unknown', 'checks': [], 'charCount': 0}

    checks = []
    url_removed = re.sub(r'https?://\S+', '', text).strip()

    # ① 店名が入っているか（改行を除去して比較）
    text_flat = text.replace('\n', '').replace('\r', '')
    store_in_text = store_name and (store_name[:4] in text_flat or store_name[:6] in text_flat)
    checks.append({
        'label': '店名の記載',
        'status': 'ok' if store_in_text else 'warn',
        'detail': '店名が記載されており、どこからのSMSか明確です' if store_in_text
                  else '店名がないと迷惑SMSと勘違いされる可能性があります'
    })

    # ② お客様氏名（個人名差し込みのみ検出。「お客様」等の汎用語は除外）
    personal_name_patterns = ['{', '【氏名】', '[氏名]', '氏名', 'お名前', '○○様', '〇〇様', '◯◯様']
    has_personal_name = any(p in text for p in personal_name_patterns)
    # 「様」が「お客様」以外で使われている場合（実名差し込み）も検出
    if not has_personal_name:
        import re as _re
        has_personal_name = bool(_re.search(r'(?<!客)様', text))
    checks.append({
        'label': 'お客様名の記載',
        'status': 'ok' if has_personal_name else 'na',
        'detail': 'お客様名が差し込まれており承認欲求に働きかけています' if has_personal_name
                  else '個人名差し込みは確認できません（KOレポートからは判定不可）'
    })

    # ③ お店の思いが伝わるか（温かみ・感情語の有無）
    warm_words = ['ありがとう', 'おかげさま', 'うれし', '感謝', '喜', 'お待ち', 'ぜひ',
                  '皆様', 'みなさま', '心より', '特別', 'とっておき', '大切', '思い']
    has_warmth = any(w in text for w in warm_words)
    content_len = len(re.sub(r'https?://\S+|\s|　', '', url_removed))
    too_short = content_len < 15
    checks.append({
        'label': 'お店の思い・温かみ',
        'status': 'ok' if (has_warmth and not too_short) else 'warn',
        'detail': 'お店の気持ちが感じられる文章です' if (has_warmth and not too_short)
                  else '作業的・短すぎる文章はお客様の興味を惹きにくい場合があります。感謝・期待感を添えると効果的です'
    })

    # ④ 汎用フレーズのみチェック
    generic_phrases = ['大切なお客様へのお知らせ', 'お知らせです', 'ご案内です', '詳しくはこちら']
    has_generic_only = any(p in text for p in generic_phrases) and content_len < 20

    # ok時：URLを除いたテキストで具体的な内容を特定
    text_no_url = re.sub(r'https?://\S+', '', text)
    found_specifics = []
    nums = re.findall(r'\d+[周年台日月曜火水木金土%倍円]?', text_no_url)
    if nums: found_specifics.append(f"「{'・'.join(nums[:3])}」などの具体的な数字")
    hook_found = [w for w in ['周年', '限定', '新台', 'イベント', '増台', '復活', 'お得'] if w in text_no_url]
    if hook_found: found_specifics.append(f"「{'・'.join(hook_found)}」といった具体的なキーワード")
    specific_detail = '、'.join(found_specifics) + 'が含まれており、汎用文だけで終わっていません' if found_specifics \
                      else '汎用フレーズに依存しない文章構成です'

    checks.append({
        'label': '汎用フレーズのみ',
        'status': 'warn' if has_generic_only else 'ok',
        'detail': '汎用フレーズだけで具体的なフックがありません。「何があるの？」と思わせる一言を追加しましょう' if has_generic_only
                  else specific_detail
    })

    # ⑤ 興味喚起フック
    has_number = bool(re.search(r'\d+', text))
    has_hook_words = any(w in text for w in ['周年', '限定', '初', '新台', '復活', '秘密', 'お得', '増台', 'イベント'])
    has_hook = has_number or has_hook_words
    checks.append({
        'label': '興味喚起フック',
        'status': 'ok' if has_hook else 'warn',
        'detail': '数字・固有フレーズで「気になる」を演出できています' if has_hook
                  else 'URLを開きたくなる具体的なフックがありません'
    })

    # 文字数（URL・改行含む）
    char_count = len(text)

    # ⑥ 文字数（情報表示のみ。長い・短いで優劣なし）
    if char_count <= 70:
        char_status, char_detail = 'info', f'{char_count}字：基本料金範囲内（70字以内）'
    elif char_count <= 133:
        char_status, char_detail = 'info', f'{char_count}字：推奨範囲内（133字以内）。追加料金あり'
    else:
        char_status, char_detail = 'warn', f'{char_count}字：133字を超えています。料金が上がる可能性があります'
    checks.append({
        'label': f'文字数 {char_count}字',
        'status': char_status,
        'detail': char_detail
    })

    warn_count = sum(1 for c in checks if c['status'] == 'warn')
    na_count   = sum(1 for c in checks if c['status'] == 'na')
    if warn_count >= 3:
        score = 'review'
    elif warn_count > 0 or na_count > 0:
        score = 'caution'
    else:
        score = 'good'
    return {'score': score, 'checks': checks, 'charCount': char_count}


def inject_sms_analysis(html, sms_text, analysis):
    """SMS本文と分析結果をDATAオブジェクトのプレースホルダーに注入"""
    # json.dumpsで確実にエスケープ
    text_json   = json.dumps(sms_text, ensure_ascii=False)
    checks_json = json.dumps(analysis['checks'], ensure_ascii=False)
    # プレースホルダーを実データで置換
    html = html.replace(
        "  smsText: '',\n  smsScore: 'unknown',\n  smsChecks: [],",
        f"  smsText: {text_json},\n  smsScore: '{analysis['score']}',\n  smsChecks: {checks_json},"
    )
    return html


def inject_age_segments(html, age_segments):
    new_str = 'ageSegments: ' + json.dumps(age_segments, ensure_ascii=False, indent=4) + ','
    pattern = r'// 年代別データ.*?ageSegments: \[.*?\],\n'
    result = re.sub(pattern, '  // 年代別データ（自動生成）\n  ' + new_str + '\n', html, flags=re.DOTALL)
    if result == html:  # パターンが見つからない場合はsegmentsの直後に挿入
        html = html.replace('  // 離反期間別データ（自動生成）', '  // 年代別データ（自動生成）\n  ' + new_str + '\n\n  // 離反期間別データ（自動生成）')
        return html
    return result

def inject_segments(html, segments):
    new_str = 'segments: ' + json.dumps(segments, ensure_ascii=False, indent=4) + ','
    pattern = r'// 離反期間別データ.*?segments: \[.*?\],\n'
    return re.sub(pattern, '  // 離反期間別データ（自動生成）\n  ' + new_str + '\n', html, flags=re.DOTALL)

def inject_unsent_segments(html, segments, member_estimates, visit_rate_data, measurement_days,
                           max_segment=None, birthday_mode=False):
    """設置台数ベースの推定会員数と送信済み数から未送信層を計算してDATAに注入。
    max_segment: この離反期間を超えるセグメントはシミュレーションに含めない
    birthday_mode: True の場合、追加対象人数を 1/12 に補正（誕生日月限定施策）
    """
    ref_days = min(measurement_days, 30) if measurement_days else None

    # 最大対象期間のインデックスを決定
    max_idx = len(SEG_ORDER) - 1  # デフォルト：全セグメント
    if max_segment and max_segment in SEG_ORDER:
        max_idx = SEG_ORDER.index(max_segment)

    unsent = []
    for i, label in enumerate(SEG_ORDER):
        if label == '3年以上': continue
        if i > max_idx: continue  # キャンペーン種別の最大対象期間を超えたらスキップ
        est_total = member_estimates.get(label, 0)
        seg = next((s for s in segments if s['label'] == label), None)
        sent_count = seg['sent'] if seg else 0
        unsent_count = max(0, est_total - sent_count)
        if unsent_count == 0: continue
        # バースデー施策：誕生日月のお客様のみ対象 → 1/12 に補正
        if birthday_mode:
            unsent_count = max(1, unsent_count // 12)
        if visit_rate_data and ref_days:
            card_rate = visit_rate_data.get(ref_days, {}).get(label, CARD_RATES.get(label, 0.1))
        else:
            card_rate = CARD_RATES.get(label, 0.1)
        lp_rate = round(seg['lpRate'] / 100, 3) if seg else 0.25
        unsent.append({
            'label':       label,
            'estTotal':    est_total,
            'sentCount':   sent_count,
            'unsentCount': unsent_count,
            'lpRate':      lp_rate,
            'cardRate':    round(card_rate, 4),
        })
    new_str = 'unsentSegments: ' + json.dumps(unsent, ensure_ascii=False, indent=4) + ','
    pattern = r'unsentSegments:\s*\[.*?\],\n'
    return re.sub(pattern, new_str + '\n', html, flags=re.DOTALL)

def inject_kpi(html, meta):
    # KPIブロック全体を一括置換（行頭の2スペースインデントで限定）
    fields = {
        'reserved':        meta.get('reserved',        0),
        'sent':            meta.get('sent',             0),
        'lpViews':         meta.get('lpViews',          0),
        'multiViews':      meta.get('multiViews',       0),
        'totalPV':         meta.get('pageViews',        0),
        'visits':          meta.get('visits',           0),
        'measurementDays': meta.get('measurementDays',  0),
        'machines':        meta.get('machines',          0),
    }
    for key, val in fields.items():
        # 行頭2スペース + キー名: でKPIフィールドだけをマッチ
        html = re.sub(rf'(\n  {key}:\s*)\d+', rf'\g<1>{val}', html, count=1)
    return html

def inject_header(html, store=None, campaign=None, sendid=None, dt=None, axis=None):
    if store:    html = re.sub(r"store:\s*'[^']*'",    f"store: '{store}'",       html)
    if campaign: html = re.sub(r"campaign:\s*'[^']*'", f"campaign: '{campaign}'", html)
    if sendid:   html = re.sub(r"sendId:\s*'[^']*'",   f"sendId: '{sendid}'",     html)
    if dt:       html = re.sub(r"datetime:\s*'[^']*'", f"datetime: '{dt}'",       html)
    if axis:     html = re.sub(r"axis:\s*'[^']*'",     f"axis: '{axis}'",         html)
    return html

def inject_image(html, image_src):
    return re.sub(r'<img src="data:image/[^"]*" class="heatmap-img"',
                  f'<img src="{image_src}" class="heatmap-img"', html)

def inject_campaign_meta(html, campaign_type, max_segment):
    """campaignType / maxSegment を DATA に注入"""
    if campaign_type:
        html = re.sub(r"campaignType:\s*'[^']*'", f"campaignType: '{campaign_type}'", html)
    if max_segment:
        html = re.sub(r"maxSegment:\s*'[^']*'",   f"maxSegment: '{max_segment}'",     html)
    return html

def find_image(base_path):
    for ext in ['.jpg','.jpeg','.png','.webp']:
        p = base_path + ext
        if os.path.exists(p): return p
    return None


# ── コア処理（CLIもStreamlitも共通） ─────────────────────
def generate_report_core(
    xlsx_path=None,
    scroll_csv_path=None,
    attention_csv_path=None,
    image_path=None,
    machines=None,
    campaign_type=None,
    template_path=None,
    script_dir=None,
):
    """
    レポートHTMLを生成して文字列で返す。
    script_dir を指定しない場合はこのファイルのディレクトリを使用。
    """
    if script_dir is None:
        script_dir = os.path.dirname(os.path.abspath(__file__))

    # テンプレート読み込み
    tpl = template_path or os.path.join(script_dir, 'sms_report_template.html')
    with open(tpl, encoding='utf-8') as f:
        html = f.read()

    meta = {}
    segments = []
    scroll_depths = []

    # 来店結果報告データ
    visit_rate_data = None
    rate_file = find_data_file(script_dir, ['来店結果報告_平滑化版.xlsx', '報告結果_平滑化版.xlsx', 'visit_rates.xlsx'])
    if rate_file:
        visit_rate_data = load_visit_rates(rate_file)

    # XLSX（KO形式）
    age_segments = []
    if xlsx_path:
        meta, segments, age_segments = parse_ko_xlsx(xlsx_path, visit_rate_data=visit_rate_data)
        html = inject_segments(html, segments)
        html = inject_age_segments(html, age_segments)

    # 設置台数・キャンペーン種別
    target_file = find_data_file(script_dir, ['SMSターゲット.xlsx', 'SMS対象.xlsx', 'sms_targets.xlsx'])
    member_file = find_data_file(script_dir, ['会員データ調査_想定台数100刻み_算出済.xlsx', '会員データ調査_想定人数100刻み_算出済.xlsx', 'member_data.xlsx'])
    campaign_targets = load_campaign_targets(target_file) if target_file else []

    max_segment   = None
    birthday_mode = False
    if campaign_type:
        matched = next((ct for ct in campaign_targets if ct['name'] == campaign_type), None)
        if matched:
            max_segment   = matched['max_seg']
            birthday_mode = matched['birthday']

    if machines:
        meta['machines'] = machines
    if machines and xlsx_path and os.path.exists(member_file):
        member_estimates = load_member_estimates(member_file, machines)
        html = inject_unsent_segments(html, segments, member_estimates, visit_rate_data,
                                      meta.get('measurementDays', 0),
                                      max_segment=max_segment, birthday_mode=birthday_mode)
        if campaign_type:
            html = inject_campaign_meta(html, campaign_type, max_segment or '')

    # Scroll CSV（到達率）
    reach_map = {}
    if scroll_csv_path:
        reach_map = parse_scroll_csv(scroll_csv_path)

    # Attention CSV（注目割合）
    if attention_csv_path:
        csv_type = detect_csv_type(attention_csv_path)
        if csv_type in ('attention', 'clarity'):
            clarity_meta, scroll_depths = parse_clarity_csv(attention_csv_path)
            meta.update({k: v for k, v in clarity_meta.items() if k not in meta})
            if not meta.get('sendId') and clarity_meta.get('url'):
                m = re.search(r'/(\d+)$', clarity_meta['url'])
                if m: meta['sendId'] = m.group(1)
            if reach_map:
                for d in scroll_depths:
                    d['reach'] = reach_map.get(d['depth'], None)
            html = inject_scroll_depths(html, scroll_depths)
    elif reach_map:
        # Scroll のみの場合も到達率データを注入
        html = inject_scroll_depths(html, [])

    # SMS本文分析
    if xlsx_path and meta.get('smsText'):
        sms_analysis = analyze_sms(meta['smsText'], store_name=meta.get('store', ''))
        html = inject_sms_analysis(html, meta['smsText'], sms_analysis)

    # KPI・ヘッダー
    if xlsx_path:
        html = inject_kpi(html, meta)
        html = inject_header(html,
            store    = meta.get('store'),
            campaign = meta.get('campaign'),
            sendid   = meta.get('sendId'),
            dt       = meta.get('datetime'),
            axis     = '来店転換率 × LP支持率',
        )
    elif attention_csv_path:
        html = inject_header(html, sendid=meta.get('sendId'))

    # 画像
    if image_path:
        html = inject_image(html, image_to_base64(image_path))

    return html


# ── メイン ────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(description='SMS分析レポート自動生成')
    parser.add_argument('--xlsx',       default=None, help='KO XLSXファイル')
    parser.add_argument('--csv',        default=None, help='Clarity注目割合CSVファイル（Attention）')
    parser.add_argument('--scroll-csv', default=None, help='Clarityスクロール深度CSVファイル（Scroll）')
    parser.add_argument('--image',      default=None, help='LPページ画像ファイル')
    parser.add_argument('--machines',      default=None, type=int, help='店舗の設置台数（緩和シミュレーション用）')
    parser.add_argument('--campaign-type', default=None, help='SMSキャンペーン種別（未指定時はインタラクティブ選択）')
    parser.add_argument('--template', default=None, help='テンプレートHTMLパス')
    parser.add_argument('--output',   default='report_output.html', help='出力ファイル名')
    args = parser.parse_args()

    if not args.xlsx and not args.csv:
        print('エラー: --xlsx または --csv を指定してください')
        sys.exit(1)

    # テンプレート読み込み
    script_dir = os.path.dirname(os.path.abspath(__file__))
    template_path = args.template or os.path.join(script_dir, 'sms_report_template.html')
    if not os.path.exists(template_path):
        print(f'エラー: テンプレートが見つかりません → {template_path}')
        sys.exit(1)
    with open(template_path, encoding='utf-8') as f:
        html = f.read()

    meta = {}
    segments = []
    scroll_depths = []

    # ── XLSX（KO形式）処理
    # 来店結果報告データの自動読み込み
    visit_rate_data = None
    rate_file = find_data_file(script_dir, ['来店結果報告_平滑化版.xlsx', '報告結果_平滑化版.xlsx', 'visit_rates.xlsx'])
    if rate_file:
        visit_rate_data = load_visit_rates(rate_file)
        print(f'📈 来店結果報告データ読み込み済み: {len(visit_rate_data)}日分')
    else:
        print(f'⚠️  来店結果報告ファイルが見つかりません。固定値で代替します。')

    if args.xlsx:
        print(f'📊 KO XLSXを読み込み中: {args.xlsx}')
        meta, segments, age_segments = parse_ko_xlsx(args.xlsx, visit_rate_data=visit_rate_data)
        print(f'   → 店舗: {meta.get("store")}  送信ID: {meta.get("sendId")}')
        print(f'   → 送信成功: {meta.get("sent")}名  LP閲覧: {meta.get("lpViews")}名  来店: {meta.get("visits")}名')
        print(f'   → 離反セグメント: {len(segments)}  年代セグメント: {len(age_segments)}')
        html = inject_segments(html, segments)
        html = inject_age_segments(html, age_segments)

    # ── 設置台数・キャンペーン種別（インタラクティブ入力 or CLI引数）
    target_file  = find_data_file(script_dir, ['SMSターゲット.xlsx', 'SMS対象.xlsx', 'sms_targets.xlsx'])
    member_file  = find_data_file(script_dir, ['会員データ調査_想定台数100刻み_算出済.xlsx', '会員データ調査_想定人数100刻み_算出済.xlsx', 'member_data.xlsx'])
    campaign_targets = load_campaign_targets(target_file) if target_file else []

    if args.xlsx:
        # 設置台数：未指定ならプロンプト
        if not args.machines:
            try:
                val = input('🏪 店舗の設置台数を入力してください（例: 721）: ').strip()
                args.machines = int(val)
            except (ValueError, EOFError):
                print('⚠️  設置台数が無効のためシミュレーションをスキップします')

        # キャンペーン種別：未指定ならプロンプト
        max_segment   = None
        birthday_mode = False
        if campaign_targets and not args.campaign_type:
            print('\n📋 SMSの内容を選択してください:')
            for i, ct in enumerate(campaign_targets, 1):
                note = '（÷12補正あり）' if ct['birthday'] else ''
                print(f'  {i}. {ct["name"]}　→　〜{ct["max_seg"]}まで対象{note}')
            try:
                choice = int(input('番号を入力: ').strip()) - 1
                if 0 <= choice < len(campaign_targets):
                    selected = campaign_targets[choice]
                    args.campaign_type = selected['name']
                    max_segment        = selected['max_seg']
                    birthday_mode      = selected['birthday']
                    print(f'   → 選択: {args.campaign_type}（最大 {max_segment}）')
            except (ValueError, EOFError):
                print('⚠️  種別未選択のため全セグメントを対象とします')
        elif args.campaign_type:
            matched = next((ct for ct in campaign_targets if ct['name'] == args.campaign_type), None)
            if matched:
                max_segment   = matched['max_seg']
                birthday_mode = matched['birthday']
                print(f'📋 キャンペーン種別: {args.campaign_type}（最大 {max_segment}）')
            else:
                print(f'⚠️  キャンペーン種別 "{args.campaign_type}" が見つかりません。全セグメントを対象とします')

    if args.machines:
        meta['machines'] = args.machines
    if args.machines and args.xlsx and os.path.exists(member_file):
        print(f'🏪 設置台数 {args.machines}台 → 推定会員数を算出中...')
        member_estimates = load_member_estimates(member_file, args.machines)
        html = inject_unsent_segments(html, segments, member_estimates, visit_rate_data,
                                      meta.get('measurementDays', 0),
                                      max_segment=max_segment, birthday_mode=birthday_mode)
        print(f'   → 未送信層シミュレーションデータを注入完了')
        if args.campaign_type:
            html = inject_campaign_meta(html, args.campaign_type, max_segment or '')
    elif args.machines and not member_file:
        print(f'⚠️  会員データ調査ファイルが見つかりません')

    # ── Clarity CSV処理（Attention + Scroll 両対応）
    reach_map = {}
    if args.scroll_csv:
        print(f'📄 スクロールCSVを読み込み中: {args.scroll_csv}')
        reach_map = parse_scroll_csv(args.scroll_csv)
        print(f'   → 到達率データ {len(reach_map)}行')

    if args.csv:
        csv_type = detect_csv_type(args.csv)
        print(f'📄 注目割合CSVを読み込み中: {args.csv}')
        if csv_type in ('attention', 'clarity'):
            clarity_meta, scroll_depths = parse_clarity_csv(args.csv)
            print(f'   → Attentionデータ {len(scroll_depths)}行  総PV: {clarity_meta.get("pageViews")}')
            meta.update({k: v for k, v in clarity_meta.items() if k not in meta})
            if not meta.get('sendId') and clarity_meta.get('url'):
                m = re.search(r'/(\d+)$', clarity_meta['url'])
                if m: meta['sendId'] = m.group(1)
            # Scroll CSVから到達率を注入
            if reach_map:
                for d in scroll_depths:
                    d['reach'] = reach_map.get(d['depth'], None)
            html = inject_scroll_depths(html, scroll_depths)
        else:
            print('   ⚠️  未対応のCSV形式です')

    # ── SMS本文分析
    if args.xlsx and meta.get('smsText'):
        sms_analysis = analyze_sms(meta['smsText'], store_name=meta.get('store', ''))
        html = inject_sms_analysis(html, meta['smsText'], sms_analysis)
        score_label = {'good':'✅ 良好', 'caution':'⚠️ 要確認', 'review':'❗ 要改善'}.get(sms_analysis['score'], '')
        print(f'📝 SMS本文分析: {score_label}（{len(meta["smsText"])}字）')

    # ── KPI・ヘッダー注入（CSV読み込み後にまとめて実行）
    if args.xlsx:
        html = inject_kpi(html, meta)
        html = inject_header(html,
            store    = meta.get('store'),
            campaign = meta.get('campaign'),
            sendid   = meta.get('sendId'),
            dt       = meta.get('datetime'),
            axis     = '来店転換率 × LP支持率',
        )
    elif args.csv:
        html = inject_header(html, sendid=meta.get('sendId'))

    # ── 画像処理
    image_path = args.image
    if not image_path:
        base = os.path.splitext(args.xlsx or args.csv)[0]
        image_path = find_image(base)
        if image_path:
            print(f'🖼  画像を自動検出: {image_path}')
    if image_path:
        print(f'🖼  画像変換中...')
        html = inject_image(html, image_to_base64(image_path))
        print(f'   → 埋め込み完了')
    else:
        print('🖼  画像なし（同名の画像ファイルを置くか --image で指定）')

    # ── 出力
    out_path = os.path.join(script_dir, args.output)
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f'\n✅ レポート生成完了: {out_path}')
    print(f'   ファイルサイズ: {os.path.getsize(out_path) // 1024} KB')


if __name__ == '__main__':
    main()
