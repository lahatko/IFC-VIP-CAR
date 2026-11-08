#!/usr/bin/env python3
"""IFC 카페테리아 주간식단표 PDF → menu.json + menu.png

사용법:
  python3 tools/build_menu.py --pdf 주간메뉴.pdf --json menu.json --png menu.png [--source "gmail:<id> ..."]

출력(표준출력, JSON 한 줄): {"ok": bool, "week": "11.02~11.08", "days": 7, "problem": "..."}
"""
import argparse, datetime, json, os, re, subprocess, sys, tempfile

import pdfplumber

MEALS = ['아침', '점심', '저녁', '야식']


def clean_lines(cell):
    out = []
    for ln in (cell or '').split('\n'):
        ln = ln.strip()
        if not ln:
            continue
        if ln.startswith('&') and out:          # '치킨가라아게' + '&양념치킨소스' → 한 줄로
            out[-1] += ln
        else:
            out.append(ln)
    return out


def parse(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        tables = page.extract_tables()
        text = page.extract_text() or ''
    if not tables:
        raise ValueError('표를 찾지 못함')
    rows = max(tables, key=len)
    hi = next(i for i, r in enumerate(rows) if any(re.search(r'\d{4}-\d{2}-\d{2}', c or '') for c in r))
    header = rows[hi]
    # 날짜 칸 → 요일 / 날짜가 없는 칸(None)은 바로 앞 요일의 B코너로 붙임
    colday, days, cur = {}, [], None
    for ci, c in enumerate(header):
        m = re.search(r'([월화수목금토일])\s*\((\d{4}-\d{2}-\d{2})\)', c or '')
        if m:
            cur = len(days)
            days.append({'dow': m.group(1), 'date': m.group(2), 'meals': {}})
            colday[ci] = cur
        elif c is None and cur is not None:
            colday[ci] = cur
    body = rows[hi + 1:]
    is_label = lambda r: len(r) > 1 and len((r[1] or '').strip()) <= 8 and any(k in (r[1] or '') for k in MEALS)   # 아래 안내문(운영시간 등) 제외
    labeled = [i for i, r in enumerate(body) if is_label(r)]
    for n, li in enumerate(labeled):
        meal = next(k for k in MEALS if k in (body[li][1] or ''))
        nxt_main = labeled[n + 1] - 1 if n + 1 < len(labeled) else len(body)
        main_row = body[li - 1] if li > 0 and not (body[li - 1][1] or '').strip() else None
        extra_rows = [r for r in body[li + 1:nxt_main] if not is_label(r) and len((r[1] or '').strip()) <= 8]
        for d in days:
            d['meals'][meal] = {'main': [], 'sides': []}
        for ci, di in colday.items():
            m = days[di]['meals'][meal]
            if main_row and ci < len(main_row):
                for x in clean_lines(main_row[ci]):
                    if x not in m['main']:
                        m['main'].append(x)
            for r in [body[li]] + extra_rows:
                if ci < len(r):
                    for x in clean_lines(r[ci]):
                        if x not in m['sides'] and x not in m['main']:
                            m['sides'].append(x)
        for d in days:
            m = d['meals'][meal]
            if m['sides'] == ['미운영'] or (not m['main'] and '미운영' in m['sides']):
                m.update(main=[], sides=[], closed=True)
    hours = ''
    mh = re.search(r'운영시간\s*-?\s*(.+)', text)
    if mh:
        hours = mh.group(1).strip()
    title = next((ln.strip() for ln in text.split('\n') if '식단' in ln), 'IFC 카페테리아 주간식단차림표')
    title = re.sub(r'^\S\s+', '', title) if len(title.split()[0]) == 1 else title
    return {'title': title, 'days': days, 'hours': hours}


def render_png(pdf_path, out_png):
    with tempfile.TemporaryDirectory() as td:
        subprocess.run(['pdftoppm', '-r', '150', '-png', '-singlefile', pdf_path, os.path.join(td, 'p')], check=True)
        os.replace(os.path.join(td, 'p.png'), out_png)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--pdf', required=True)
    ap.add_argument('--json', required=True)
    ap.add_argument('--png', required=True)
    ap.add_argument('--source', default='')
    a = ap.parse_args()
    try:
        m = parse(a.pdf)
    except Exception as e:
        print(json.dumps({'ok': False, 'problem': f'식단표 PDF를 읽지 못함: {e}'}, ensure_ascii=False))
        sys.exit(0)
    if len(m['days']) < 5:
        print(json.dumps({'ok': False, 'problem': f"요일 칸이 {len(m['days'])}개뿐임 (양식 변경 의심)"}, ensure_ascii=False))
        sys.exit(0)
    d0, d1 = m['days'][0]['date'], m['days'][-1]['date']
    m['week'] = f'{d0[5:7]}.{d0[8:]}~{d1[5:7]}.{d1[8:]}'
    kst = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=9)))
    m['updated'] = kst.strftime('%Y-%m-%d %H:%M')
    m['source'] = a.source
    render_png(a.pdf, a.png)
    json.dump(m, open(a.json, 'w'), ensure_ascii=False, indent=1)
    print(json.dumps({'ok': True, 'week': m['week'], 'days': len(m['days'])}, ensure_ascii=False))


if __name__ == '__main__':
    main()
