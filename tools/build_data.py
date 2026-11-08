#!/usr/bin/env python3
"""비표 / 정기권 엑셀 → 암호화된 data.enc

admin.html 과 같은 규칙으로 엑셀을 읽고, 사이트의 keys.json(공개 열쇠)으로 잠급니다.
비밀번호는 필요 없습니다.

사용법:
  python3 tools/build_data.py --pass 비표.xlsx [--season 정기권.xlsx] [--prev 이전비표.xlsx]
                              --keys keys.json --out data.enc [--state tools/state.json] [--force]

출력(표준출력, JSON 한 줄): {"changed": bool, "summary": {...}, "diff": {...}, "warnings": [...]}
내용이 이전과 같으면(state.json 의 content_sha 와 같으면) data.enc 를 쓰지 않고 changed=false.
"""
import argparse, base64, datetime, hashlib, json, os, re, sys

import openpyxl
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import padding
from cryptography.hazmat.primitives.ciphers.aead import AESGCM

BLD = {'ONE': '1IFC', 'TWO': '2IFC', 'THREE': '3IFC', '1IFC': '1IFC', '2IFC': '2IFC', '3IFC': '3IFC',
       'ONE IFC': '1IFC', 'TWO IFC': '2IFC', 'THREE IFC': '3IFC'}


def s(v):
    if v is None:
        return ''
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    if isinstance(v, (datetime.datetime, datetime.date)):
        return ds(v)
    return str(v).strip()


def ds(v):
    if isinstance(v, (datetime.datetime, datetime.date)):
        return f'{v.year}-{v.month:02d}-{v.day:02d}'
    return s(v)


def norm(p):
    return re.sub(r'[\s\-·.]', '', str(p or ''))


def bld(v):
    u = s(v).upper()
    return BLD.get(u) or BLD.get(re.sub(r'\s', '', u)) or s(v)


def read_sheet(ws):
    rows = []
    for r in ws.iter_rows(max_col=41, max_row=5001, values_only=True):
        if any(c is not None and c != '' for c in r):
            rows.append(list(r))
    hi = next((i for i, r in enumerate(rows[:8]) if any(re.sub(r'\s', '', s(c)) == '차량번호' for c in r)), -1)
    if hi < 0:
        return None
    H = {}
    for i, c in enumerate(rows[hi]):
        k = re.sub(r'\s', '', s(c))
        if k and k not in H:
            H[k] = i

    def col(r, *names):
        for nm in names:
            idx = H.get(nm)
            if idx is None:
                idx = next((H[h] for h in H if nm in h), None)
            if idx is not None:
                return r[idx] if idx < len(r) else None
        return None

    out = []
    for r in rows[hi + 1:]:
        x = dict(gu=s(col(r, '구분')), dt=ds(col(r, '일자', '신청일자', '등록일')), c=s(col(r, '회사명', '회사', '입주사')),
                 ow=s(col(r, '차주(직급)', '차주')), no=s(col(r, '관리번호', '비표번호')), m=s(col(r, '차종')),
                 p=s(col(r, '차량번호')), d=s(col(r, '차량기사명', '기사명', '기사')), dp=s(col(r, '연락처')),
                 mg=s(col(r, '담당자')), ext=s(col(r, '내선번호', '내선')), mp=s(col(r, '핸드폰번호', '핸드폰', '휴대폰')),
                 st=s(col(r, '등록여부', '상태')), o=s(col(r, '비고')), exp=ds(col(r, '만료일', '종료일', '만료')))
        if x['p']:
            out.append(x)
    return out


def parse(pass_path, season_path=None):
    PASS, VOID, SEASON, SKIP, used = [], [], [], [], []
    if pass_path:
        wb = openpyxl.load_workbook(pass_path, data_only=True, read_only=True)
        for name in wb.sheetnames:
            b = BLD.get(name.upper().strip())
            is_void = bool(re.search('해지|분실', name))
            if not b and not is_void:
                continue
            rows = read_sheet(wb[name])
            if not rows:
                continue
            used.append(name)
            for r in rows:
                n = norm(r['p'])
                base = dict(p=r['p'], n=n, c=r['c'], m=r['m'], d=r['d'], dp=r['dp'], mg=r['mg'], ext=r['ext'],
                            mp=r['mp'], no=r['no'], dt=r['dt'], o=r['o'])
                if b:
                    if r['st'] != '등록':
                        continue
                    if re.fullmatch(r'\d{1,5}', n):
                        SKIP.append(f"{b} {r['c']} {r['p']} (차량번호 없는 공용비표)")
                        continue
                    PASS.append({'b': b, **base, 'w': []})
                else:
                    VOID.append({'b': bld(r['gu']), **base, 'st': r['st'] or '해지'})
    if season_path:
        wb = openpyxl.load_workbook(season_path, data_only=True, read_only=True)
        for name in wb.sheetnames:
            rows = read_sheet(wb[name])
            if not rows:
                continue
            used.append('정기권:' + name)
            for r in rows:
                if re.search('해지|만료|분실|취소|대기', r['st']):
                    continue
                SEASON.append(dict(b=bld(r['gu']) or BLD.get(name.upper().strip()) or '', p=r['p'], n=norm(r['p']), c=r['c'],
                                   m=r['m'], d=r['d'] or r['ow'], dp=r['dp'], mg=r['mg'], ext=r['ext'], mp=r['mp'],
                                   no=r['no'], dt=r['dt'], exp=r['exp'], o=r['o'], w=[]))
    WARN = []

    def flag(r, msg):
        if msg not in r['w']:
            r['w'].append(msg)

    byN = {}
    for r in PASS:
        byN.setdefault(r['n'], []).append(r)
    for n, lst in byN.items():
        if len(lst) > 1:
            msg = f'같은 차량이 {len(lst)}번 등록됨 (' + ', '.join(f"{r['b']} {r['c']} {r['no']}번" for r in lst) + ')'
            for r in lst:
                flag(r, msg)
            WARN.append([n, msg])
    for r in PASS:
        for v in VOID:
            if v['n'] == r['n'] and v['dt'] and r['dt'] and (v['dt'] > r['dt'] or (v['dt'] == r['dt'] and not re.search('분실|훼손', v['st']))):
                msg = f"해지·분실 명단에도 있음 ({v['dt']} {v['st']}{', ' + v['o'] if v['o'] else ''})"
                flag(r, msg)
                WARN.append([r['p'], msg])
    byNo = {}
    for r in PASS:
        if r['no']:
            byNo.setdefault(r['b'] + '|' + r['no'], []).append(r)
    for k, lst in byNo.items():
        plates = list(dict.fromkeys(r['n'] for r in lst))
        if len(plates) < 2:
            continue
        b, no = k.split('|', 1)
        msg = f"{b} 관리번호 {no}번을 다른 차량과 같이 씀 ({', '.join(plates)})"
        for r in lst:
            flag(r, f'관리번호 {no}번을 다른 차량과 같이 씀')
        WARN.append([', '.join(plates), msg])
    for r in PASS:
        if not re.fullmatch(r'[가-힣]{0,2}\d{2,3}[가-힣]\d{4}', norm(r['p'])):
            flag(r, '번호판 형식 확인 필요')
            WARN.append([r['p'], f"번호판 형식이 이상함: '{r['p']}' ({r['c']})"])
    warns = list(dict.fromkeys(w[1] for w in WARN))
    return dict(pass_=PASS, void=VOID, season=SEASON, warnings=warns, skip=SKIP, used=used)


def b64(b):
    return base64.b64encode(b).decode()


def encrypt(obj, pub_b64):
    pub = serialization.load_der_public_key(base64.b64decode(pub_b64))
    raw, iv = os.urandom(32), os.urandom(12)
    ct = AESGCM(raw).encrypt(iv, json.dumps(obj, ensure_ascii=False, separators=(',', ':')).encode(), None)
    ek = pub.encrypt(raw, padding.OAEP(mgf=padding.MGF1(hashes.SHA256()), algorithm=hashes.SHA256(), label=None))
    return {'v': 2, 'updated': obj['updated'], 'ek': b64(ek), 'iv': b64(iv), 'ct': b64(ct)}


def plate_index(lst):
    out = {}
    for r in lst:
        out.setdefault(r['n'], r)
    return out


def diff(prev, cur):
    a, b = plate_index(prev['pass_']), plate_index(cur['pass_'])
    added = [f"{r['b']} {r['p']} {r['c']}" for n, r in b.items() if n not in a]
    removed = [f"{r['b']} {r['p']} {r['c']}" for n, r in a.items() if n not in b]
    changed = []
    for n, r in b.items():
        if n in a:
            o = a[n]
            ch = [f'{lbl} {o[k] or "-"}→{r[k] or "-"}' for k, lbl in (('b', '건물'), ('c', '회사'), ('d', '기사'), ('dp', '기사연락처'), ('mg', '담당자'), ('no', '관리번호')) if o[k] != r[k]]
            if ch:
                changed.append(f"{r['p']}: " + ', '.join(ch))
    va, vb = plate_index(prev['void']), plate_index(cur['void'])
    new_void = [f"{r['b']} {r['p']} {r['c']} ({r['st']})" for n, r in vb.items() if n not in va]
    return dict(added=added, removed=removed, changed=changed, new_void=new_void)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--pass', dest='pass_path', required=True)
    ap.add_argument('--season')
    ap.add_argument('--prev')
    ap.add_argument('--keys', required=True)
    ap.add_argument('--out', required=True)
    ap.add_argument('--state')
    ap.add_argument('--force', action='store_true')
    ap.add_argument('--source', default='', help='state.json 에 남길 출처 메모 (예: 메일 ID, 첨부 파일명)')
    ap.add_argument('--dump-json', help='(검증용) 암호화 전 내용을 이 경로에 저장')
    a = ap.parse_args()

    cur = parse(a.pass_path, a.season)
    content = {'pass': cur['pass_'], 'void': cur['void'], 'season': cur['season']}
    sha = hashlib.sha256(json.dumps(content, ensure_ascii=False, sort_keys=True).encode()).hexdigest()
    state = {}
    if a.state and os.path.exists(a.state):
        state = json.load(open(a.state))
    cnt = lambda lst, b=None: len({r['n'] for r in lst if not b or r['b'] == b})
    summary = {'비표': {b: cnt(cur['pass_'], b) for b in ('1IFC', '2IFC', '3IFC')}, '비표합계': cnt(cur['pass_']),
               '정기권합계': cnt(cur['season']), '해지분실기록': len(cur['void']), '읽은시트': cur['used'], '제외': cur['skip']}
    result = {'changed': a.force or state.get('content_sha') != sha, 'content_sha': sha, 'summary': summary,
              'warnings': cur['warnings']}
    if a.prev:
        result['diff'] = diff(parse(a.prev), cur)
    if result['changed']:
        kst = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=9)))
        content['updated'] = kst.strftime('%Y-%m-%d %H:%M')
        if a.dump_json:
            json.dump(content, open(a.dump_json, 'w'), ensure_ascii=False)
        keys = json.load(open(a.keys))
        json.dump(encrypt(content, keys['pub']), open(a.out, 'w'), separators=(',', ':'))
        result['updated'] = content['updated']
        if a.state:
            json.dump({'content_sha': sha, 'updated': content['updated'], 'source': a.source}, open(a.state, 'w'), ensure_ascii=False, indent=1)
    print(json.dumps(result, ensure_ascii=False))


if __name__ == '__main__':
    main()
