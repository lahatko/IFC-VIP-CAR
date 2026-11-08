#!/usr/bin/env python3
"""지하 정기권 목록(주차 시스템 내보내기 CSV 또는 엑셀) → 암호화된 season.enc

admin.html 의 정기권 처리와 같은 규칙입니다. 사이트의 keys.json(공개 열쇠)으로 잠그므로 비밀번호는 필요 없습니다.
비표 명단(data.enc)과 별도 파일이라, 매일 비표 자동 업데이트가 정기권을 지우지 않습니다.

사용법:
  일반 정기권:            python3 tools/build_season.py --file export_....csv --keys keys.json --out season.enc
  지하4층 로딩덕 정기권:  python3 tools/build_season.py --file export_....csv --zone B4 --keys keys.json --out dock.enc
"""
import argparse, base64, csv, datetime, gzip, io, json, os, re

from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import padding
from cryptography.hazmat.primitives.ciphers.aead import AESGCM

JUNK = {'null', '-', '.', 'x', 'X', 'q', '0'}


def clean(v):
    v = '' if v is None else str(v).strip()
    if isinstance(v, str) and re.fullmatch(r'\d+\.0', v):
        v = v[:-2]
    return '' if v in JUNK else v


def norm(p):
    return re.sub(r'[\s\-·.]', '', str(p or ''))


def building(company, code):
    m = re.search(r'([123])\s*IFC', company or '', re.I)
    if m:
        return m.group(1) + 'IFC'
    m = re.match(r'F([123])', code or '')
    return m.group(1) + 'IFC' if m else ''


def read_rows(path):
    if path.lower().endswith('.csv'):
        raw = open(path, 'rb').read()
        try:
            text = raw.decode('utf-8-sig')
        except UnicodeDecodeError:
            text = raw.decode('cp949', errors='replace')
        rows = list(csv.reader(io.StringIO(text)))
    else:
        import openpyxl
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        rows = []
        for ws in wb:
            rows = [['' if c is None else (c.strftime('%Y-%m-%d %H:%M:%S') if isinstance(c, datetime.datetime) else str(c)) for c in r]
                    for r in ws.iter_rows(values_only=True)]
            if any('차량번호' in ''.join(r) for r in rows[:5]):
                break
    hi = next(i for i, r in enumerate(rows[:5]) if any(re.sub(r'\s', '', c) == '차량번호' for c in r))
    H = {}
    for i, c in enumerate(rows[hi]):
        k = re.sub(r'\s', '', c)
        if k and k not in H:
            H[k] = i
    return H, rows[hi + 1:]


def parse(path, zone=''):
    H, rows = read_rows(path)

    def col(r, *names):
        for nm in names:
            i = H.get(nm)
            if i is None:
                i = next((H[h] for h in H if nm in h), None)
            if i is not None:
                return r[i] if i < len(r) else ''
        return ''

    out, seen = [], set()
    for r in rows:
        p = clean(col(r, '차량번호'))
        if not p:
            continue
        c, code = clean(col(r, '업체명')), clean(col(r, '업체코드'))
        x = {'b': building(c, code), 'p': p, 'n': norm(p), 'c': c,
             'prod': clean(col(r, '상품명', '싱픔명')), 'pt': clean(col(r, '상품유형')),
             'fr': clean(col(r, '사용시작일시', '시작')), 'to': clean(col(r, '사용종료일시', '종료')),
             'use': 1 if clean(col(r, '사용여부')) == '사용' else 0,
             'm': clean(col(r, '모델')) or clean(col(r, '차종')), 'd': clean(col(r, '이름')),
             'dp': clean(col(r, '연락처')), 'room': clean(col(r, '주소')), 'o': clean(col(r, '비고')), 'zone': zone}
        x = {k: v for k, v in x.items() if v != '' and not (k == 'pt' and v == '기간상품')}   # 빈 칸·기본값은 빼서 크기 줄임
        key = json.dumps(x, ensure_ascii=False, sort_keys=True)
        if key in seen:
            continue
        seen.add(key)
        out.append(x)
    return out


def encrypt(obj, pub_b64):
    pub = serialization.load_der_public_key(base64.b64decode(pub_b64))
    raw, iv = os.urandom(32), os.urandom(12)
    body = gzip.compress(json.dumps(obj, ensure_ascii=False, separators=(',', ':')).encode(), mtime=0)   # 압축 후 잠금
    ct = AESGCM(raw).encrypt(iv, body, None)
    ek = pub.encrypt(raw, padding.OAEP(mgf=padding.MGF1(hashes.SHA256()), algorithm=hashes.SHA256(), label=None))
    b64 = lambda b: base64.b64encode(b).decode()
    return {'v': 2, 'z': 'gzip', 'updated': obj['updated'], 'ek': b64(ek), 'iv': b64(iv), 'ct': b64(ct)}


def exported_from_name(name):
    m = re.search(r'(20\d{2})(\d{2})(\d{2})_(\d{2})(\d{2})', os.path.basename(name))
    return f'{m.group(1)}-{m.group(2)}-{m.group(3)} {m.group(4)}:{m.group(5)}' if m else ''


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--file', required=True)
    ap.add_argument('--keys', required=True)
    ap.add_argument('--out', required=True)
    ap.add_argument('--exported', default='', help='목록을 내보낸 시각 (없으면 파일명 export_YYYYMMDD_HHMMSS 에서 읽음)')
    ap.add_argument('--zone', default='', help="구역 표시 (지하4층 로딩덕 = B4)")
    ap.add_argument('--dump-json')
    a = ap.parse_args()
    season = parse(a.file, a.zone)
    kst = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=9)))
    obj = {'updated': kst.strftime('%Y-%m-%d %H:%M'), 'exported': a.exported or exported_from_name(a.file), 'zone': a.zone, 'season': season}
    if a.dump_json:
        json.dump(obj, open(a.dump_json, 'w'), ensure_ascii=False)
    keys = json.load(open(a.keys))
    json.dump(encrypt(obj, keys['pub']), open(a.out, 'w'), separators=(',', ':'))
    now = kst.strftime('%Y-%m-%d %H:%M:%S')
    valid = {x['n'] for x in season if x.get('use') and x.get('fr', '') <= now <= x.get('to', '')}
    by_b = {b: len({x['n'] for x in season if x['n'] in valid and x.get('b', '') == b}) for b in ('1IFC', '2IFC', '3IFC', '')}
    print(json.dumps({'rows': len(season), 'plates': len({x['n'] for x in season}), 'valid_now': len(valid),
                      'valid_by_building': by_b, 'exported': obj['exported']}, ensure_ascii=False))


if __name__ == '__main__':
    main()
