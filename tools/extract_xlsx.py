# 사용법: python3 -I tools/extract_xlsx.py <Gmail RAW 결과 JSON> <저장 폴더> [확장자들, 기본 xlsx,xls / 예: pdf]
import json,base64,email,sys,hashlib,re
from email import policy
src,outdir=sys.argv[1],sys.argv[2]
exts=tuple('.'+e.strip('.').lower() for e in (sys.argv[3].split(',') if len(sys.argv)>3 else ['xlsx','xls']))  # 예: pdf
raw=json.load(open(src))['raw']
raw=base64.urlsafe_b64decode(raw+'='*(-len(raw)%4))
msg=email.message_from_bytes(raw,policy=policy.default)
n=0
for part in msg.walk():
    fn=part.get_filename()
    if fn and fn.lower().endswith(exts):
        data=part.get_payload(decode=True)
        safe=re.sub(r'[^\w.\-가-힣]','_',fn)
        open(f'{outdir}/{safe}','wb').write(data); n+=1
        print(json.dumps({'file':f'{outdir}/{safe}','name':fn,'bytes':len(data),'sha256':hashlib.sha256(data).hexdigest()},ensure_ascii=False))
if not n: print('NO_MATCH')
