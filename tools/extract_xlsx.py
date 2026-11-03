import json,base64,email,sys,hashlib,re
from email import policy
src,outdir=sys.argv[1],sys.argv[2]
raw=json.load(open(src))['raw']
raw=base64.urlsafe_b64decode(raw+'='*(-len(raw)%4))
msg=email.message_from_bytes(raw,policy=policy.default)
n=0
for part in msg.walk():
    fn=part.get_filename()
    if fn and fn.lower().endswith(('.xlsx','.xls')):
        data=part.get_payload(decode=True)
        safe=re.sub(r'[^\w.\-가-힣]','_',fn)
        open(f'{outdir}/{safe}','wb').write(data); n+=1
        print(json.dumps({'file':f'{outdir}/{safe}','name':fn,'bytes':len(data),'sha256':hashlib.sha256(data).hexdigest()},ensure_ascii=False))
if not n: print('NO_XLSX')
