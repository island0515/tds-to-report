import json,os,re,shutil,subprocess,tempfile,threading,urllib.parse,zipfile
from http.server import BaseHTTPRequestHandler,ThreadingHTTPServer
from pathlib import Path
HOST=os.getenv("CONVERTER_HOST","0.0.0.0");PORT=int(os.getenv("CONVERTER_PORT","8080"));TOKEN=os.getenv("CONVERTER_TOKEN","");TIMEOUT=int(os.getenv("CONVERSION_TIMEOUT_SECONDS","120"));MAX_MB=int(os.getenv("MAX_UPLOAD_MB","80"));LOCK=threading.Semaphore(1);STATE_LOCK=threading.Lock();ACTIVE=False;WAITING=0
def state():
 with STATE_LOCK:return {"active":ACTIVE,"waiting":WAITING}
def soffice():
 x=shutil.which("soffice") or shutil.which("libreoffice")
 if not x:raise RuntimeError("容器中未找到 soffice")
 return x
def convert(content,name):
 global ACTIVE,WAITING
 with STATE_LOCK:WAITING+=1
 acquired=False
 try:
  LOCK.acquire();acquired=True
  with STATE_LOCK:WAITING=max(0,WAITING-1);ACTIVE=True
  with tempfile.TemporaryDirectory(prefix="lo-") as td:
   root=Path(td);src=root/(re.sub(r'[^A-Za-z0-9._-]+','_',Path(name).stem)+'.doc');out=root/'out';profile=root/'profile';out.mkdir();profile.mkdir();src.write_bytes(content)
   cmd=[soffice(),"--headless","--nologo","--nodefault","--nolockcheck","--norestore",f"-env:UserInstallation={profile.resolve().as_uri()}","--convert-to","docx:Office Open XML Text","--outdir",str(out),str(src)]
   result=subprocess.run(cmd,stdout=subprocess.PIPE,stderr=subprocess.PIPE,text=True,timeout=TIMEOUT)
   files=list(out.glob('*.docx'))
   if result.returncode or len(files)!=1:raise RuntimeError(result.stderr.strip() or result.stdout.strip() or '转换失败')
   with zipfile.ZipFile(files[0]) as z:
    if 'word/document.xml' not in z.namelist() or z.testzip():raise RuntimeError('转换结果不是有效 DOCX')
   return files[0].read_bytes(),files[0].name
 finally:
  with STATE_LOCK:
   if not acquired:WAITING=max(0,WAITING-1)
   ACTIVE=False
  if acquired:LOCK.release()
class H(BaseHTTPRequestHandler):
 def log_message(self,f,*a):print('[converter]',f%a,flush=True)
 def js(self,c,p):
  b=json.dumps(p,ensure_ascii=False).encode();self.send_response(c);self.send_header('Content-Type','application/json; charset=utf-8');self.send_header('Content-Length',str(len(b)));self.end_headers();self.wfile.write(b)
 def do_GET(self):
  if self.path!='/health':return self.js(404,{"message":"Not found"})
  s=state()
  try:self.js(200,{"status":"ok","soffice":soffice(),**s,"concurrency":1})
  except Exception as e:self.js(503,{"status":"error","message":str(e)})
 def do_POST(self):
  if self.path!='/convert':return self.js(404,{"message":"Not found"})
  if TOKEN and self.headers.get('X-Converter-Token','')!=TOKEN:return self.js(401,{"message":"Unauthorized"})
  try:
   n=int(self.headers.get('Content-Length','0'))
   if n<=0 or n>MAX_MB*1024*1024:raise RuntimeError('请求为空或文件过大')
   data,name=convert(self.rfile.read(n),urllib.parse.unquote(self.headers.get('X-Filename','input.doc')))
   self.send_response(200);self.send_header('Content-Type','application/vnd.openxmlformats-officedocument.wordprocessingml.document');self.send_header('Content-Length',str(len(data)));self.end_headers();self.wfile.write(data)
  except Exception as e:self.js(422,{"message":str(e)})
if __name__=='__main__':print(f'Converter API {HOST}:{PORT}',flush=True);ThreadingHTTPServer((HOST,PORT),H).serve_forever()
