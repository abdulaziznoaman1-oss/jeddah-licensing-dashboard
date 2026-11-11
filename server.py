"""Deployable, self-contained licensing dashboard HTTP server.
Uses Python standard library + python-pptx. Keeps uploaded data in SQLite.
"""
from __future__ import annotations
import base64
import csv
import datetime as dt
import hashlib
import hmac
import io
import json
import logging
from collections import defaultdict, deque
import mimetypes
import os
import secrets
import sqlite3
import sys
import time
import zipfile
from email.parser import BytesParser
from email.policy import default
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

from analytics import COLS, clean_record, compute, read_rows, record_key
from ppt_report import export_ppt

ROOT=Path(__file__).resolve().parent
DATA_DIR=Path(os.getenv('DATA_DIR',str(ROOT/'data')))
DATA_DIR.mkdir(parents=True,exist_ok=True)
DB=str(DATA_DIR/'licensing.sqlite')
SEED=ROOT/'data'/'example.xlsx'
VIEW_PASSWORD=os.getenv('VIEW_PASSWORD','').strip()
ADMIN_PASSWORD=os.getenv('ADMIN_PASSWORD','').strip()
IS_PROD=os.getenv('APP_ENV','').lower()=='production'
SECRET=os.getenv('SECRET_KEY','').encode('utf-8') or secrets.token_bytes(32)
MAX_UPLOAD=28*1024*1024
FAILED_LOGINS=defaultdict(deque)
SQL_SCHEMA='''
CREATE TABLE IF NOT EXISTS records (record_key TEXT PRIMARY KEY, payload TEXT NOT NULL, updated_at TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS imports (id INTEGER PRIMARY KEY AUTOINCREMENT, filename TEXT, mode TEXT, uploaded_at TEXT, incoming INTEGER, total INTEGER);
'''

def db_connect():
    conn=sqlite3.connect(DB,timeout=30)
    conn.execute('PRAGMA journal_mode=WAL')
    conn.execute('PRAGMA busy_timeout=30000')
    return conn

def init_db():
    with db_connect() as conn:conn.executescript(SQL_SCHEMA)
    with db_connect() as conn:
        existing=conn.execute('SELECT COUNT(*) FROM records').fetchone()[0]
    if existing==0 and SEED.exists() and not IS_PROD:
        rows=read_rows(SEED.read_bytes(),SEED.name)
        put_rows(rows,'استبدال','example.xlsx')

def put_rows(rows,mode,filename):
    now=dt.datetime.now(dt.timezone.utc).isoformat()
    mapping={record_key(clean_record(r)):json.dumps(clean_record(r),ensure_ascii=False) for r in rows}
    with db_connect() as conn:
        conn.execute('BEGIN IMMEDIATE')
        if mode=='استبدال':conn.execute('DELETE FROM records')
        conn.executemany('INSERT INTO records(record_key,payload,updated_at) VALUES(?,?,?) '
                         'ON CONFLICT(record_key) DO UPDATE SET payload=excluded.payload,updated_at=excluded.updated_at',
                         [(key,val,now) for key,val in mapping.items()])
        total=conn.execute('SELECT COUNT(*) FROM records').fetchone()[0]
        conn.execute('INSERT INTO imports(filename,mode,uploaded_at,incoming,total) VALUES(?,?,?,?,?)',
                     (filename,mode,now,len(rows),total))
        conn.commit()
    return {'received':len(rows),'unique_upload':len(mapping),'total':total,'mode':mode}

def read_saved():
    with db_connect() as conn:
        payloads=conn.execute('SELECT payload FROM records').fetchall()
        last=conn.execute('SELECT filename,mode,uploaded_at,incoming,total FROM imports ORDER BY id DESC LIMIT 1').fetchone()
    return [json.loads(v[0]) for v in payloads], dict(zip(('filename','mode','uploaded_at','incoming','total'),last)) if last else None

def signed_session(role,csrf):
    data=json.dumps({'role':role,'exp':int(time.time())+60*60*12,'csrf':csrf},separators=(',',':')).encode()
    payload=base64.urlsafe_b64encode(data).decode().rstrip('=')
    sig=hmac.new(SECRET,payload.encode(),hashlib.sha256).hexdigest()
    return payload+'.'+sig

def parse_session(raw):
    if not raw:return None
    try:
        payload,signature=raw.rsplit('.',1)
        if not hmac.compare_digest(signature,hmac.new(SECRET,payload.encode(),hashlib.sha256).hexdigest()):return None
        item=json.loads(base64.urlsafe_b64decode(payload+'='*(-len(payload)%4)))
        return item if item['exp']>=time.time() else None
    except (ValueError,KeyError,TypeError):return None

def safe_compare(a,b):return hmac.compare_digest(str(a).encode(),str(b).encode())

class Handler(BaseHTTPRequestHandler):
    server_version='LicensingDashboard/1.0'
    def log_message(self, fmt, *args):
        logging.info('%s %s',self.address_string(),fmt%args)

    def respond(self,data,status=200,ctype='application/json; charset=utf-8',headers=None):
        data=data if isinstance(data,bytes) else data.encode('utf-8') if isinstance(data,str) else json.dumps(data,ensure_ascii=False).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type',ctype)
        self.send_header('Content-Length',str(len(data)))
        self.send_header('X-Content-Type-Options','nosniff')
        self.send_header('Referrer-Policy','same-origin')
        self.send_header('X-Frame-Options','DENY')
        self.send_header('Cache-Control','no-store' if ctype.startswith('application/json') or ctype.startswith('text/html') else 'private,max-age=300')
        if ctype.startswith('text/html'):self.send_header('Content-Security-Policy',"default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' data:; connect-src 'self'; object-src 'none'; frame-ancestors 'none'")
        for name,val in (headers or {}).items():self.send_header(name,val)
        self.end_headers()
        try:self.wfile.write(data)
        except BrokenPipeError:pass

    def get_cookies(self):
        raw=self.headers.get('Cookie','')
        return dict(x.strip().split('=',1) for x in raw.split(';') if '=' in x)

    def auth(self):
        if IS_PROD and (not VIEW_PASSWORD or not ADMIN_PASSWORD):return None
        if not VIEW_PASSWORD and not ADMIN_PASSWORD:return {'role':'admin','csrf':'local'}
        return parse_session(self.get_cookies().get('session'))

    def authorized(self,write=False):
        sess=self.auth()
        if sess is None:
            self.respond({'error':'يجب تسجيل الدخول أولا'},401)
            return None
        if write and sess['role']!='admin':
            self.respond({'error':'رفع البيانات متاح للمسؤول فقط'},403)
            return None
        return sess

    def csrf_verified(self,sess):
        if not safe_compare(self.headers.get('X-CSRF-Token',''),sess['csrf']):
            self.respond({'error':'انتهت صلاحية الجلسة، أعد تسجيل الدخول'},403)
            return False
        return True

    def do_GET(self):
      try:self._get()
      except Exception as ex:
        logging.exception('GET error')
        self.respond({'error':'حدث خطأ أثناء المعالجة'},500)

    def _get(self):
        req=urlparse(self.path);path=req.path
        if path=='/healthz':return self.respond({'status':'ok'})
        if path=='/favicon.ico':return self.respond(b'',204,'image/x-icon')
        if path=='/':
            return self.respond((ROOT/'static'/'index.html').read_bytes(),ctype='text/html; charset=utf-8')
        if path=='/static/plotly.min.js':
            return self.respond((ROOT/'static/plotly.min.js').read_bytes(),ctype='text/javascript; charset=utf-8')
        if path in ['/static/style.css','/static/main.js','/static/assets/moc-logo.png','/static/assets/jeddah.jpg']:
            file=ROOT/path.lstrip('/')
            return self.respond(file.read_bytes(),ctype=mimetypes.guess_type(str(file))[0] or 'application/octet-stream')
        sess=self.authorized()
        if not sess:return
        if path=='/api/bootstrap':
            data,last=read_saved()
            return self.respond({'ok':True,'role':sess['role'],'csrf':sess['csrf'],'count':len(data),'last_import':last})
        if path in ('/api/dashboard','/api/export.pptx','/api/export.csv'):
            query={k:values[0] for k,values in parse_qs(req.query).items() if values}
            data,last=read_saved();result=compute(data,query)
            if result.get('empty'):return self.respond(result)
            if path=='/api/dashboard':
                result['last_import']=last
                # raw records are deliberately served only by explicit CSV export
                result.pop('records',None)
                return self.respond(result)
            if path=='/api/export.pptx':
                result.pop('records',None)
                b=export_ppt(result)
                return self.respond(b,ctype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
                    headers={'Content-Disposition':'attachment; filename="licensing-impact-report.pptx"'})
            buffer=io.StringIO();out=csv.DictWriter(buffer,fieldnames=COLS);out.writeheader()
            for record in result['records']:
                out.writerow({k:("'"+str(v) if str(v).startswith(('=','+','-','@')) else v) for k,v in record.items()})
            return self.respond('\ufeff'+buffer.getvalue(),ctype='text/csv; charset=utf-8',
                headers={'Content-Disposition':'attachment; filename="licensing-data.csv"'})
        return self.respond({'error':'المسار غير متاح'},404)

    def do_POST(self):
      try:self._post()
      except Exception as ex:
        logging.exception('POST error')
        self.respond({'error':'تعذر تنفيذ العملية. تحقق من صيغة الملف.'},500)

    def _post(self):
        req=urlparse(self.path)
        size=int(self.headers.get('Content-Length','0'))
        if size<=0 or size>MAX_UPLOAD+100000:
            return self.respond({'error':'البيانات فارغة أو تتجاوز 28 ميجابايت'},413)
        if req.path=='/api/login':
            if IS_PROD and (not VIEW_PASSWORD or not ADMIN_PASSWORD):
                return self.respond({'error':'يجب ضبط VIEW_PASSWORD و ADMIN_PASSWORD على الخادم'},503)
            try:body=json.loads(self.rfile.read(size))
            except ValueError:return self.respond({'error':'طلب غير صالح'},400)
            pwd=body.get('password','')
            source=self.client_address[0]
            now=time.time()
            attempts=FAILED_LOGINS[source]
            while attempts and attempts[0]<now-900:attempts.popleft()
            if len(attempts)>=8:return self.respond({'error':'تم تجاوز عدد محاولات الدخول. حاول لاحقا.'},429)
            if ADMIN_PASSWORD and safe_compare(pwd,ADMIN_PASSWORD):role='admin'
            elif VIEW_PASSWORD and safe_compare(pwd,VIEW_PASSWORD):role='viewer'
            elif not VIEW_PASSWORD and not ADMIN_PASSWORD:role='admin'
            else:
                attempts.append(now)
                return self.respond({'error':'كلمة المرور غير صحيحة'},401)
            FAILED_LOGINS.pop(source,None)
            csrf=secrets.token_urlsafe(24);token=signed_session(role,csrf)
            secure='; Secure' if (IS_PROD or self.headers.get('X-Forwarded-Proto','')=='https') else ''
            return self.respond({'ok':True,'role':role},headers={'Set-Cookie':f'session={token}; HttpOnly; SameSite=Lax; Path=/; Max-Age=43200{secure}'})
        if req.path=='/api/logout':
            return self.respond({'ok':True},headers={'Set-Cookie':'session=; HttpOnly; SameSite=Lax; Path=/; Max-Age=0'})
        if req.path=='/api/upload':
            sess=self.authorized(write=True)
            if not sess:return
            if not self.csrf_verified(sess):return
            content_type=self.headers.get('Content-Type','')
            if 'multipart/form-data' not in content_type:
                return self.respond({'error':'يجب اختيار ملف'},400)
            raw=self.rfile.read(size)
            full=b'Content-Type: '+content_type.encode()+b'\r\nMIME-Version: 1.0\r\n\r\n'+raw
            message=BytesParser(policy=default).parsebytes(full)
            uploaded=None; filename='';mode='استبدال'
            for part in message.iter_parts():
                if part.get_content_disposition()!='form-data':continue
                name=part.get_param('name',header='content-disposition')
                if name=='file':
                    uploaded=part.get_payload(decode=True)
                    filename=Path(part.get_filename() or 'upload.xlsx').name
                if name=='mode':mode=(part.get_payload(decode=True) or b'').decode('utf-8','replace')
            if uploaded is None or len(uploaded)>MAX_UPLOAD:
                return self.respond({'error':'الملف مفقود أو كبير جدا'},400)
            if mode not in ('دمج','استبدال'):return self.respond({'error':'وضع التحديث غير صالح'},400)
            try:rows=read_rows(uploaded,filename)
            except Exception as ex:
                return self.respond({'error':str(ex)[:300]},400)
            answer=put_rows(rows,mode,filename)
            return self.respond({'ok':True,**answer})
        return self.respond({'error':'المسار غير متاح'},404)

if __name__=='__main__':
    import zipfile
    logging.basicConfig(level=logging.INFO,format='%(asctime)s %(levelname)s %(message)s')
    init_db()
    port=int(os.environ.get('PORT','8501'))
    server=ThreadingHTTPServer(('0.0.0.0',port),Handler)
    print(f'Licensing dashboard: http://localhost:{port}',flush=True)
    server.serve_forever()
