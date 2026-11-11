"""Analytics for licensing requests. No pandas/openpyxl dependency.
All ratios are explicitly defined to avoid treating open requests as completed.
"""
from __future__ import annotations
import calendar
import csv
import datetime as dt
import hashlib
import io
import json
import re
import unicodedata
import xml.etree.ElementTree as ET
import zipfile
from collections import Counter, defaultdict
from statistics import mean

AR_DIACRITICS = re.compile(r'[\u064b-\u065f\u0670\u06d6-\u06ed]')
MONTHS = ['يناير','فبراير','مارس','أبريل','مايو','يونيو','يوليو','أغسطس','سبتمبر','أكتوبر','نوفمبر','ديسمبر']
COLS = ['نوع الخدمة','نوع المنشأة','حالة الطلب','الجهة الحالية للمعالجة','النشاط الرئيسي','تاريخ ادخال الطلب','تاريخ الاعتماد النهائي','عدد أيام العمل','تصنيف الحالة','رقم الطلب']
ALIASES = {
 'نوع الخدمة': ['الخدمة','نوع الخدمة لخدمة الرخص التجارية','نوع طلب الرخصة'],
 'نوع المنشأة': ['تصنيف المنشأة','نوع النشاط'],
 'حالة الطلب':['حالة المعاملة','الحالة'],
 'الجهة الحالية للمعالجة':['الجهة المعالجة','القسم الحالي'],
 'النشاط الرئيسي':['النشاط','النشاط الاقتصادي','اسم النشاط'],
 'تاريخ ادخال الطلب':['تاريخ إدخال الطلب','تاريخ استقبال الطلب','تاريخ إستقبال المعاملة','تاريخ الطلب','تاريخ إنشاء الطلب'],
 'تاريخ الاعتماد النهائي':['تاريخ الإعتماد النهائي','تاريخ القرار النهائي','تاريخ الاعتماد','تاريخ الاغلاق'],
 'عدد أيام العمل':['مدة المعالجة','أيام العمل','عدد ايام العمل'],
 'تصنيف الحالة':['الموائمة','المواءمة','حالة المواءمة','التزام SLA'],
 'رقم الطلب':['رقم المعاملة','رقم الطلب من منصة بلدي','Request ID','Order ID','معرف الطلب']
}
XML = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'


def text_clean(value):
    return AR_DIACRITICS.sub('', str(value or '').replace('\u200f','').replace('\u200e','')).strip()


def norm(value):
    s=text_clean(value)
    s=s.replace('إ','ا').replace('أ','ا').replace('آ','ا').replace('ى','ي').replace('ة','ه')
    return re.sub(r'\s+',' ',s).strip().lower()

MAPPING = {norm(c):c for c in COLS}
for canon, names in ALIASES.items():
 for n in names: MAPPING[norm(n)] = canon


def get_col_name(value):
    return MAPPING.get(norm(value))


def parse_time(v):
    if isinstance(v,dt.datetime): return v
    if isinstance(v,dt.date): return dt.datetime.combine(v,dt.time.min)
    v=text_clean(v)
    if not v or norm(v) in {'جاري العمل','قيد العمل','غير متوفر','لا يوجد'}: return None
    if re.fullmatch(r'\d+(?:\.\d+)?',v):
        f=float(v)
        if 20000<f<100000:
            return dt.datetime(1899,12,30)+dt.timedelta(days=f)
    for fmt in ('%d/%m/%Y, %H:%M:%S','%d/%m/%Y %H:%M:%S','%d/%m/%Y','%Y-%m-%d %H:%M:%S','%Y-%m-%d','%m/%d/%Y %H:%M:%S'):
        try: return dt.datetime.strptime(v,fmt)
        except ValueError: pass
    try:return dt.datetime.fromisoformat(v.replace('Z','+00:00')).replace(tzinfo=None)
    except ValueError:return None


def _parse_xlsx(content):
    with zipfile.ZipFile(io.BytesIO(content)) as z:
        members=z.infolist()
        if sum(m.file_size for m in members)>300*1024*1024 or any(m.file_size>180*1024*1024 for m in members):
            raise ValueError('الملف المضغوط يحتوي على بيانات أكبر من الحد الآمن')
        paths=z.namelist()
        strings=[]
        if 'xl/sharedStrings.xml' in paths:
            ss=ET.fromstring(z.read('xl/sharedStrings.xml'))
            strings=[''.join(e.itertext()) for e in ss.findall(f'{XML}si')]
        sheet_files=sorted((name for name in paths if re.fullmatch(r'xl/worksheets/sheet\d+\.xml',name)),key=lambda n:int(re.search(r'(\d+)\.xml',n).group(1)))
        for sheet_file in sheet_files:
            root=ET.fromstring(z.read(sheet_file))
            sheet_rows=[]
            for node in root.findall(f'.//{XML}sheetData/{XML}row'):
                cells={}
                for c in node.findall(f'{XML}c'):
                    ref=c.get('r',''); cidx=0
                    for ch in re.match(r'[A-Z]+',ref).group(0): cidx=cidx*26+ord(ch)-64
                    cidx-=1
                    value=c.find(f'{XML}v')
                    if value is not None and value.text is not None:
                        if c.get('t')=='s':
                            idx=int(value.text); val=strings[idx] if 0<=idx<len(strings) else ''
                        else:val=value.text
                    else:
                        inline=c.find(f'{XML}is')
                        val=''.join(inline.itertext()) if inline is not None else ''
                    cells[cidx]=text_clean(val)
                sheet_rows.append(cells)
                if len(sheet_rows)>175000: raise ValueError('الملف يتجاوز الحد الأقصى 175 ألف صف')
            if not sheet_rows: continue
            header_idx=None
            for i,head in enumerate(sheet_rows[:12]):
                cols={get_col_name(v) for v in head.values()}
                if 'نوع الخدمة' in cols and 'تاريخ ادخال الطلب' in cols:
                    header_idx=i;break
            if header_idx is None:continue
            mapping={pos:get_col_name(label) for pos,label in sheet_rows[header_idx].items() if get_col_name(label)}
            result=[]
            for item in sheet_rows[header_idx+1:]:
                row={c:'' for c in COLS}
                for pos,col in mapping.items():row[col]=item.get(pos,'')
                if any(row.values()):result.append(row)
            return result
    raise ValueError('لم أجد شيت يحتوي على نوع الخدمة وتاريخ إدخال الطلب. راجع أسماء الأعمدة.')


def read_rows(content, filename):
    if len(content)>28*1024*1024: raise ValueError('الملف كبير جدا: الحد الأقصى 28 ميجابايت')
    filename=filename.lower()
    if filename.endswith('.xlsx'): rows=_parse_xlsx(content)
    elif filename.endswith('.csv'):
        try: decoded=content.decode('utf-8-sig')
        except UnicodeDecodeError: decoded=content.decode('cp1256')
        reader=csv.DictReader(io.StringIO(decoded))
        rows=[]
        for row in reader:
            item={c:'' for c in COLS}
            for key,value in row.items():
                col=get_col_name(key)
                if col:item[col]=text_clean(value)
            rows.append(item)
            if len(rows)>175000:raise ValueError('الملف يتجاوز الحد الأقصى 175 ألف صف')
    else:raise ValueError('ارفع ملف Excel بصيغة XLSX أو CSV')
    if not rows:raise ValueError('الملف لا يحتوي على صفوف بيانات')
    valid=[r for r in rows if parse_time(r.get('تاريخ ادخال الطلب')) and r.get('نوع الخدمة')]
    if len(valid)!=len(rows):
        raise ValueError(f'يوجد {len(rows)-len(valid)} صف دون نوع خدمة أو بتاريخ إدخال غير صالح. صحح البيانات ثم أعد الرفع لتجنب إسقاط أي طلب.')
    return valid


def record_key(r):
    # Future exports can supply request IDs. Otherwise use stable record fields.
    if text_clean(r.get('رقم الطلب')):
        return hashlib.sha256(('req:'+norm(r['رقم الطلب'])).encode('utf-8')).hexdigest()
    parts=[norm(r.get(k)) for k in ('تاريخ ادخال الطلب','نوع الخدمة','النشاط الرئيسي','نوع المنشأة')]
    return hashlib.sha256('\x1e'.join(parts).encode('utf-8')).hexdigest()


def clean_record(r):
    return {col:text_clean(r.get(col,'')) for col in COLS}


def enrich(r):
    date=parse_time(r.get('تاريخ ادخال الطلب'))
    if date is None:return None
    status=norm(r.get('حالة الطلب'))
    has_final_date=bool(parse_time(r.get('تاريخ الاعتماد النهائي')))
    open_= ('جاري' in status or 'قيد' in status or 'انتظار' in status) or (not has_final_date and not any(x in status for x in ('موافقه','رفض','مكتمل','منتهي')))
    approved=('موافقه' in status or 'اعتماد' in status) and not open_
    rejected='رفض' in status and not open_
    compliant=norm(r.get('تصنيف الحالة')) in ('موائم','ملتزم','مطابق')
    unclassified=not r.get('تصنيف الحالة')
    try: days=float(str(r.get('عدد أيام العمل','')).replace(',','.'))
    except ValueError: days=None
    if days is not None and (days<0 or days>3650):days=None
    service=norm(r.get('نوع الخدمة'))
    svc='تجديد' if 'تجديد' in service else 'إصدار' if 'اصدار' in service else 'أخرى'
    ent=norm(r.get('نوع المنشأة'))
    entity='تجاري' if 'تجاري' in ent else 'صحي' if 'صحي' in ent else 'غير محدد'
    return dict(r, _date=date, month=date.strftime('%Y-%m'), service=svc,
                entity=entity, open=open_, approved=approved, rejected=rejected,
                compliant=compliant, classified=not unclassified, days=days)


def month_sequence(a,b):
    ay,am=map(int,a.split('-'));by,bm=map(int,b.split('-'))
    seq=[]
    while (ay,am)<=(by,bm) and len(seq)<120:
        seq.append(f'{ay:04d}-{am:02d}')
        am+=1
        if am==13:ay+=1;am=1
    return seq


def month_label(s):
    y,m=map(int,s.split('-'))
    return f'{MONTHS[m-1]} {y}'


def pct(n,d):return round(n/d*100,1) if d else None


def metrics(rows):
    n=len(rows);closed=[r for r in rows if not r['open']]; c=len(closed)
    good=sum(r['compliant'] for r in rows)
    cls_closed=[r for r in closed if r['classified']]
    dvals=[r['days'] for r in closed if r['days'] is not None]
    return dict(total=n, renew=sum(r['service']=='تجديد' for r in rows),issue=sum(r['service']=='إصدار' for r in rows),
       commercial=sum(r['entity']=='تجاري' for r in rows),health=sum(r['entity']=='صحي' for r in rows),
       unspecified=sum(r['entity']=='غير محدد' for r in rows),
       approved=sum(r['approved'] for r in rows),rejected=sum(r['rejected'] for r in rows),open=sum(r['open'] for r in rows),
       overdue_open=sum(r['open'] and r['classified'] and not r['compliant'] for r in rows),
       closed=c,compliant=good,
       sla=pct(sum(r['compliant'] for r in cls_closed),len(cls_closed)),
       sla_base=len(cls_closed),
       alignment=pct(good,n),
       avg_days=round(mean(dvals),2) if dvals else None,
       approval_rate=pct(sum(r['approved'] for r in rows),c))


def compute(all_raw, filters=None):
    filters=filters or {}
    enriched=[e for r in all_raw if (e:=enrich(r))]
    if not enriched:
        return {'empty':True,'months':[],'monthly':[],'kpis':metrics([]),'insights':[],'top_activities':[],'meta':{}}
    months_all=sorted({r['month'] for r in enriched})
    start=filters.get('start') if filters.get('start') in month_sequence(months_all[0],months_all[-1]) else (months_all[-4] if len(months_all)>=4 else months_all[0])
    end=filters.get('end') if filters.get('end') in month_sequence(months_all[0],months_all[-1]) else months_all[-1]
    if start>end:start,end=end,start
    types=filters.get('service','')
    entities=filters.get('entity','')
    state=filters.get('state','')
    activity=filters.get('activity','')
    dept=filters.get('department','')
    selected=[r for r in enriched if start<=r['month']<=end and (not types or r['service']==types) and
             (not entities or r['entity']==entities) and (not activity or r.get('النشاط الرئيسي')==activity) and
             (not dept or r.get('الجهة الحالية للمعالجة')==dept) and
             (not state or ('مفتوحة' if r['open'] else 'موافق عليها' if r['approved'] else 'مرفوضة' if r['rejected'] else 'مغلقة أخرى')==state)]
    sequence=month_sequence(start,end)
    series=[]
    last=None
    for m in sequence:
        mm=metrics([r for r in selected if r['month']==m])
        mm['month']=m; mm['label']=month_label(m)
        mm['growth']=round((mm['total']/last-1)*100,1) if last else None
        last=mm['total']
        series.append(mm)
    total=metrics(selected)
    total['avg_monthly']=round(total['total']/len(sequence),1) if sequence else None
    total['growth']=series[-1]['growth'] if series else None
    total['period_start']=start; total['period_end']=end
    top=Counter(r.get('النشاط الرئيسي') or 'غير محدد' for r in selected)
    top_activities=[{'name':name,'count':count} for name,count in top.most_common(10)]
    deptcount=Counter(r.get('الجهة الحالية للمعالجة') or 'غير محدد' for r in selected)
    open_dept=Counter(r.get('الجهة الحالية للمعالجة') or 'غير محدد' for r in selected if r['open'])
    insights=[]
    if len(series)>1:
        hi=max(series,key=lambda x:x['total']);lo=min(series,key=lambda x:x['total'])
        insights.append(f"أعلى حجم طلبات كان في {hi['label']} بعدد {hi['total']} طلب، وأدنى حجم في {lo['label']} بعدد {lo['total']} طلب.")
        s0=series[0].get('sla'); s1=series[-1].get('sla')
        if s0 is not None and s1 is not None:
            diff=round(s1-s0,1)
            insights.append(f"تغير التزام SLA للطلبات المصنفة والمنتهية بمقدار {abs(diff)} نقطة مئوية {'ارتفاعا' if diff>=0 else 'انخفاضا'} بين أول شهر وآخر شهر.")
        growth=series[-1]['growth']
        if growth is not None:
            insights.append(f"تغير حجم الطلبات في آخر شهر بنسبة {abs(growth)}% {'ارتفاعا' if growth>=0 else 'انخفاضا'} عن الشهر السابق.")
    if total['open']:
        insights.append(f"يوجد {total['open']} طلب مفتوح ضمن الفترة المختارة، منها {total['overdue_open']} مصنف متأخر. هذه الطلبات تدخل في المواءمة العامة ولا تدخل في مؤشر SLA للطلبات المنتهية.")
    latest=max(r['_date'] for r in enriched)
    last_month_day=calendar.monthrange(latest.year,latest.month)[1]
    if latest.day<last_month_day:
        insights.append(f'آخر تاريخ إدخال متاح في البيانات هو {latest.strftime("%d/%m/%Y")}. قد لا يكون آخر شهر مكتملا؛ لذا يفسر النمو الشهري بحذر.')
    insights.append('هذه المؤشرات تقيس انتظام التسجيل والمعالجة كما يظهر في الملف، ولا تثبت وحدها الأثر السببي للمنصة دون خط أساس قبل التشغيل.')
    return dict(empty=False,kpis=total,months=sequence,monthly=series,
       meta=dict(all_months=months_all,month_options=[{'value':m,'label':month_label(m)} for m in month_sequence(months_all[0],months_all[-1])],
                 first=months_all[0],last=months_all[-1],start=start,end=end,
                 activities=[v for v,c in Counter(r.get('النشاط الرئيسي') for r in enriched).most_common() if v],
                 departments=[v for v,c in Counter(r.get('الجهة الحالية للمعالجة') for r in enriched).most_common() if v],
                 source_records=len(enriched),filtered_records=len(selected),latest_input_date=latest.strftime('%d/%m/%Y')),
       top_activities=top_activities,
       departments=[{'name':name,'count':count,'open':open_dept[name]} for name,count in deptcount.most_common()],
       insights=insights,
       records=[{k:r.get(k,'') for k in COLS} for r in selected],
       filters=filters)
