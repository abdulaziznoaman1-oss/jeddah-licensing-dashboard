"""Editable Arabic PowerPoint export with Ministry-of-Culture inspired styling."""
from __future__ import annotations
from io import BytesIO
from pathlib import Path
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches,Pt
from pptx.oxml.xmlchemy import OxmlElement
from analytics import MONTHS, month_label, text_clean

NAVY='#0B2230'; INK='#102633'; PLUM='#852555'; TEAL='#4F9F9C'; CORAL='#F2655A'; ORANGE='#E49B3E'; PALE='#F5F6F6'; LIGHT='#EAF0F1'; WHITE='#FFFFFF'; GREY='#667780'
P=Path(__file__).resolve().parent
LOGO=P/'static/assets/moc-logo.png'; COVER=P/'static/assets/jeddah.jpg'
W=13.333; H=7.5

def rgb(h):return RGBColor.from_string(h.lstrip('#'))
def shape(s,x,y,w,h,fill=WHITE,line=None,rounded=False):
    q=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE, Inches(x),Inches(y),Inches(w),Inches(h))
    q.fill.solid();q.fill.fore_color.rgb=rgb(fill)
    if line:q.line.color.rgb=rgb(line)
    else:q.line.fill.background()
    if rounded:
        try:q.adjustments[0]=0.13
        except Exception: pass
    return q

def text(s,x,y,w,h,value,size=15,bold=False,color=INK,align='right',valign='middle',rtl=True):
    box=s.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h))
    tf=box.text_frame;tf.clear();tf.word_wrap=True
    tf.margin_left=Inches(.07);tf.margin_right=Inches(.07);tf.margin_top=Inches(.02);tf.margin_bottom=Inches(.02)
    tf.vertical_anchor={'top':MSO_ANCHOR.TOP,'middle':MSO_ANCHOR.MIDDLE,'bottom':MSO_ANCHOR.BOTTOM}.get(valign,MSO_ANCHOR.MIDDLE)
    for i,line in enumerate(str(value).split('\n')):
        par=tf.paragraphs[0] if i==0 else tf.add_paragraph()
        par.alignment={'right':PP_ALIGN.RIGHT,'center':PP_ALIGN.CENTER,'left':PP_ALIGN.LEFT}[align]
        if rtl:par._p.get_or_add_pPr().set('rtl','1')
        run=par.add_run();run.text=text_clean(line)
        run.font.name='Arial';run.font.size=Pt(size);run.font.bold=bold;run.font.color.rgb=rgb(color)
        run._r.get_or_add_rPr().set('lang','ar-SA' if rtl else 'en-US')
    return box

def base(prs,num,title,subtitle=''):
    s=prs.slides.add_slide(prs.slide_layouts[6]);shape(s,0,0,W,H,WHITE)
    shape(s,0,0,W,.11,NAVY)
    shape(s,11.08,.40,1.88,.56,NAVY,rounded=True)
    text(s,11.17,.48,1.69,.37,'مركز الخدمات البلدية',12,True,WHITE,'center')
    text(s,.72,.48,9.8,.56,title,26,True,PLUM)
    if subtitle:text(s,.72,1.00,11.75,.40,subtitle,12,False,GREY)
    shape(s,.74,1.53,8.0,.045,PLUM);shape(s,8.78,1.53,1.08,.045,CORAL);shape(s,9.92,1.53,1.32,.045,TEAL);shape(s,11.3,1.53,.6,.045,ORANGE)
    shape(s,.72,7.08,11.89,.012,LIGHT)
    text(s,.76,7.10,4,.23,'Confidential - مقيد',8,False,GREY,'left')
    text(s,5.2,7.10,3,.23,'جدة التاريخية',9,False,GREY,'center')
    text(s,11.72,7.10,.82,.23,str(num).zfill(2),9,True,GREY,'right',rtl=False)
    return s

def card(s,x,y,w,h,label,value,sub='',accent=PLUM):
    shape(s,x,y,w,h,WHITE,LIGHT,rounded=True)
    shape(s,x+w-.11,y+.19,.045,h-.38,accent)
    text(s,x+.17,y+.10,w-.40,.34,label,13,True,GREY)
    text(s,x+.17,y+.51,w-.39,.68,value,30,True,accent)
    if sub:text(s,x+.17,y+h-.49,w-.4,.35,sub,10,False,GREY)

def hbar(s,x,y,w,h,items,accent=PLUM,limit=None,labelsize=11):
    if not items:return
    maxv=max(float(it[1] or 0) for it in items) or 1
    n=len(items)
    stride=h/n
    for i,(name,value) in enumerate(items):
        yy=y+i*stride
        lbl=str(name)
        if len(lbl)>50:lbl=lbl[:47]+'...'
        text(s,x+w*.48,yy,w*.50,.24,lbl,labelsize,False,INK)
        barx=x+.54
        barw=w*.43
        shape(s,barx,yy+.06,barw,.19,LIGHT,rounded=True)
        if value:shape(s,barx,yy+.06,max(.03,barw*float(value)/maxv),.19,accent,rounded=True)
        text(s,x+.07,yy,.45,.27,str(value),10,True,accent,'left',rtl=False)

def draw_stacks(s,series,x,y,w,h,colors,keys,with_counts=True):
    n=max(1,len(series));gap=.20; bw=min(.78,(w-.3)/n-gap)
    maxval=max(1,max((v['total'] for v in series),default=1))
    axisy=y+h-.36
    shape(s,x,axisy,w,.015,LIGHT)
    for i,m in enumerate(series):
        xi=x+.13+i*w/n+(w/n-bw)/2
        baseline=axisy
        for j,k in enumerate(keys):
            v=m.get(k,0);bh=(h-.82)*float(v)/maxval
            if bh>0:
                baseline-=bh;shape(s,xi,baseline,bw,bh,colors[j])
        if with_counts:text(s,xi-.06,baseline-.37,bw+.12,.29,str(m['total']),11,True,PLUM,'center',rtl=False)
        text(s,xi-.20,axisy+.08,bw+.4,.29,m['label'].split()[0],10,False,GREY,'center')

def simple_line(s,series,key,x,y,w,h,color=TEAL,suffix='%'):
    valid=[(i,m.get(key)) for i,m in enumerate(series) if m.get(key) is not None]
    if not valid:return
    values=[v for _,v in valid]; ymin=min(0,min(values)-10) if key=='growth' else 0; ymax=max(100,max(values)+10) if suffix=='%' and key!='growth' else max(values+[1])*1.2
    if ymin==ymax:ymax=ymin+1
    mid=y+h*.5
    for k in range(5):
        yy=y+(h-.68)*k/4;shape(s,x,yy,w,.012,LIGHT)
    pts=[]
    for i,v in valid:
        xi=x+.22+(w-.42)*(i/(max(1,len(series)-1)))
        yi=y+(h-.69)*(1-(v-ymin)/(ymax-ymin))
        pts.append((xi,yi,v))
    for a,b in zip(pts,pts[1:]):
        from math import atan2,sqrt,degrees
        dx=b[0]-a[0];dy=b[1]-a[1]
        ln=s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,Inches(a[0]),Inches(a[1]),Inches(b[0]),Inches(b[1]));ln.line.color.rgb=rgb(color);ln.line.width=Pt(2.5)
    for idx,(xx,yy,v) in enumerate(pts):
        shape(s,xx-.056,yy-.056,.112,.112,color,rounded=True)
        text(s,xx-.49,yy-.39,.98,.30,f'{v:.1f}{suffix}' if isinstance(v,float) else f'{v}{suffix}',10,True,color,'center',rtl=False)
    for i,m in enumerate(series):
        xx=x+.20+(w-.42)*(i/(max(1,len(series)-1)))
        text(s,xx-.50,y+h-.44,1.0,.25,m['label'].split()[0],9,False,GREY,'center')

def prose(s,items,starty=2.05,linegap=.78,fontsize=17):
    for i,item in enumerate(items):
        y=starty+i*linegap
        shape(s,12.15,y+.12,.10,.10,[PLUM,TEAL,ORANGE,CORAL][i%4],rounded=True)
        text(s,.97,y,11.05,.57,item,fontsize,False,INK)

def format_pct(v):return 'غير متاح' if v is None else f'{v:.1f}%'
def fmt(v,unit=''):return 'غير متاح' if v is None else f'{v:,.1f}{unit}' if isinstance(v,float) else f'{v:,}{unit}'

def export_ppt(summary):
    k=summary['kpis'];months=summary['monthly'];top=summary['top_activities'];dept=summary['departments'];period=f"{month_label(summary['meta']['start'])} - {month_label(summary['meta']['end'])}"
    prs=Presentation();prs.slide_width=Inches(W);prs.slide_height=Inches(H)
    # Cover
    s=prs.slides.add_slide(prs.slide_layouts[6]);s.shapes.add_picture(str(COVER),0,0,width=Inches(W),height=Inches(H))
    shape(s,.42,.39,3.10,1.03,WHITE,rounded=True);s.shapes.add_picture(str(LOGO),Inches(.70),Inches(.52),width=Inches(2.05))
    shape(s,.62,3.23,7.3,2.62,WHITE,rounded=True)
    text(s,1.04,3.45,6.25,.54,'مركز الخدمات البلدية',27,True,NAVY)
    text(s,1.04,4.11,6.28,.58,'تقرير أثر منصة التراخيص',29,True,PLUM)
    text(s,1.04,4.82,6.12,.39,period,16,True,INK)
    text(s,1.04,5.30,6.12,.30,'تحليل الأداء التشغيلي والالتزام والمواءمة',12,False,GREY)
    shape(s,0,7.14,W,.36,NAVY)
    # 2 purpose
    s=base(prs,2,'لماذا منصة متابعة التراخيص؟','تحويل معاملات التراخيص إلى قرارات تشغيلية مبنية على البيانات')
    prose(s,['توحيد بيانات طلبات الإصدار والتجديد والأنشطة في مصدر واحد.','متابعة حالة الطلب والجهة المسؤولة وتاريخ الإدخال والاعتماد.','قياس التزام SLA ونسبة المواءمة بصورة شهرية قابلة للمقارنة.','إبراز الطلبات المفتوحة والمتأخرة لتوجيه المتابعة التشغيلية.'])
    # 3 kpis
    s=base(prs,3,'الملخص التنفيذي','مؤشرات الفترة المختارة')
    for i,(lab,val,sub,ac) in enumerate([
        ('إجمالي الطلبات',fmt(k['total']),'طلبات مسجلة',PLUM),('طلبات تجديد',fmt(k['renew']),'جميع حالات الطلب',TEAL),
        ('طلبات إصدار',fmt(k['issue']),'جميع حالات الطلب',ORANGE),('متوسط شهري',fmt(k['avg_monthly']),'طلب في الشهر',CORAL),
        ('التزام SLA',format_pct(k['sla']),f"من {k['sla_base']} طلب منته ومصنف",TEAL),
        ('المواءمة العامة',format_pct(k['alignment']),'من جميع الطلبات',PLUM),('طلبات مفتوحة',fmt(k['open']),'حتى تاريخ استخراج البيانات',CORAL),
        ('متوسط أيام العمل',fmt(k['avg_days']),'للطلبات المنتهية',ORANGE)]):
        col=i%4;row=i//4;card(s,.78+col*3.17,1.91+row*2.18,3.00,1.97,lab,val,sub,ac)
    # 4 flow
    s=base(prs,4,'كيف نظمت المنصة مسار الطلب؟','تتبع مراحل المعالجة وفقا للحقول المتوفرة في قاعدة الطلبات')
    flow=[('1','تسجيل الطلب','نوع الخدمة والنشاط وتاريخ الإدخال'),('2','توجيه المعالجة','الجهة الحالية وحالة الطلب'),('3','قياس المدة','أيام العمل وتصنيف المواءمة'),('4','قرار ومتابعة','الموافقة أو الرفض والطلبات المفتوحة')]
    for i,(n,h,b) in enumerate(flow):
        x=9.58-i*3.07
        shape(s,x,2.33,2.77,2.55,WHITE,LIGHT,rounded=True)
        shape(s,x+.98,2.55,.74,.74,[PLUM,TEAL,ORANGE,CORAL][i],rounded=True)
        text(s,x+1.01,2.67,.67,.40,n,22,True,WHITE,'center',rtl=False)
        text(s,x+.12,3.53,2.53,.42,h,17,True,PLUM,'center')
        text(s,x+.15,4.04,2.48,.63,b,12,False,INK,'center')
    text(s,1.05,5.70,11.15,.67,'الأثر المرصود: وجود سجل موحد يتيح رصد حجم الطلبات ومسارها ومدة معالجتها ومواطن التأخير.',16,True,INK)
    # 5 monthly volume
    s=base(prs,5,'حجم الطلبات الشهري','إجمالي طلبات التراخيص حسب شهر إدخال الطلب')
    draw_stacks(s,months,.98,2.04,11.2,4.24,[PLUM],['total'])
    # 6 mix
    s=base(prs,6,'تجديد مقابل إصدار','توزيع أنواع الخدمة لكل شهر، لجميع حالات الطلب')
    draw_stacks(s,months,.98,2.10,9.2,4.13,[PLUM,TEAL],['renew','issue'])
    shape(s,10.35,2.22,2.06,1.78,PALE,rounded=True);text(s,10.45,2.52,1.81,.48,'تجديد',15,True,PLUM,'center');text(s,10.45,3.07,1.81,.58,str(k['renew']),27,True,PLUM,'center',rtl=False)
    shape(s,10.35,4.18,2.06,1.78,PALE,rounded=True);text(s,10.45,4.45,1.81,.48,'إصدار',15,True,TEAL,'center');text(s,10.45,4.99,1.81,.58,str(k['issue']),27,True,TEAL,'center',rtl=False)
    # 7 SLA vs alignment
    s=base(prs,7,'الالتزام بالـ SLA والمواءمة','التمييز بين الطلبات المنتهية وجميع الطلبات المسجلة')
    simple_line(s,months,'sla',.82,2.06,7.48,3.93,TEAL)
    simple_line(s,months,'alignment',8.85,2.06,3.65,3.93,PLUM)
    text(s,.85,6.23,7.16,.41,'SLA: الطلبات المنتهية والمصنفة فقط',13,True,TEAL)
    text(s,8.76,6.23,3.83,.41,'المواءمة: جميع الطلبات',13,True,PLUM)
    # 8 days
    s=base(prs,8,'كفاءة المعالجة','متوسط أيام العمل في الطلبات ذات الحالة المنتهية فقط')
    simple_line(s,months,'avg_days',1.05,2.11,10.92,4.11,ORANGE,' يوم')
    # 9 state
    s=base(prs,9,'حالات الطلبات الشهرية','عدد الموافقات والرفض والطلبات المفتوحة')
    draw_stacks(s,months,.92,2.10,9.35,4.12,[TEAL,PLUM,CORAL],['approved','rejected','open'])
    for i,(name,val,col) in enumerate([('موافق عليها',k['approved'],TEAL),('مرفوضة',k['rejected'],PLUM),('مفتوحة',k['open'],CORAL)]):
        card(s,10.48,2.07+i*1.47,1.9,1.35,name,str(val),'',col)
    # 10 growth
    s=base(prs,10,'نمو حجم الطلبات','نسبة التغير الشهري في عدد المعاملات مقارنة بالشهر السابق')
    simple_line(s,months,'growth',1.05,2.16,10.85,3.97,PLUM)
    text(s,2.20,6.12,8.9,.48,'ملاحظة: لا يحسب نمو الشهر الأول لعدم وجود شهر سابق داخل الفترة المختارة.',12,False,GREY)
    # 11 service
    s=base(prs,11,'المتوسط ومزيج الخدمات','حصة التجديد والإصدار ومتوسط حجم المعاملات الشهري')
    card(s,1.1,2.15,3.5,2.25,'متوسط الطلبات الشهري',fmt(k['avg_monthly']),'طلب',PLUM)
    card(s,4.95,2.15,3.5,2.25,'حصة طلبات التجديد',format_pct(k['renew']/k['total']*100 if k['total'] else None),'من إجمالي الطلبات',TEAL)
    card(s,8.79,2.15,3.5,2.25,'حصة طلبات الإصدار',format_pct(k['issue']/k['total']*100 if k['total'] else None),'من إجمالي الطلبات',ORANGE)
    prose(s,['إجمالي الطلبات يشمل الحالات المفتوحة والمرفوضة.','طلبات الإصدار والتجديد لا تعني بالضرورة أن الرخص صدرت أو جددت بنجاح.'],starty=5.00,linegap=.66,fontsize=14)
    # 12 entity
    s=base(prs,12,'تصنيف المنشآت','توزيع الطلبات حسب نوع المنشأة الوارد في البيانات')
    for i,(lab,v,accent) in enumerate([('تجاري',k['commercial'],PLUM),('صحي',k['health'],TEAL),('غير محدد',k['unspecified'],ORANGE)]):
        card(s,1.15+i*4.17,2.24,3.80,2.57,lab,str(v),f"{format_pct(v/k['total']*100 if k['total'] else None)} من الطلبات",accent)
    # 13 top activities
    s=base(prs,13,'الأنشطة الأكثر طلبا','أكثر 10 أنشطة تقدمت بطلبات تراخيص')
    hbar(s,1.12,1.94,11.05,4.95,[(it['name'],it['count']) for it in top],PLUM,labelsize=11)
    # 14 open
    s=base(prs,14,'الطلبات المفتوحة والمتأخرة','متابعة المخاطر التشغيلية في الفترة المختارة')
    card(s,.93,2.06,3.64,2.10,'الطلبات المفتوحة',str(k['open']),'غير محسوبة في SLA للطلبات المنتهية',CORAL)
    card(s,4.81,2.06,3.64,2.10,'المتأخرة المفتوحة',str(k['overdue_open']),'وفقا لتصنيف الحالة في الملف',PLUM)
    card(s,8.69,2.06,3.64,2.10,'المواءمة العامة',format_pct(k['alignment']),'تتضمن الطلبات المفتوحة',TEAL)
    text(s,1.0,4.97,11.27,.93,'ضرورة عرض الرصيد المفتوح بشكل مستقل حتى لا تخفي نسب الالتزام الجيدة للطلبات المنتهية أي تراكم في المعاملات قيد المعالجة.',17,False,INK)
    # 15 departments
    s=base(prs,15,'توزيع العمل حسب الجهة','عدد الطلبات المسجلة لدى كل جهة معالجة')
    hbar(s,1.18,2.24,10.9,3.77,[(it['name'],it['count']) for it in dept[:7]],TEAL,labelsize=14)
    # 16 impact
    s=base(prs,16,'الأثر التشغيلي المثبت','ما يمكن إثباته من بيانات التراخيص دون افتراض خط أساس غير متوفر')
    prose(s,[f"إتاحة متابعة موحدة لـ {k['total']} طلب خلال الفترة المختارة.",
        f"توثيق {k['renew']} طلب تجديد و{k['issue']} طلب إصدار عبر حقول مشتركة.",
        f"قياس SLA بنسبة {format_pct(k['sla'])} والمواءمة العامة بنسبة {format_pct(k['alignment'])}.",
        f"كشف {k['open']} طلب مفتوح يتطلب متابعة مستقلة عن الطلبات المنتهية.",
        'مقارنة الأداء شهريا وتسليط الضوء على الأنشطة الأكثر طلبا.'],2.01,.82,16)
    # 17 recommendations
    s=base(prs,17,'التوصيات التشغيلية','أولويات لرفع كفاءة المتابعة وجودة البيانات')
    prose(s,['متابعة الطلبات المفتوحة يوميا وتخصيص مالك واضح لكل معاملة متأخرة.','تثبيت تعريف SLA والمواءمة ومراجعة طريقة تصنيف الحالات.','توحيد بيانات النشاط والجهة وتفادي تصنيف المنشأة على أنها غير محددة.','تحديث الداشبورد عند كل رفع جديد واعتماد مراجعة شهرية للمؤشرات.','إضافة رقم طلب فريد وبيانات تاريخية قبل تشغيل المنصة لقياس الأثر السببي.'],1.95,.80,16)
    # 18 methodology
    s=base(prs,18,'منهجية الاحتساب وحدود النتائج','التعاريف المستخدمة في تحليل البيانات')
    prose(s,['الطلبات: عدد الصفوف الفريدة وفق تاريخ إدخال الطلب ونوعه والنشاط ونوع المنشأة.','الالتزام بـ SLA: عدد الطلبات المنتهية المصنفة موائم ÷ الطلبات المنتهية ذات تصنيف.','المواءمة العامة: عدد الطلبات المصنفة موائم ÷ جميع الطلبات في الفترة.','معدل النمو: (طلبات الشهر - طلبات الشهر السابق) ÷ طلبات الشهر السابق.','متوسط أيام العمل: المتوسط للطلبات المنتهية ذات أيام عمل صالحة.','بعض الأشهر قد تكون جزئية. لا توجد مقارنة سببية بدون بيانات ما قبل المنصة.'],1.86,.74,14)
    out=BytesIO();prs.save(out);return out.getvalue()
