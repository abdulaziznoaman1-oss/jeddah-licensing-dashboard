'use strict';
const qs=s=>document.querySelector(s);
const appState={role:'viewer',csrf:'',data:null,boot:false};
const colors={navy:'#0b2230',plum:'#862656',teal:'#4b9997',coral:'#ec7064',orange:'#de9d41',grey:'#8298a2'};
const formatN=n=>n==null?'غير متاح':Number(n).toLocaleString('en-US',{maximumFractionDigits:1});
const formatP=n=>n==null?'—':Number(n).toLocaleString('en-US',{maximumFractionDigits:1})+'%';
const pctV=n=>n==null?'—':formatP(n);
function toast(msg){const e=qs('#toast');e.textContent=msg;e.classList.remove('hidden');clearTimeout(toast.t);toast.t=setTimeout(()=>e.classList.add('hidden'),4500)}
function error(msg){const e=qs('#alert');e.textContent=msg;e.classList.remove('hidden')}
async function api(path,opts={}){
 const resp=await fetch(path,{credentials:'same-origin',...opts});
 let data;const typ=resp.headers.get('Content-Type')||'';
 try{data=typ.includes('json')?await resp.json():await resp.text()}catch(_){data={error:'تعذر قراءة استجابة الخادم'}}
 if(!resp.ok)throw Error((data&&data.error)||'تعذر تنفيذ الطلب');
 return data;
}
function visible(name){qs('#loading').classList.add('hidden');qs('#loginScreen').classList.toggle('hidden',name!=='login');qs('#application').classList.toggle('hidden',name!=='app')}
const filmap={start:'#filterStart',end:'#filterEnd',service:'#filterService',entity:'#filterEntity',state:'#filterState',department:'#filterDepartment',activity:'#filterActivity'};
function currentFilters(){let res={};for(const [key,sel] of Object.entries(filmap)){const value=qs(sel).value;if(value)res[key]=value}return res}
function query(includeUrl=true){const fil=currentFilters();return new URLSearchParams(fil).toString()}
function updateUrl(){const q=query();history.replaceState(null,'',location.pathname+(q?'?'+q:''))}
function setOptions(select,items,baseText,value){
 const current=value||select.value;select.replaceChildren();
 if(baseText){let opt=new Option(baseText,'');select.append(opt)}
 for(const it of items){const val=typeof it==='string'?it:it.value;const label=typeof it==='string'?it:it.label;select.append(new Option(label,val))}
 if([...select.options].some(x=>x.value===current))select.value=current;
}
function syncFilters(meta){
 const urlQ=new URLSearchParams(location.search);
 for(const key of ['start','end']){
   const el=qs(filmap[key]);let choice=el.value||urlQ.get(key)||(key==='start'?meta.start:meta.end);
   setOptions(el,meta.month_options,'',choice);
 }
 setOptions(qs('#filterDepartment'),meta.departments,'جميع الجهات',qs('#filterDepartment').value||urlQ.get('department'));
 setOptions(qs('#filterActivity'),meta.activities,'جميع الأنشطة',qs('#filterActivity').value||urlQ.get('activity'));
 for(const key of ['service','entity','state'])if(urlQ.has(key)&&!qs(filmap[key]).value)qs(filmap[key]).value=urlQ.get(key);
}
function resizeAllCharts(){
 if(!window.Plotly)return;
 document.querySelectorAll('.chart.js-plotly-plot').forEach(el=>{
   if(el.offsetWidth>0)Promise.resolve(Plotly.Plots.resize(el)).catch(()=>{});
 });
}
function chart(target,traces,extra={}){
 const obj=document.getElementById(target); if(!obj||!window.Plotly)return;
 const layout={autosize:true,paper_bgcolor:'rgba(0,0,0,0)',plot_bgcolor:'rgba(0,0,0,0)',font:{family:'Arial, Tahoma',size:11,color:'#6b8490'},
 margin:{l:44,r:16,t:13,b:45},showlegend:true,legend:{orientation:'h',x:.5,xanchor:'center',y:-.23,font:{size:11}},
 xaxis:{gridcolor:'#eff2f3',zeroline:false,automargin:true},yaxis:{gridcolor:'#edf1f2',zeroline:false,automargin:true},
 hovermode:'x unified',barmode:'stack',...extra};
 Plotly.react(obj,traces,layout,{displayModeBar:false,responsive:true,locale:'en'});
}
function bars(months){return months.map(x=>x.label)}
function render(data){
 if(data.empty){error('لا توجد بيانات صالحة في الملف الحالي.');return}
 const k=data.kpis,m=data.monthly;
 qs('#periodText').textContent='الفترة: '+data.monthly[0].label+' إلى '+data.monthly[m.length-1].label;
 const info=data.last_import;
 qs('#updatedText').textContent=info?'آخر رفع: '+info.filename+' • إجمالي السجلات: '+formatN(info.total):'البيانات المحفوظة';
 const cards=[
 ['إجمالي الطلبات',formatN(k.total),'خلال الفترة المختارة',colors.plum],
 ['طلبات التجديد',formatN(k.renew),'من جميع حالات الطلب',colors.teal],
 ['طلبات الإصدار',formatN(k.issue),'من جميع حالات الطلب',colors.orange],
 ['متوسط الطلبات الشهري',formatN(k.avg_monthly),'طلب لكل شهر',colors.plum],
 ['الالتزام بـ SLA',formatP(k.sla),`من ${formatN(k.sla_base)} طلب منته ومصنف`,colors.teal],
 ['المواءمة العامة',formatP(k.alignment),'من جميع الطلبات المسجلة',colors.plum],
 ['الطلبات المفتوحة',formatN(k.open),`${formatN(k.overdue_open)} منها متأخرة`,colors.coral],
 ['متوسط أيام المعالجة',k.avg_days==null?'—':formatN(k.avg_days),'للمعاملات المنتهية',colors.orange]];
 const grid=qs('#kpiGrid');grid.replaceChildren();
 for(const [label,value,hint,color] of cards){
 const el=document.createElement('article');el.className='kpi';el.style.setProperty('--accent',color);
 for(const [cls,txt] of [['label',label],['value',value],['hint',hint]]){
 const sub=document.createElement('div');sub.className=cls;sub.textContent=txt;el.append(sub)}grid.append(el)}
 const x=bars(m);const traces=[
 {x,y:m.map(v=>v.renew),type:'bar',name:'تجديد',marker:{color:colors.plum},hovertemplate:'%{y} طلب تجديد<extra></extra>'},
 {x,y:m.map(v=>v.issue),type:'bar',name:'إصدار',marker:{color:colors.teal},hovertemplate:'%{y} طلب إصدار<extra></extra>'}];
 chart('volumeChart',traces,{bargap:.36});
 chart('growthChart',[{x,y:m.map(v=>v.growth),type:'bar',marker:{color:m.map(v=>v.growth!=null&&v.growth<0?colors.coral:colors.teal)},hovertemplate:'%{y:.1f}%<extra></extra>'}],{showlegend:false,barmode:'relative',yaxis:{ticksuffix:'%',gridcolor:'#edf1f2',zeroline:true,zerolinecolor:'#b8c8cc'},margin:{l:43,r:5,t:18,b:45}});
 chart('statesChart',[
 {x,y:m.map(v=>v.approved),name:'موافقة',type:'bar',marker:{color:colors.teal}},
 {x,y:m.map(v=>v.rejected),name:'رفض',type:'bar',marker:{color:colors.plum}},
 {x,y:m.map(v=>v.open),name:'قيد العمل',type:'bar',marker:{color:colors.coral}}],{bargap:.36});
 chart('daysChart',[{x,y:m.map(v=>v.avg_days),name:'متوسط الأيام',mode:'lines+markers+text',type:'scatter',marker:{color:colors.orange,size:9},line:{color:colors.orange,width:3},text:m.map(v=>v.avg_days==null?'':formatN(v.avg_days)),textposition:'top center',textfont:{size:11,color:colors.orange}}],{showlegend:false,yaxis:{title:'يوم عمل',gridcolor:'#edf1f2'},margin:{l:48,r:8,t:31,b:41}});
 chart('slaChart',[{x,y:m.map(v=>v.sla),name:'التزام SLA للمنتهية',mode:'lines+markers',type:'scatter',line:{color:colors.teal,width:3},marker:{size:10}},
 {x,y:m.map(v=>v.alignment),name:'المواءمة لجميع الطلبات',type:'scatter',mode:'lines+markers',line:{color:colors.plum,width:3},marker:{size:10}}],
 {yaxis:{range:[0,102],ticksuffix:'%',gridcolor:'#edf1f2',zeroline:false},margin:{l:53,r:14,t:35,b:45}});
 const entities=[['تجاري',k.commercial,colors.plum],['صحي',k.health,colors.teal],['غير محدد',k.unspecified,colors.orange]];
 chart('entityChart',[{labels:entities.map(x=>x[0]),values:entities.map(x=>x[1]),type:'pie',hole:.65,sort:false,marker:{colors:entities.map(x=>x[2])},textinfo:'label+percent',textposition:'outside',hovertemplate:'%{label}: %{value} طلب (%{percent})<extra></extra>'}],
 {margin:{l:38,r:38,t:30,b:25},showlegend:false,annotations:[{text:formatN(k.total)+'<br><span style="font-size:12px">طلب</span>',x:.5,y:.5,showarrow:false,font:{size:29,color:colors.navy}}]});
 const top=data.top_activities.slice().reverse();
 const labels=top.map(x=>x.name.length>34?x.name.slice(0,32)+'…':x.name);
 chart('topChart',[{y:labels,x:top.map(x=>x.count),type:'bar',orientation:'h',marker:{color:colors.plum},text:top.map(x=>x.count),textposition:'outside',customdata:top.map(x=>x.name),hovertemplate:'%{customdata}<br>%{x} طلب<extra></extra>'}],
 {showlegend:false,margin:{l:36,r:26,t:8,b:30},xaxis:{gridcolor:'#edf1f2'},yaxis:{automargin:true,tickfont:{size:10}}});
 const depts=data.departments.slice().reverse();
 chart('deptChart',[{y:depts.map(v=>v.name),x:depts.map(v=>v.count-v.open),name:'منتهية',orientation:'h',type:'bar',marker:{color:colors.teal}},
 {y:depts.map(v=>v.name),x:depts.map(v=>v.open),name:'مفتوحة',orientation:'h',type:'bar',marker:{color:colors.coral}}],
 {margin:{l:30,r:20,t:15,b:32},yaxis:{automargin:true}});
 const table=qs('#monthlyTable tbody');table.replaceChildren();
 for(const row of m){const tr=document.createElement('tr');for(const t of [row.label,row.total,row.renew,row.issue,pctV(row.sla),pctV(row.alignment),pctV(row.growth),row.avg_days==null?'—':formatN(row.avg_days),row.open]){
 const td=document.createElement('td');td.textContent=String(t);tr.append(td)}table.append(tr)}
 const ul=qs('#insightList');ul.replaceChildren();data.insights.forEach(str=>{let li=document.createElement('li');li.textContent=str;ul.append(li)});
 if(k.open>0){let a=qs('#alert');a.classList.remove('hidden');a.textContent=`تنبيه متابعة: ${formatN(k.open)} طلب مفتوح ضمن الفترة المختارة، منها ${formatN(k.overdue_open)} مصنفة متأخرة. لا تدخل الطلبات المفتوحة في مقام SLA للطلبات المنتهية.`}else{qs('#alert').classList.add('hidden')}
}
async function loadDashboard(){
 try{
  const args=query();const data=await api('/api/dashboard'+(args?'?'+args:''));
  if(!appState.boot){syncFilters(data.meta);appState.boot=true;return loadDashboard()}
  appState.data=data;visible('app');render(data);
  requestAnimationFrame(()=>requestAnimationFrame(resizeAllCharts));
 }catch(e){error(e.message);visible('app');}
}
async function bootstrap(){
 try{const r=await api('/api/bootstrap');appState.role=r.role;appState.csrf=r.csrf;
    qs('#sessionRole').textContent=r.role==='admin'?'صلاحية المسؤول':'صلاحية مشاهدة';qs('#uploadBtn').classList.toggle('hidden',r.role!=='admin');
    appState.boot=false;await loadDashboard();
 }catch(e){visible('login')}
}
async function doLogin(e){
 e.preventDefault();const pass=qs('#loginPassword').value;
 try{await api('/api/login',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({password:pass})});qs('#loginError').classList.add('hidden');await bootstrap()}
 catch(e){qs('#loginError').textContent=e.message;qs('#loginError').classList.remove('hidden')}
}
async function doUpload(e){e.preventDefault();const file=qs('#uploadFile').files[0];if(!file)return;
 const mode=qs('input[name="mode"]:checked').value;const form=new FormData();form.append('file',file);form.append('mode',mode);
 const b=qs('#doUpload');b.disabled=true;b.textContent='جاري التحقق من البيانات ورفعها...';
 try{const r=await api('/api/upload',{method:'POST',headers:{'X-CSRF-Token':appState.csrf},body:form});
 qs('#uploadError').classList.add('hidden');qs('#uploadModal').classList.add('hidden');qs('#uploadFile').value='';
 toast(`تم تحديث الداشبورد: ${formatN(r.total)} سجل متاح`);appState.boot=false;await loadDashboard();}
 catch(e){qs('#uploadError').textContent=e.message;qs('#uploadError').classList.remove('hidden')}
 finally{b.disabled=false;b.textContent='رفع وتحديث لوحة المعلومات'}
}
function fileDownload(url){const link=document.createElement('a');link.href=url+(query()?'?'+query():'');document.body.append(link);link.click();link.remove()}
async function shareLink(){
 updateUrl();try{await navigator.clipboard.writeText(location.href);toast('تم نسخ رابط العرض الحالي')}
 catch(e){window.prompt('انسخ رابط العرض',location.href)}
}
function bindUI(){
 qs('#loginForm').addEventListener('submit',doLogin);
 qs('#uploadForm').addEventListener('submit',doUpload);
 qs('#uploadBtn').addEventListener('click',()=>qs('#uploadModal').classList.remove('hidden'));
 qs('#closeModal').addEventListener('click',()=>qs('#uploadModal').classList.add('hidden'));
 qs('#uploadModal').addEventListener('click',e=>{if(e.target===qs('#uploadModal'))qs('#uploadModal').classList.add('hidden')});
 qs('#uploadFile').addEventListener('change',e=>qs('#selectedFile').textContent=e.target.files[0]?.name||'اختر ملف');
 for(const sel of Object.values(filmap))qs(sel).addEventListener('change',()=>{updateUrl();loadDashboard()});
 qs('#clearFilters').addEventListener('click',()=>{for(const key of ['service','entity','state','department','activity'])qs(filmap[key]).value='';qs('#filterStart').replaceChildren();qs('#filterEnd').replaceChildren();appState.boot=false;history.replaceState(null,'',location.pathname);loadDashboard()});
 qs('#reloadBtn').addEventListener('click',loadDashboard);
 qs('#shareBtn').addEventListener('click',shareLink);
 qs('#exportBtn').addEventListener('click',()=>fileDownload('/api/export.pptx'));
 qs('#csvBtn').addEventListener('click',()=>fileDownload('/api/export.csv'));
 qs('#logoutBtn').addEventListener('click',async()=>{await api('/api/logout',{method:'POST',body:'{}'});appState.boot=false;qs('#loginPassword').value='';visible('login')});
 document.querySelectorAll('.nav-link').forEach(a=>a.addEventListener('click',()=>{document.querySelectorAll('.nav-link').forEach(x=>x.classList.remove('active'));a.classList.add('active')}));
 setInterval(()=>{if(appState.data&&document.visibilityState==='visible')loadDashboard()},60000);
}
window.addEventListener('resize',()=>requestAnimationFrame(resizeAllCharts));
window.addEventListener('DOMContentLoaded',()=>{bindUI();bootstrap()});
