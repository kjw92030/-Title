const $=id=>document.getElementById(id);
const fmt=x=>Number(x).toLocaleString('ko-KR',{maximumFractionDigits:3});
let worker, nextId=0, requests=new Map(), sectors=[], rowChoices=[], cells=[], basket=[];
let workerReady=false, activeDataset='basic', datasetInfo={};
const datasets={
 basic:{label:'기본부문',file:'io-2023.xlsx'},
 small:{label:'통합소분류',file:'io-small.xlsx'},
 medium:{label:'통합중분류',file:'io-medium.xlsx'},
 large:{label:'통합대분류',file:'io-large.xlsx'}
};
function text(tag,value){const e=document.createElement(tag);e.textContent=value;return e;}
function call(type,payload){return new Promise((resolve,reject)=>{const id=++nextId;requests.set(id,{resolve,reject});worker.postMessage({id,type,payload});});}
function startWorker(){
 worker=new Worker('./worker.js');
 worker.onmessage=({data:d})=>{
  if(d.type==='status'){$('status').textContent=d.text;return;}
  const p=requests.get(d.id);if(!p)return;requests.delete(d.id);
  d.ok?p.resolve(d.result):p.reject(Error(d.error));
 };
 worker.onerror=e=>{for(const p of requests.values())p.reject(Error(e.message||'Worker 실행 오류'));requests.clear();};
}
function options(el,list){el.replaceChildren(...list.map(x=>{const o=document.createElement('option');o.value=x.code;o.textContent=x.code+' '+x.name;return o;}));}

function setExplorerDisabled(disabled){
 for(const id of ['sheet','row','row-search','col-search'])$(id).disabled=disabled;
}
const pending=new Map();
async function getDataset(id){
 if(datasetInfo[id])return datasetInfo[id];
 if(!pending.has(id)){
  const job=(async()=>{
   const config=datasets[id],response=await fetch('./data/'+config.file);
   if(!response.ok)throw Error(config.file+' 다운로드 실패 (HTTP '+response.status+')');
   if(!worker)startWorker();
   if(!workerReady){
    workerReady=call('init',{}).catch(e=>{workerReady=false;throw e;});
   }
   await workerReady;
   const info=await call('load',{dataset:id,label:config.label,buffer:await response.arrayBuffer()});
   datasetInfo[id]=info;refreshSectors();return info;
  })().finally(()=>pending.delete(id));
  pending.set(id,job);
 }
 return pending.get(id);
}
async function loadDataset(id){
 const buttons=[...document.querySelectorAll('.dataset-button')];
 buttons.forEach(b=>b.disabled=true);
 $('status').textContent=datasets[id].label+' 자료 다운로드 중…';
 try{
  const info=await getDataset(id);activeDataset=id;
  $('row-search').value='';$('col-search').value='';$('row').replaceChildren();
  $('sheet').replaceChildren(...info.sheets.map(s=>text('option',s)));
  $('validation').textContent=datasets[id].label+' · '+info.sectors.length+'개 부문\n'+JSON.stringify(info.residuals,null,2)+'\n허용값: 1e-8 미만';
  setExplorerDisabled(false);await explore();
  $('status').textContent=datasets[id].label+' 준비 완료 · 계수 검증 통과';
 }catch(e){$('status').textContent='불러오기 실패: '+e.message;}
 finally{buttons.forEach(b=>b.disabled=false);}
}
let exploreSeq=0;
async function explore(){
 const seq=++exploreSeq;
 try{
  const d=await call('explore',{dataset:activeDataset,sheet:$('sheet').value,row:$('row').value});
  if(seq!==exploreSeq)return;
  rowChoices=d.rows;filterRows(d.selected);cells=d.cells;
  $('unit').textContent=d.unit;
  $('explore-caption').textContent=$('sheet').value+' / 행: '+d.selected+' '+d.name;
  renderCells();
 }catch(e){$('status').textContent='자료 탐색 실패: '+e.message;}
}
function filterRows(selected){
 const q=$('row-search').value.trim();
 options($('row'),rowChoices.filter(r=>(r.code+' '+r.name).includes(q)));
 if([...$('row').options].some(o=>o.value===selected))$('row').value=selected;
}
function renderCells(){
 const q=$('col-search').value.trim();
 $('cells').replaceChildren(...cells.filter(c=>(c.code+' '+c.name).includes(q)).map(c=>{
  const tr=document.createElement('tr');
  for(const v of [c.code,c.name,c.value===null?'빈 셀':typeof c.value==='number'?fmt(c.value):c.value])tr.append(text('td',v));
  return tr;
 }));
}
document.querySelectorAll('.dataset-button').forEach(b=>b.addEventListener('click',()=>loadDataset(b.dataset.dataset)));
$('sheet').onchange=()=>{$('row-search').value='';$('row').replaceChildren();explore();};
$('row').onchange=explore;
$('row-search').oninput=()=>{filterRows();if($('row').value)explore();else{++exploreSeq;$('cells').replaceChildren();$('explore-caption').textContent='일치하는 행이 없습니다.';}};
$('col-search').oninput=renderCells;
let saved=[], compared=[], editing=null, serial=0, busy=false;
const cache=new Map();
const status=s=>{$('sim-status').textContent=s;};
const selectedDataset=()=>$('sim-dataset').value;
const metrics=[['demand','최종수요 증가액'],['production','생산유발액'],['imports','수입유발액'],['valueAdded','부가가치유발액'],['multiplier','투입 1원당 생산'],['importRatio','투입 1원당 수입'],['valueRatio','투입 1원당 부가가치']];
function refreshSectors(){
 sectors=datasetInfo[selectedDataset()]?.sectors||[];
 const q=$('sector-search').value.trim().toLowerCase(), prev=$('sector').value;
 options($('sector'),sectors.filter(s=>(s.code+' '+s.name).toLowerCase().includes(q)));
 if([...$('sector').options].some(o=>o.value===prev))$('sector').value=prev;
 for(const id of ['sector-search','amount','sector'])$(id).disabled=!sectors.length;
 $('add').disabled=!$('sector').value;
}
function button(label,fn){const b=text('button',label);b.type='button';b.onclick=fn;b.disabled=busy;return b;}
function renderBasket(){
 $('basket').replaceChildren(...basket.map((s,i)=>{
  const li=text('li',`${s.code} ${s.name} · ${fmt(s.amount)} 백만원 `);
  li.append(button('삭제',()=>{basket.splice(i,1);renderBasket();}));return li;
 }));
 $('save').disabled=!basket.length;
 $('save').textContent=editing===null?'시나리오 저장':'수정 저장';
}
function resetDraft(){editing=null;basket=[];$('scenario-name').value='';$('sector-search').value='';refreshSectors();renderBasket();}
function resetResults(){compared=[];$('results').hidden=true;}
function renderSaved(){
 $('scenario-list').replaceChildren(...saved.map(s=>{
  const li=text('li',`${s.name} · ${datasets[s.dataset].label} · ${s.shocks.length}개 항목 · ${fmt(s.shocks.reduce((t,x)=>t+x.amount,0))} 백만원 `);
  li.append(button('수정',()=>{
   editing=s.id;$('sim-dataset').value=s.dataset;$('scenario-name').value=s.name;
   basket=s.shocks.map(x=>({...x}));$('sector-search').value='';refreshSectors();renderBasket();$('scenario-name').focus();status('시나리오 수정 중');
  }),button('삭제',()=>{
   saved=saved.filter(x=>x.id!==s.id);cache.delete(s.id);if(editing===s.id)resetDraft();resetResults();renderSaved();status('시나리오를 삭제했습니다.');
  }));return li;
 }));
 $('run').disabled=busy||!saved.length;
 syncDetail();
}
$('sector-search').oninput=refreshSectors;
$('sim-dataset').onchange=()=>{resetDraft();status(sectors.length?'산업과 증가액을 입력하세요.':'선택 자료를 먼저 불러오세요.');};
$('sim-load').onclick=async()=>{
 const id=selectedDataset();$('sim-load').disabled=true;status(datasets[id].label+' 불러오는 중…');
 try{const d=await getDataset(id);if(id===selectedDataset())status(`${datasets[id].label} · ${d.sectors.length}개 부문 · 계수 검증 통과`);}
 catch(e){status('불러오기 실패: '+e.message);}
 finally{$('sim-load').disabled=false;}
};
$('scenario').onsubmit=e=>{
 e.preventDefault();const amount=Number($('amount').value),s=sectors.find(x=>x.code===$('sector').value);
 if(!s||!Number.isFinite(amount)||amount<=0){status('산업과 양의 증가액을 확인하세요.');return;}
 basket.push({...s,amount});renderBasket();status('산업을 추가했습니다.');
};
$('clear').onclick=()=>{resetDraft();status('새 시나리오를 작성하세요.');};
$('save').onclick=()=>{
 if(!basket.length)return;
 const id=editing??++serial,name=$('scenario-name').value.trim()||`시나리오 ${id}`;
 const s={id,name,dataset:selectedDataset(),shocks:basket.map(x=>({...x}))};
 if(editing===null)saved.push(s);else saved=saved.map(x=>x.id===id?s:x);
 cache.delete(id);resetDraft();resetResults();renderSaved();status('저장했습니다. 저장된 시나리오를 분석할 수 있습니다.');
};
$('run').onclick=async()=>{
 if(busy||!saved.length)return;
 busy=true;$('editor').disabled=true;renderSaved();resetResults();status('계산 중…');
 const out=[];
 try{
  for(const s of saved){
   let r;
   try{r=await scenarioResult(s);}
   catch(e){throw Error(s.name+': '+e.message);}
   out.push({...s,result:{...r,importRatio:r.imports/r.demand,valueRatio:r.valueAdded/r.demand}});
  }
  compared=out.sort((a,b)=>b.result.production-a.result.production);
  const choices=compared.map(s=>({code:s.id,name:`${s.name} · ${datasets[s.dataset].label}`}));
  options($('baseline'),choices);
  renderSummary();$('results').hidden=false;
  status(compared.length===1?'계산 완료. 시나리오를 더 저장하면 비교할 수 있습니다.':`${compared.length}개 시나리오 비교 완료`);
 }catch(e){status('계산 실패: '+e.message);}
 finally{busy=false;$('editor').disabled=false;renderSaved();}
};
function tableRow(values,tag='td'){
 const tr=document.createElement('tr');for(const v of values)tr.append(text(tag,v));return tr;
}
function baseline(){return compared.find(s=>String(s.id)===$('baseline').value);}
function rank(i){return compared.findIndex(s=>s.result.production===compared[i].result.production)+1;}
function diff(base,v){return [v-base,base===0?null:(v-base)/Math.abs(base)*100];}
function renderSummary(){
 const base=baseline();
 $('summary-head').replaceChildren(tableRow(['지표',...compared.map(s=>`${s.name} · ${datasets[s.dataset].label}`)],'th'));
 $('summary-rows').replaceChildren(tableRow(['생산유발액 순위',...compared.map((_,i)=>rank(i))]),...metrics.map(([key,label])=>tableRow([label,...compared.map(s=>{
  const v=s.result[key];if(s===base)return fmt(v)+' (기준)';
  const [d,p]=diff(base.result[key],v);return `${fmt(v)} (차이 ${fmt(d)} · ${p===null?'—':fmt(p)+'%'})`;
 })])));
}
let detailSeq=0, showing=null;
const inflight=new Map();
function scenarioResult(s){
 if(cache.has(s.id))return Promise.resolve(cache.get(s.id));
 if(!inflight.has(s))inflight.set(s,call('simulate',{dataset:s.dataset,shocks:s.shocks}).then(r=>{
  if(saved.includes(s))cache.set(s.id,r);
  return r;
 }).finally(()=>inflight.delete(s)));
 return inflight.get(s);
}
function syncDetail(){
 const el=$('detail'), previous=el.value;
 options(el,saved.map(s=>({code:s.id,name:`${s.name} · ${datasets[s.dataset].label}`})));
 if(saved.some(s=>String(s.id)===previous))el.value=previous;
 el.disabled=$('metric').disabled=!saved.length;
 const current=saved.find(s=>String(s.id)===el.value);
 if(!showing||showing.scenario!==current)loadDetail();
}
async function loadDetail(){
 const seq=++detailSeq,s=saved.find(x=>String(x.id)===$('detail').value);
 showing=null;$('detail-content').hidden=true;
 $('detail-status').textContent=s?'산업별 효과 계산 중…':'1번에서 시나리오를 저장하면 여기에 표시됩니다.';
 if(!s)return;
 try{
  const result=await scenarioResult(s);
  if(seq!==detailSeq)return;
  showing={scenario:s,result};renderDetail();$('detail-content').hidden=false;$('detail-status').textContent='';
 }catch(e){if(seq===detailSeq)$('detail-status').textContent='계산 실패: '+e.message;}
}
function drawBars(rows,key){
 const canvas=$('impact-chart'), w=760,h=rows.length*28+36,dpr=Math.min(window.devicePixelRatio||1,3);
 canvas.width=w*dpr;canvas.height=h*dpr;
 const c=canvas.getContext('2d');if(!c)return;
 c.scale(dpr,dpr);c.fillStyle='#FDFDFF';c.fillRect(0,0,w,h);
 c.font='14px system-ui, sans-serif';c.textBaseline='middle';
 const left=248,right=626,min=Math.min(0,...rows.map(r=>r[key])),max=Math.max(0,...rows.map(r=>r[key]));
 const scale=v=>left+(v-min)/(max-min||1)*(right-left),zero=scale(0);
 const ticks=[...new Set([min,0,max])];
 for(const tick of ticks){const x=scale(tick);c.fillStyle=tick===0?'#CDD3E0':'#E2E6F0';c.fillRect(x,4,1,h-29);}
 rows.forEach((r,i)=>{
  const y=17+i*28,v=r[key],x=scale(v);let label=`${r.code} ${r.name}`;
  if(c.measureText(label).width>left-18){while(c.measureText(label+'…').width>left-18)label=label.slice(0,-1);label+='…';}
  c.fillStyle='#1E2433';c.textAlign='left';c.fillText(label,8,y);
  c.fillStyle=v<0?'#C1232C':'#1751D0';c.fillRect(Math.min(x,zero),y-7,Math.abs(x-zero),14);
  c.fillStyle='#1E2433';c.textAlign='right';c.fillText(fmt(v),w-10,y,122);
 });
 c.fillStyle='#505A74';c.textAlign='left';c.fillText('단위: 백만원',left,h-10);
 c.textAlign='right';c.fillText('0 기준 막대',right,h-10);
}
function renderDetail(){
 if(!showing)return;
 const {scenario:s,result}=showing,key=$('metric').value;
 const rows=[...result.rows].sort((a,b)=>b[key]-a[key]);
 $('chart-title').textContent=`${$('metric').selectedOptions[0].textContent} 상위 ${Math.min(15,rows.length)}개 산업`;
 $('detail-caption').textContent=`${s.name} · ${datasets[s.dataset].label} · ${rows.length}개 부문 · 백만원`;
 drawBars(rows.slice(0,15),key);
 $('result-rows').replaceChildren(...rows.map(r=>tableRow([r.code,r.name,fmt(r.production),fmt(r.imports),fmt(r.valueAdded)])));
}
$('baseline').onchange=renderSummary;$('detail').onchange=loadDetail;$('metric').onchange=renderDetail;
function csvCell(v){let s=v==null?'':String(v);if(typeof v==='string'&&(/^[\s\uFEFF]*[=+@-]/.test(s)||/^[\t\r\n]/.test(s)))s="'"+s;return '"'+s.replaceAll('"','""')+'"';}
function csv(name,rows){
 const url=URL.createObjectURL(new Blob(['\ufeff'+rows.map(r=>r.map(csvCell).join(',')).join('\r\n')],{type:'text/csv;charset=utf-8'}));
 const a=document.createElement('a');a.href=url;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);
}
$('download').onclick=()=>{
 const base=baseline();
 csv('io-2023-comparison.csv',[['시나리오 ID','시나리오','분류표','생산유발액 순위','기준 시나리오 ID',...metrics.flatMap(([key,l])=>{const label=l+(key.endsWith('Ratio')||key==='multiplier'?'':' (백만원)');return [label,label+' 차이',l+' 증감률 (%)'];})],...compared.map((s,i)=>[s.id,s.name,datasets[s.dataset].label,rank(i),base.id,...metrics.flatMap(([k])=>[s.result[k],...diff(base.result[k],s.result[k])])])]);
};
$('download-detail').onclick=()=>csv('io-2023-industry-details.csv',[
 ['시나리오 ID','시나리오','분류표','산업코드','산업명','생산유발액 (백만원)','수입유발액 (백만원)','부가가치유발액 (백만원)'],
 ...compared.flatMap(s=>s.result.rows.map(r=>[s.id,s.name,datasets[s.dataset].label,r.code,r.name,r.production,r.imports,r.valueAdded]))]);
resetDraft();renderSaved();
