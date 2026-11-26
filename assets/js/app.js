let rawRows=[], columns=[], sourceName='', charts={}, workbook=null, activeSheetName='';
const $=id=>document.getElementById(id), fmt=n=>Number(n||0).toLocaleString('es-HN',{maximumFractionDigits:2});
function toast(msg){const t=$('toast');t.textContent=msg;t.classList.add('show');clearTimeout(toast.timer);toast.timer=setTimeout(()=>t.classList.remove('show'),3200)}
function go(s){document.querySelectorAll('.page-section').forEach(x=>x.classList.toggle('visible',x.id==='section-'+s));document.querySelectorAll('.nav-item').forEach(x=>x.classList.toggle('active',x.dataset.section===s));$('crumb').textContent=s==='riego'?'Análisis de riego':'Conexión Excel';$('sidebar').classList.remove('open')}
document.querySelectorAll('[data-section]').forEach(b=>b.onclick=()=>go(b.dataset.section));document.querySelectorAll('[data-goto]').forEach(b=>b.onclick=()=>go(b.dataset.goto));$('mobileMenu').onclick=()=>$('sidebar').classList.toggle('open');
function esc(v){return String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))}
function norm(v){return String(v??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z0-9]/g,'')}
const synonyms={duration:['horastrabajadas','tiempoderiego','tiempoderiegohoras','duracion','duracionderiego','horasderiego','horariego','tiempo','horas','minutosderiego','tiempominutos','tiempohoras','tiempototal','horasefectivas','tiempodeaplicacion','tiempoprogramado','tiempodefuncionamiento','riegoenhoras','riegoenminutos','duracionhoras','duracionminutos']};
function detectValueColumn(){const keys=synonyms.duration.map(norm);return columns.find(c=>keys.includes(norm(c)))||columns.find(c=>keys.some(k=>norm(c).includes(k)))||columns.find(c=>rawRows.some(r=>numericValue(r[c])!==null))||''}
function numericValue(v){if(v===null||v===undefined||v==='')return null;if(typeof v==='number')return Number.isFinite(v)?v:null;const s=String(v).trim(),m=s.match(/^(\\d+):(\\d{1,2})(?::(\\d{1,2}))?$/);if(m)return +m[1]+(+m[2]/60)+((+m[3]||0)/3600);const n=Number(s.replace(',','.'));return Number.isFinite(n)?n:null}
function durationHours(v){const n=numericValue(v);if(n===null)return null;const unit=$('durationUnit')?.value||'auto';if(unit==='minutes')return n/60;if(unit==='hours')return n;if(unit==='auto')return n;return n}
function fillSelect(id,values,selectedValue){const el=$(id);el.innerHTML=values.map(v=>'<option value="'+esc(v)+'">'+esc(v)+'</option>').join('');if(selectedValue&&values.includes(selectedValue))el.value=selectedValue}
function configureMappings(){
 $('dashboard').hidden=false;$('emptyState').hidden=true;
 const value=detectValueColumn();fillSelect('valueColumn',columns,value);
 $('durationUnit').value='auto';
 $('filterColumnChoices').innerHTML=columns.map((c,i)=>'<label class="column-choice"><input type="checkbox" value="'+esc(c)+'" checked><span>'+esc(c)+'</span></label>').join('');
 fillSelect('chartDimension',columns.filter(c=>c!==value),columns.find(c=>c!==value)||'');
 $('valueColumn').onchange=()=>{fillSelect('chartDimension',columns.filter(c=>c!==$('valueColumn').value),columns.find(c=>c!==$('valueColumn').value)||'');renderDashboard()};
 $('durationUnit').onchange=renderDashboard;
 $('chartDimension').onchange=renderDashboard;
 $('filterColumnChoices').querySelectorAll('input').forEach(x=>x.onchange=renderDashboard);
 $('resetColumnConfig').onclick=()=>{configureMappings();toast('Configuración restablecida para esta hoja.')};
 $('resetFilters').onclick=()=>{document.querySelectorAll('#dynamicFilters select').forEach(s=>Array.from(s.options).forEach(o=>o.selected=false));renderDashboard()};
 renderDashboard();
}
function importRows(rows,name){if(!rows?.length){toast('No se encontraron registros.');return}rawRows=rows.filter(r=>r&&Object.values(r).some(v=>v!==null&&v!==''));if(!rawRows.length){toast('La hoja no contiene filas con datos.');return}columns=Object.keys(rawRows[0]||{});sourceName=name||'Archivo Excel';configureMappings();toast('Archivo cargado: '+rawRows.length+' registros y '+columns.length+' columnas.')}
function extractSheetTable(sheet){
 const matrix=XLSX.utils.sheet_to_json(sheet,{header:1,defval:null,raw:true,blankrows:false});
 if(!matrix.length)return {rows:[],headerRow:1,headers:[]};
 const headers=matrix[0]||[],width=Math.max(headers.length,...matrix.slice(1).map(r=>(r||[]).length));
 const names=[],used=new Set();
 for(let j=0;j<width;j++){let name=headers[j]===null||headers[j]===undefined?'':String(headers[j]).trim();if(!name)name='Columna '+(j+1);let uniqueName=name,n=2;while(used.has(norm(uniqueName)))uniqueName=name+' ('+(n++)+')';used.add(norm(uniqueName));names.push(uniqueName)}
 const rows=matrix.slice(1).map(line=>{const obj={};names.forEach((name,j)=>obj[name]=line?.[j]??null);return obj}).filter(r=>Object.values(r).some(v=>v!==null&&String(v).trim()!==''));
 return {rows,headerRow:1,headers:names};
}
function importWorkbook(wb,name){workbook=wb;sourceName=name||'Libro Excel';const selector=$('sheetSelector');selector.innerHTML=wb.SheetNames.map(n=>'<option value="'+esc(n)+'">'+esc(n)+'</option>').join('');selector.onchange=()=>loadSelectedSheet(selector.value);loadSelectedSheet(wb.SheetNames[0])}
function loadSelectedSheet(sheetName){if(!workbook||!workbook.Sheets[sheetName])return;activeSheetName=sheetName;const table=extractSheetTable(workbook.Sheets[sheetName]);if(!table.rows.length){toast('La hoja seleccionada no contiene registros debajo de la fila 1.');return}importRows(table.rows,sheetName);$('sheetInfo').textContent='Hoja: '+sheetName+' · Encabezados en fila 1 · '+table.headers.length+' columnas · '+table.rows.length+' registros';toast('Hoja "'+sheetName+'" cargada.')}
function readFile(file){const reader=new FileReader();reader.onload=e=>{try{const wb=file.name.toLowerCase().endsWith('.csv')?XLSX.read(e.target.result,{type:'string',cellDates:true}):XLSX.read(e.target.result,{type:'array',cellDates:true});importWorkbook(wb,file.name)}catch(err){console.error(err);toast('No se pudo leer el archivo.')}};file.name.toLowerCase().endsWith('.csv')?reader.readAsText(file):reader.readAsArrayBuffer(file)}
$('fileInput').onchange=e=>{const f=e.target.files[0];if(f)readFile(f)};
function valueFor(row,col){return row[col]===null||row[col]===undefined||String(row[col]).trim()===''?'(Vacío)':String(row[col])}
function getSelectedFilterColumns(){return Array.from($('filterColumnChoices').querySelectorAll('input:checked')).map(x=>x.value)}
function buildDynamicFilters(){
 const selectedCols=getSelectedFilterColumns(),container=$('dynamicFilters'),previous={};
 container.querySelectorAll('select').forEach(s=>previous[s.dataset.column]=Array.from(s.selectedOptions).map(o=>o.value));
 container.innerHTML='';
 selectedCols.forEach((col,i)=>{const values=[...new Set(rawRows.map(r=>valueFor(r,col)))].sort((a,b)=>a.localeCompare(b,'es',{numeric:true}));const label=document.createElement('label');label.textContent=col;const select=document.createElement('select');select.multiple=true;select.dataset.column=col;select.innerHTML=values.map(v=>'<option value="'+esc(v)+'">'+esc(v)+'</option>').join('');(previous[col]||[]).forEach(v=>{const o=Array.from(select.options).find(x=>x.value===v);if(o)o.selected=true});select.onchange=renderDashboard;label.appendChild(select);container.appendChild(label)});
}
function getFilteredRows(){const filters=Array.from($('dynamicFilters').querySelectorAll('select')).map(s=>({column:s.dataset.column,values:Array.from(s.selectedOptions).map(o=>o.value)}));return rawRows.filter(r=>filters.every(f=>!f.values.length||f.values.includes(valueFor(r,f.column))))}
const chartPalette=['#16A085','#2874A6','#E67E22','#8E44AD','#D35400','#2E86C1','#C0392B','#27AE60','#7D3C98','#B7950B','#148F77','#AF601A','#5D6D7E','#CA6F1E','#1F618D','#A93226'];
function makeDynamicChart(canvas,labels,values,title,index){
 const colors=labels.map((_,i)=>chartPalette[i%chartPalette.length]);
 const chart=new Chart(canvas,{type:'bar',data:{labels,datasets:[{label:$('valueColumn').value||'Valor',data:values,backgroundColor:colors,borderColor:colors.map(c=>c),borderWidth:1,borderRadius:8,maxBarThickness:48,hoverBorderWidth:2}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},title:{display:false},tooltip:{backgroundColor:'#173746',padding:12,displayColors:true,callbacks:{label:c=>fmt(c.raw)+' '+($('durationUnit').value==='minutes'?'minutos':'horas')}}},scales:{y:{beginAtZero:true,grid:{color:'#edf1f3'},ticks:{color:'#74858d'},title:{display:true,text:$('durationUnit').value==='minutes'?'Minutos':'Horas',color:'#536b75'}},x:{grid:{display:false},ticks:{color:'#536b75',autoSkip:true,maxRotation:35,minRotation:0}}}}});
 charts['dynamic_'+index]=chart;
}
function renderDynamicCharts(valid){
 const container=$('dynamicCharts'),selectedCols=getSelectedFilterColumns();
 Object.keys(charts).filter(k=>k.startsWith('dynamic_')).forEach(k=>{charts[k].destroy();delete charts[k]});
 container.innerHTML='';
 if(!selectedCols.length){container.innerHTML='<div class="empty-chart-state">Selecciona columnas en «Configurar columnas» para generar sus gráficas.</div>';return}
 selectedCols.forEach((col,index)=>{
  const grouped=new Map();
  valid.forEach(x=>{const label=valueFor(x.row,col);grouped.set(label,(grouped.get(label)||0)+x.value)});
  const entries=[...grouped.entries()].sort((a,b)=>b[1]-a[1]);
  const card=document.createElement('article');card.className='dynamic-chart-card';
  const heading=document.createElement('div');heading.className='dynamic-chart-heading';
  const eyebrow=document.createElement('span');eyebrow.textContent='ANÁLISIS POR COLUMNA';
  const title=document.createElement('h3');title.textContent=col;
  const meta=document.createElement('small');meta.textContent=entries.length+' categorías · '+fmt(entries.reduce((s,x)=>s+x[1],0))+' '+($('durationUnit').value==='minutes'?'min':'h');
  heading.append(eyebrow,title,meta);
  const box=document.createElement('div');box.className='dynamic-chart-box';
  const canvas=document.createElement('canvas');canvas.id='dynamicChart'+index;box.appendChild(canvas);
  card.append(heading,box);container.appendChild(card);
  if(entries.length)makeDynamicChart(canvas,entries.map(x=>x[0]),entries.map(x=>Number(x[1].toFixed(3))),col,index);
  else box.innerHTML='<div class="empty-chart-state">No hay valores para mostrar con los filtros actuales.</div>';
 });
}
function renderDashboard(){
 if(!rawRows.length)return;
 const valueCol=$('valueColumn').value,filterCols=getSelectedFilterColumns();
 buildDynamicFilters();
 const data=getFilteredRows(),valid=data.map(r=>({row:r,value:durationHours(r[valueCol])})).filter(x=>x.value!==null&&x.value>=0);
 const total=valid.reduce((s,x)=>s+x.value,0),avg=valid.length?total/valid.length:0,max=valid.length?Math.max(...valid.map(x=>x.value)):0;
 $('metricRecords').textContent=fmt(data.length);$('metricTotal').textContent=fmt(total)+($('durationUnit').value==='minutes'?' min':' h');$('metricAverage').textContent=fmt(avg)+($('durationUnit').value==='minutes'?' min':' h');$('metricMax').textContent=fmt(max)+($('durationUnit').value==='minutes'?' min':' h');
 $('mappingStatus').textContent='Columnas detectadas: '+columns.length+' · Filtros disponibles: '+filterCols.length+'. Los campos proceden de la fila 1 de la hoja.';
 $('resultCount').textContent=fmt(valid.length)+' registros con valor válido · '+fmt(data.length)+' registros filtrados';
 const displayCols=[...new Set([valueCol,...filterCols].filter(Boolean))],shown=data.slice(0,250);
 $('resultsTable').innerHTML=shown.length?'<table class="data-table"><thead><tr>'+displayCols.map(c=>'<th>'+esc(c)+'</th>').join('')+'</tr></thead><tbody>'+shown.map(r=>'<tr>'+displayCols.map(c=>'<td>'+esc(r[c]??'—')+'</td>').join('')+'</tr>').join('')+'</tbody></table>':'No hay registros que coincidan con los filtros seleccionados.';
 renderDynamicCharts(valid);
}
$('clearDataset').onclick=()=>{if(!rawRows.length)return;if(!confirm('¿Eliminar los registros cargados de DATA VISTA? El Excel original de SharePoint permanecerá intacto.'))return;Object.values(charts).forEach(c=>c.destroy());charts={};rawRows=[];columns=[];$('dashboard').hidden=true;$('emptyState').hidden=false;$('connectionStatus').textContent='Datos eliminados · puedes volver a conectar el Excel';$('syncMeta').textContent='';toast('Registros eliminados. El archivo original sigue intacto.')};
const proxyField=$('proxyUrl');proxyField.value=localStorage.getItem('dataVistaSharePointProxy')||'';proxyField.addEventListener('change',()=>localStorage.setItem('dataVistaSharePointProxy',proxyField.value.trim()));
async function fetchSheet(){const url=$('sheetUrl').value.trim(),proxy=$('proxyUrl').value.trim();if(!url)return toast('Pega el enlace del archivo Excel.');if(!proxy)return toast('Configura el servicio de conexión.');try{const endpoint=new URL(proxy);endpoint.searchParams.set('url',url);$('connectionStatus').textContent='Descargando archivo desde SharePoint…';const res=await fetch(endpoint.toString(),{cache:'no-store'});if(!res.ok){let detail='HTTP '+res.status;try{const err=await res.json();detail=err.error||detail}catch{}throw Error(detail)}const blob=await res.blob();if(!blob.size)throw Error('El archivo descargado está vacío.');const wb=XLSX.read(await blob.arrayBuffer(),{type:'array',cellDates:true});importWorkbook(wb,'Libro de SharePoint');$('connectionStatus').textContent='Conectado · libro importado correctamente';$('syncMeta').textContent=wb.SheetNames.length+' hojas disponibles · Consulta: '+new Date().toLocaleString('es-HN')+' · Puedes elegir la hoja que deseas analizar';go('riego')}catch(e){console.error(e);$('connectionStatus').textContent='No se pudo importar el enlace';$('syncMeta').textContent=e.message;toast('Error de conexión: '+e.message)}}
$('connectSheet').onclick=fetchSheet;
