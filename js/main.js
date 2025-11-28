const excelInput = document.getElementById('excelFile');
const listEl = document.getElementById('list');
const detailHeader = document.getElementById('detailHeader');
const productSummary = document.getElementById('productSummary');
const chartEl = document.getElementById('chart');
const tableContainer = document.getElementById('tableContainer');
const sortMode = document.getElementById('sortMode');
const txtFilter = document.getElementById('txtFilter');
const btnReset = document.getElementById('btnReset');

let rawRows = [], monthKeys = [], normalizedMonths = [], processed = [], selectedIndex = -1;
const chart = echarts.init(chartEl, null, {renderer: 'canvas'});

function parseExcelFile(file) {
  const reader = new FileReader();
  reader.onload = (e) => {
    const data = e.target.result;
    let workbook;
    try { workbook = XLSX.read(data, {type: 'binary'}); }
    catch { try { workbook = XLSX.read(data, {type:'array'}); } 
            catch { alert('读取文件失败，请确认是有效的 Excel 文件。'); return; } }
    const firstSheetName = workbook.SheetNames[0];
    const ws = workbook.Sheets[firstSheetName];
    const json = XLSX.utils.sheet_to_json(ws, {defval:null});
    if(!json || json.length===0){ listEl.innerHTML='<div class="no-data">Excel 表为空或未识别到数据。</div>'; return; }
    rawRows=json; detectMonthColumns(Object.keys(json[0]||{})); buildProcessedData(); renderList(); clearDetail();
  };
  reader.readAsBinaryString(file);
}

function detectMonthColumns(headers){
  monthKeys.length=0; normalizedMonths.length=0;
  headers.forEach(h=>{
    if(!h) return;
    let dateObj=null;
    if(typeof h==="number") dateObj=new Date((h-25569)*86400*1000);
    else if(typeof h==="string"){ const parsed=new Date(h); if(!isNaN(parsed)) dateObj=parsed; }
    if(dateObj){ const y=dateObj.getFullYear(), m=String(dateObj.getMonth()+1).padStart(2,"0"); monthKeys.push(h); normalizedMonths.push(`${y}-${m}`); return; }
    const s=String(h).trim(), m=s.match(/(\d{4})[^\d]{0,3}(\d{1,2})/);
    if(m){ monthKeys.push(s); normalizedMonths.push(`${m[1]}-${m[2].padStart(2,'0')}`); }
  });
}

function parseNumber(v){ 
  if(v===null||v===undefined) return 0;
  if(typeof v==='number') return v;
  const s=String(v).replace(/,/g,'').trim();
  if(s===''||['#N/A','NA','n/a'].includes(s.toUpperCase())) return 0;
  const m=s.match(/-?[\d,.]+(\.\d+)?/); if(!m) return 0;
  const num=Number(m[0]); return isNaN(num)?0:num;
}

function buildProcessedData(){
  processed = rawRows.map((row, idx) => {
    let total = 0, months = {};
    monthKeys.forEach((key, i) => {
      const val = parseNumber(row[key]);
      months[normalizedMonths[i]] = val;
      total += val;
    });

    return {
      __rowIndex: idx,
      raw: row,
      customer_name: row['客户名称'] || row['客户包名称'] || row['客户包编码'] || (`行${idx+1}`),
      customer_code: row['客户编码'] || row['客户包编码'] || '', // 新增字段
      customer_id: row['客户包编码'] || row['逻辑网格id'] || '',
      district: row['区县'] || '',
      team: row['班组'] || '',
      grid: row['逻辑网格'] || row['逻辑网格id'] || '',
      manager: row['客户经理'] || '',
      products: {
        tianyi: parseNumber(row['天翼']),
        broadband: parseNumber(row['宽带']),
        phone: parseNumber(row['固话']),
        itv: parseNumber(row['ITV'])
      },
      months,
      total_income: total
    };
  });
}


function formatMoney(n){ return (typeof n==='number'?n:Number(n)||0).toLocaleString(undefined,{minimumFractionDigits:2,maximumFractionDigits:2}); }

function renderList(){
  if(!processed||processed.length===0){ listEl.innerHTML='<div class="no-data">未读取到数据。</div>'; return; }
  const filter=txtFilter.value.trim().toLowerCase(), mode=sortMode.value;
  let arr=processed.slice();
  if(filter) arr=arr.filter(r=> (r.customer_name&&r.customer_name.toLowerCase().includes(filter))||(r.manager&&r.manager.toLowerCase().includes(filter))||(r.customer_code&&r.customer_code.toLowerCase().includes(filter)));
  arr.sort((a,b)=>mode==='desc'?b.total_income-a.total_income:a.total_income-b.total_income);

  listEl.innerHTML='';
  arr.forEach((r,idx)=>{
    const item=document.createElement('div'); item.className='item';
    if(processed.indexOf(r)===selectedIndex) item.classList.add('selected');
    const left=document.createElement('div'); left.className='meta';
    const name=document.createElement('div'); name.className='name'; name.textContent=`${idx+1}. ${r.customer_name}`;
    const sub=document.createElement('div'); sub.className='sub'; sub.innerHTML = `${r.manager ? '客户经理：' + r.manager + '&nbsp;&nbsp;' : ''}${r.district ? r.district + '&nbsp;&nbsp;' : ''}${r.team ? r.team : ''}`;

    left.appendChild(name); left.appendChild(sub);
    const right=document.createElement('div'); right.style.textAlign='right';
    const amt=document.createElement('div'); amt.className='amount'; amt.textContent=formatMoney(r.total_income);
    const idsmall=document.createElement('div'); idsmall.className='small'; idsmall.textContent=r.customer_code||'';
    right.appendChild(amt); right.appendChild(idsmall);
    item.appendChild(left); item.appendChild(right);
    item.addEventListener('click',()=>{ selectedIndex=processed.indexOf(r); renderList(); renderDetail(r); });
    listEl.appendChild(item);
  });
}

function clearDetail(){
  detailHeader.innerHTML='<div class="no-data">未选择客户</div>';
  productSummary.innerHTML='<div class="no-data">选择客户后显示产品（天翼/宽带/固话/ITV）与月份明细。</div>';
  chart.clear(); tableContainer.innerHTML='';
}

function renderDetail(r){
  // header
  detailHeader.innerHTML = '';
  const h = document.createElement('div');
  
  // 客户名称
  const title = document.createElement('div');
  title.style.fontSize = '16px';
  title.style.fontWeight = '700';
  title.textContent = r.customer_name;
  
  // 客户编码 + 客户包编码 + 区域 + 网格
  const info = document.createElement('div');
  info.style.marginTop = '6px';

  info.innerHTML = `
    <div class="small">客户编码：${r.customer_code || ''}</div>
    <div class="small">客户包编码：${r.customer_id || ''}</div>
    <div class="small">区域：${r.district || ''}</div>
    <div class="small">网格：${r.team || ''}</div>
`  ;
  
  h.appendChild(title);
  h.appendChild(info);
  detailHeader.appendChild(h);

  // product summary 不变
  productSummary.innerHTML = '';
  const ps = document.createElement('div');
  ps.innerHTML = `
    <div style="display:flex;gap:8px;align-items:center">
      <div class="badge">总收入 ${formatMoney(r.total_income)}</div>
    </div>
    <div style="margin-top:8px">
      <div class="small">产品情况：</div>
      <div style="display:flex;gap:10px;margin-top:6px">
        <div class="card" style="padding:8px">
          <div style="font-weight:700">${r.products.tianyi}</div>
          <div class="small">天翼</div>
        </div>
        <div class="card" style="padding:8px">
          <div style="font-weight:700">${r.products.broadband}</div>
          <div class="small">宽带</div>
        </div>
        <div class="card" style="padding:8px">
          <div style="font-weight:700">${r.products.phone}</div>
          <div class="small">固话</div>
        </div>
        <div class="card" style="padding:8px">
          <div style="font-weight:700">${r.products.itv}</div>
          <div class="small">ITV</div>
        </div>
      </div>
    </div>
  `;
  productSummary.appendChild(ps);

  // 保持 info 和 product-summary 高度一致
  const infoHeight = detailHeader.offsetHeight;
  const productHeight = productSummary.offsetHeight;
  const maxHeight = Math.max(infoHeight, productHeight);
  detailHeader.style.height = maxHeight + 'px';
  productSummary.style.height = maxHeight + 'px';

  // 月收入图表 & 表格渲染
  const monthsArr = Object.keys(r.months || {});
  monthsArr.sort((a,b)=> a.localeCompare(b)); 
  const values = monthsArr.map(m => r.months[m] || 0);

  const option = {
    tooltip: { trigger: 'axis', formatter: function(params){
      const p = params[0];
      return `${p.axisValue}<br/>收入：${formatMoney(p.data)}`;
    }},
    xAxis: { type: 'category', data: monthsArr, boundaryGap: false },
    yAxis: { type: 'value', name: '收入（元）' },
    grid: { left: '6%', right: '6%', bottom: '12%' },
    series: [{ name: '月收入', type: 'line', data: values, smooth: true, areaStyle: {} }]
  };
  chart.setOption(option, true);

  let html = '';
  if(monthsArr.length === 0){
    html = '<div class="no-data">未识别到月份列，无法显示月度明细。</div>';
  } else {
    html = '<table><thead><tr><th>月份</th><th>收入</th></tr></thead><tbody>';
    monthsArr.forEach(m=>{
      html += `<tr><td>${m}</td><td>${formatMoney(r.months[m] || 0)}</td></tr>`;
    });
    html += '</tbody></table>';
  }
  //tableContainer.innerHTML = html;
  tableContainer.innerHTML = '';
}

// events
excelInput.addEventListener('change', (ev)=>{ const f=ev.target.files && ev.target.files[0]; if(f) parseExcelFile(f); });
sortMode.addEventListener('change', renderList);
txtFilter.addEventListener('input', renderList);
btnReset.addEventListener('click', ()=>{ txtFilter.value=''; sortMode.value='desc'; renderList(); });
window.addEventListener('resize', ()=> chart.resize());
