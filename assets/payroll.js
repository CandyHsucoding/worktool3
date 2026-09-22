/* Monthly payroll. Payroll inputs live on each staff record so existing backups include them. */
(() => {
  const h = React.createElement;
  const fields = [
    ['healthSelf', '健保自付'], ['healthPublic', '健保公付'],
    ['laborSelf', '勞保自付'], ['laborPublic', '勞保公付'], ['pensionPublic', '勞退公付']
  ];
  const money = value => value === null ? '待填時薪' : Number(value).toLocaleString('zh-TW', { maximumFractionDigits: 2 });
  const numeric = value => value === '' || value === undefined ? 0 : Number(value);
  const defaultRate = year => year === 2026 ? 196 : year === 2025 ? 190 : '';
  const groupFor = staff => staff.payrollBank || (staff.name.trim() === '官麗珠' ? '郵局' : ['溫美萍','郭恒妘'].includes(staff.name.trim()) ? '中信' : '未分組');
  const calculate = (staff, year, month, hours) => {
    const data = staff.payroll?.[year + '-' + month] || {};
    const rate = data.rate ?? defaultRate(year);
    const wage = rate === '' ? null : Math.round(hours * numeric(rate));
    const deductions = numeric(data.healthSelf) + numeric(data.laborSelf);
    return { staff, data, rate, hours, wage, claim: wage === null ? null : wage + numeric(data.healthPublic) + numeric(data.laborPublic) + numeric(data.pensionPublic), net: wage === null ? null : wage - deductions };
  };
  const sum = (rows, key) => rows.some(row => row[key] === null) ? null : rows.reduce((total, row) => total + row[key], 0);
  const saveBlob = (blob, name) => {
    const url = URL.createObjectURL(blob), link = document.createElement('a');
    link.href = url; link.download = name; document.body.appendChild(link); link.click();
    setTimeout(() => { URL.revokeObjectURL(url); link.remove(); }, 1000);
  };
  window.PayrollPanel = function PayrollPanel({ staffList, setStaffList, year, month, calculateTotal, showNames }) {
    const [busy, setBusy] = React.useState(false);
    const [showHealth, setShowHealth] = React.useState(false);
    const visibleFields = fields.filter(([key]) => showHealth || !key.startsWith('health'));
    const [error, setError] = React.useState('');
    const period = year + '-' + month;
    const rows = staffList.map(staff => calculate(staff, year, month, calculateTotal(staff)));
    const groups = ['中信','郵局','未分組'].map(bank => ({bank, rows: rows.filter(row => groupFor(row.staff) === bank)})).filter(group => group.rows.length);
    const title = bank => year + '年廚房臨時人員' + month + '月份薪資（' + bank + '）';
    const label = (row, index) => showNames ? row.staff.name || '未命名人員' : '人員' + (index + 1);
    const edit = (id, field, value) => {
      if (value !== '' && (!Number.isFinite(Number(value)) || Number(value) < 0)) return;
      setStaffList(previous => previous.map(staff => staff.id === id ? {
        ...staff, payroll: {...staff.payroll, [period]: {...staff.payroll?.[period], [field]: value === '' ? '' : Number(value)}}
      } : staff));
    };
    const input = (row, key, title, index) => h('input', {
      type:'number', min:0, step:'0.01', value:key === 'rate' ? row.rate : row.data[key] ?? '',
      'aria-label':label(row, index) + ' ' + title,
      className:key === 'laborSelf' ? 'labor-self-input' : undefined,
      onChange:event => edit(row.staff.id, key, event.target.value)
    });
    const canExport = () => {
      const invalid = rows.some(row => row.rate === '' || !Number.isFinite(Number(row.rate)) || Number(row.rate) < 0 || !Number.isFinite(row.hours) || row.hours < 0);
      if (invalid) { setError('請填入有效時薪，並確認工時為非負數後再產出報表。'); return false; }
      setError(''); return true;
    };
    const exportExcel = async () => {
      if (!canExport()) return;
      setBusy(true);
      try {
        if (!window.ExcelJS) throw new Error('Excel 匯出元件未載入，請重新整理後重試。');
        const workbook = new ExcelJS.Workbook();
        workbook.creator = '臨時人員工時管理系統';
        workbook.calcProperties.fullCalcOnLoad = true;
        const sheet = workbook.addWorksheet('薪資報表', {pageSetup: {
          paperSize:9, orientation:'landscape', fitToPage:true, fitToWidth:1, fitToHeight:0,
          margins:{left:0.3,right:0.3,top:0.35,bottom:0.35,header:0.1,footer:0.1}
        }});
        sheet.columns = [6,12,9,9,13,11,11,11,11,11,13,13].map(width => ({width}));
        sheet.getColumn(6).hidden = !showHealth;
        sheet.getColumn(7).hidden = !showHealth;
        let r = 1;
        groups.forEach(group => {
          const start = r;
          sheet.mergeCells(r,1,r,12); sheet.getCell(r,1).value = title(group.bank); sheet.getRow(r).height = 30; r++;
          const header = r;
          ['編號','姓名','工時','時薪','薪資'].forEach((text,index) => {
            sheet.mergeCells(r,index+1,r+1,index+1); sheet.getCell(r,index+1).value = text;
          });
          sheet.mergeCells(r,6,r,7); sheet.getCell(r,6).value = '健保';
          sheet.mergeCells(r,8,r,9); sheet.getCell(r,8).value = '勞保';
          sheet.getCell(r,10).value = '勞退'; sheet.getCell(r,11).value = '應請領'; sheet.getCell(r,12).value = '實領';
          ['自付','公付','自付','公付','公付','金額','金額'].forEach((text,index) => sheet.getCell(r+1,index+6).value = text);
          sheet.getRow(r).height = 25; sheet.getRow(r+1).height = 25; r+=2;
          const first = r;
          group.rows.forEach((row,index) => {
            const values = [index+1,showNames ? row.staff.name : '',row.hours,Number(row.rate),null,
              ...fields.map(([key]) => row.data[key] === '' || row.data[key] === undefined ? null : Number(row.data[key])),null,null];
            sheet.getRow(r).values = values;
            sheet.getCell(r,5).value = {formula:'ROUND(C'+r+'*D'+r+',0)',result:row.wage};
            sheet.getCell(r,11).value = {formula:'E'+r+'+N(G'+r+')+N(I'+r+')+N(J'+r+')',result:row.claim};
            sheet.getCell(r,12).value = {formula:'E'+r+'-N(F'+r+')-N(H'+r+')',result:row.net};
            sheet.getRow(r).height = 28;
            [4,6,7,8,9,10].forEach(col => {
              sheet.getCell(r,col).font = {name:'Microsoft JhengHei',size:11,color:{argb:'FF174EA6'}};
              sheet.getCell(r,col).dataValidation = {type:'decimal',operator:'greaterThanOrEqual',formulae:[0],allowBlank:col!==4,showErrorMessage:true,error:'請輸入非負數。'};
            });
            r++;
          });
          sheet.mergeCells(r,1,r,2); sheet.getCell(r,1).value = '合計';
          [3,5,6,7,8,9,10,11,12].forEach(col => {
            const letter=String.fromCharCode(64+col);
            const result = col===3 ? sum(group.rows,'hours') : col===5 ? sum(group.rows,'wage') :
              col===11 ? sum(group.rows,'claim') : col===12 ? sum(group.rows,'net') :
              group.rows.reduce((total,row)=>total+numeric(row.data[fields[col-6][0]]),0);
            sheet.getCell(r,col).value = {formula:'SUM('+letter+first+':'+letter+(r-1)+')',result};
          });
          const last=r;
          for(let rr=start;rr<=last;rr++) for(let col=1;col<=12;col++){
            const cell=sheet.getCell(rr,col);
            cell.font={name:'Microsoft JhengHei',size:11,...cell.font};
            cell.alignment={vertical:'middle',horizontal:col>=3&&rr>header+1?'right':'center',wrapText:true};
            if(rr>start) cell.border={top:{style:'thin'},left:{style:'thin'},bottom:{style:'thin'},right:{style:'thin'}};
            if(rr>=header&&rr<=header+1) cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFF1F5F9'}};
            if(col===8 && rr>header) cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFFFEDD5'}};
            if(rr===last) cell.font={...cell.font,bold:true};
            if(col>=3&&rr>header+1) cell.numFmt='General';
          }
          sheet.getCell(start,1).font={name:'Microsoft JhengHei',size:15,bold:true};
          r+=2;
          [[1,2,'製表'],[3,4,'出納'],[5,7,'學務主任'],[8,10,'會計主任'],[11,12,'校長']].forEach(([a,b,text])=>{
            sheet.mergeCells(r,a,r,b);sheet.getCell(r,a).value=text;sheet.getCell(r,a).font={name:'Microsoft JhengHei',size:11};
          });
          sheet.getRow(r).height=30; r+=3;
        });
        sheet.mergeCells(r,1,r,12);
        sheet.getCell(r,1).value='薪資＝工時×時薪（四捨五入至元）；應請領＝薪資＋健保公付＋勞保公付＋勞退公付；實領＝薪資－健保自付－勞保自付。';
        sheet.getRow(r).height=25;
        sheet.mergeCells(r+1,1,r+1,12);
        sheet.getCell(r+1,1).value='公付費用另列，不扣實領。保費未填時暫以 0 計算，空白欄位請於核定前補齊。';
        sheet.getRow(r+1).height=25;
        sheet.pageSetup.printArea='A1:L'+(r+1);
        saveBlob(new Blob([await workbook.xlsx.writeBuffer()],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),year+'年'+month+'月臨時人員薪資報表.xlsx');
      } catch (e) {setError(e.message || '匯出失敗，請重試。');}
      finally {setBusy(false);}
    };
    const print = () => {
      if (!canExport()) return;
      const clone = document.getElementById('payroll-report').cloneNode(true);
      clone.querySelectorAll('input').forEach(element => {
        const span=document.createElement('span'); span.textContent=element.value; element.replaceWith(span);
      });
      const popup = window.open('','_blank','width=1200,height=850');
      if (!popup) {setError('請允許彈出視窗後，再按列印／另存 PDF。');return;}
      popup.document.write('<!doctype html><html lang="zh-TW"><head><meta charset="utf-8"><title>'+year+'年'+month+'月臨時人員薪資報表</title><style>'+
        '@page{size:A4 landscape;margin:12mm}body{font-family:"Microsoft JhengHei",sans-serif;color:#000}h3{text-align:center;font-size:18px;margin:0 0 8px}.payroll-group{break-inside:avoid;margin-bottom:24px}table{width:100%;border-collapse:collapse;table-layout:fixed}th,td{border:1px solid #000;padding:7px 3px;text-align:center;font-size:12px;overflow-wrap:anywhere}th{background:#f1f5f9}.salary-signatures{display:flex;justify-content:space-between;padding:20px 0}.payroll-note{font-size:11px}tfoot{font-weight:bold}.labor-self{background:#ffedd5!important;-webkit-print-color-adjust:exact;print-color-adjust:exact}</style></head><body>'+clone.innerHTML+
        '<p class="payroll-note">薪資＝工時×時薪（四捨五入至元）；應請領＝薪資＋健保公付＋勞保公付＋勞退公付；實領＝薪資－健保自付－勞保自付。<br>公付費用另列，不扣實領。保費未填時暫以 0 計算，空白欄位請於核定前補齊。</p></body></html>');
      popup.document.close();popup.focus();setTimeout(()=>popup.print(),300);
    };
    return h('section',{className:'payroll-panel','aria-label':'薪資計算與報表'},
      h('div',{className:'payroll-toolbar'},
        h('div',null,h('h2',null,'薪資計算與報表'),h('p',null,'連動上方 '+year+' 年 '+month+' 月總工時，金額單位：新臺幣元')),
        h('div',{className:'payroll-actions'},
          h('button',{type:'button',onClick:exportExcel,disabled:busy},busy?'正在匯出…':'下載薪資 Excel'),
          h('button',{type:'button',onClick:print},'列印／另存薪資 PDF'))),
      h('p',{className:'payroll-explanation'},'薪資＝工時×時薪（四捨五入至元）；應請領＝薪資＋健保公付＋勞保公付＋勞退公付；實領＝薪資－健保自付－勞保自付。公付費用另列，不扣實領。'),
      h('p',{className:'payroll-explanation'},year===2026?'2026 年預設時薪 196 元，可逐人、逐月修改。保費欄留白供填寫，未填時暫以 0 計。':'請核對本年度時薪。保費欄留白供填寫，未填時暫以 0 計。'),
      h('label',{className:'payroll-health-toggle'},h('input',{type:'checkbox',checked:showHealth,onChange:event=>setShowHealth(event.target.checked)}),' 顯示健保欄位'),
      !showHealth && rows.some(row=>numeric(row.data.healthSelf)!==0 || numeric(row.data.healthPublic)!==0) && h('p',{className:'payroll-explanation'},'本月已有健保金額，隱藏欄位後仍保留並納入計算；可勾選「顯示健保欄位」查看或修改。'),
      h('div',{className:'payroll-banks'},rows.map((row,index)=>h('label',{key:row.staff.id},label(row,index)+' 薪轉分組 ',
        h('select',{value:groupFor(row.staff),'aria-label':label(row,index)+' 薪轉分組',onChange:event=>{
          const bank=event.target.value;setStaffList(previous=>previous.map(staff=>staff.id===row.staff.id?{...staff,payrollBank:bank}:staff));
        }},['中信','郵局','未分組'].map(bank=>h('option',{key:bank,value:bank},bank)))))),
      error&&h('p',{role:'alert',className:'payroll-error'},error),
      rows.some(row=>row.net!==null&&row.net<0)&&h('p',{className:'payroll-error'},'部分人員的自付保費超過薪資，請確認填入金額。'),
      h('div',{id:'payroll-report'},groups.map(group=>h('div',{key:group.bank,className:'payroll-group'},
        h('h3',null,title(group.bank)),
        h('div',{className:'payroll-scroll'},h('table',{className:'payroll-table'},
          h('thead',null,
            h('tr',null,...['編號','姓名','工時','時薪','薪資'].map(text=>h('th',{key:text,rowSpan:2},text)),
              showHealth && h('th',{colSpan:2},'健保'),h('th',{colSpan:2},'勞保'),h('th',null,'勞退'),h('th',null,'應請領'),h('th',null,'實領')),
            h('tr',null,...[...(showHealth ? ['自付','公付'] : []),'自付','公付','公付','金額','金額'].map((text,index)=>h('th',{key:index,className:index===(showHealth?2:0)?'labor-self':undefined},text)))),
          h('tbody',null,group.rows.map((row,index)=>h('tr',{key:row.staff.id},
            h('td',null,index+1),h('td',null,showNames?row.staff.name:''),
            h('td',null,money(row.hours)),h('td',null,input(row,'rate','時薪',index)),h('td',{'data-field':'wage'},money(row.wage)),
            ...visibleFields.map(([key,title])=>h('td',{key,className:key==='laborSelf'?'labor-self':undefined},input(row,key,title,index))),
            h('td',{'data-field':'claim'},money(row.claim)),h('td',{'data-field':'net'},money(row.net))))),
          h('tfoot',null,h('tr',null,h('td',{colSpan:2},'合計'),h('td',null,money(sum(group.rows,'hours'))),h('td',null,''),
            h('td',null,money(sum(group.rows,'wage'))),
            ...visibleFields.map(([key])=>h('td',{key,className:key==='laborSelf'?'labor-self':undefined},money(group.rows.reduce((total,row)=>total+numeric(row.data[key]),0)))),
            h('td',null,money(sum(group.rows,'claim'))),h('td',null,money(sum(group.rows,'net'))))))),
        h('div',{className:'salary-signatures'},['製表','出納','學務主任','會計主任','校長'].map(text=>h('span',{key:text},text)))))),
      h('p',{className:'payroll-explanation'},'時薪與保費按月份自動儲存，並隨「資料備份／遷移」一併備份。')
    );
  };
})();