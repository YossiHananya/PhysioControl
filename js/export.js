// ═══════════════════════════════════════════════════════
// export.js — PhysioControl
// Excel exports: monthly, quarterly (year+quarter picker), annual (year picker)
// ═══════════════════════════════════════════════════════
/* global currentUser, histCurrentMonth, adminCurrentMonth,
          _users, _rules, _logs, supa, SESSION, MHE,
          XLSX, supabase */
'use strict';

function logsToRows(logs){
  return logs.map(l=>{
    const r=getRuleById(l.ruleId);
    return {'תאריך':l.date,'שם פעילות':r.name||'','סוג פעילות':catHe(r.category),'נקודות':l.points,'הערות':l.notes||''};
  });
}
function styledExcelSheet(rows,wb,sheetName){
  const ws=XLSX.utils.json_to_sheet(rows);
  ws['!cols']=[{wch:12},{wch:30},{wch:20},{wch:10},{wch:24}];
  ws['!sheetView']=[{rightToLeft:true}];
  const range=XLSX.utils.decode_range(ws['!ref']||'A1');
  const headerFill={patternType:'solid',fgColor:{rgb:'14532d'}};
  const headerFont={bold:true,color:{rgb:'FFFFFF'},sz:11};
  const borderStyle={style:'thin',color:{rgb:'CCCCCC'}};
  const fullBorder={top:borderStyle,bottom:borderStyle,left:borderStyle,right:borderStyle};
  const catColors={'הערכת מטופל':'1a2f1a','טיפול שוטף':'1a2535','כללי':'E8F0FF','פעולות מנהלתיות':'1f1a2d'};
  for(let C=range.s.c;C<=range.e.c;C++){
    const addr=XLSX.utils.encode_cell({r:0,c:C});
    if(!ws[addr])continue;
    ws[addr].s={fill:headerFill,font:headerFont,border:fullBorder,alignment:{horizontal:'center',wrapText:true}};
  }
  for(let R=1;R<=range.e.r;R++){
    const catCell=ws[XLSX.utils.encode_cell({r:R,c:2})];
    const rowFill=catCell&&catColors[catCell.v]?{patternType:'solid',fgColor:{rgb:catColors[catCell.v]}}:{patternType:'solid',fgColor:{rgb:'FFFFFF'}};
    const isTot=R===range.e.r&&ws[XLSX.utils.encode_cell({r:R,c:1})]?.v==='סה"כ';
    for(let C=range.s.c;C<=range.e.c;C++){
      const addr=XLSX.utils.encode_cell({r:R,c:C});
      if(!ws[addr])ws[addr]={t:'z',v:''};
      ws[addr].s={fill:isTot?{patternType:'solid',fgColor:{rgb:'1A4F3A'}}:rowFill,font:isTot?{bold:true,color:{rgb:'FFFFFF'}}:{sz:10},border:fullBorder,alignment:{horizontal:C===0||C===4?'center':'right'}};
    }
  }
  XLSX.utils.book_append_sheet(wb,ws,sheetName.slice(0,31));
}
function downloadXlsx(wb,filename){
  try{
    const out=XLSX.write(wb,{bookType:'xlsx',type:'array',cellStyles:true});
    const blob=new Blob([out],{type:'application/octet-stream'});
    const url=URL.createObjectURL(blob);
    const a=document.createElement('a');a.href=url;a.download=filename;
    document.body.appendChild(a);a.click();
    setTimeout(()=>{URL.revokeObjectURL(url);a.remove();},1000);
    showToast('הקובץ הורד בהצלחה!');
  }catch(e){showToast('שגיאה בייצוא: '+e.message);}
}
function exportNutriExcel(){
  const monthKey=histCurrentMonth||tmKey();
  const [y,mo]=monthKey.split('-').map(Number);
  const target=calcMonthlyTarget(y,mo-1);
  const maxPay=getMaxPay(currentUser);
  const logs=_logs.filter(l=>l.date.startsWith(monthKey));
  if(!logs.length){showToast('אין פעילויות בחודש זה');return;}
  const wb=XLSX.utils.book_new();
  const cats=[{cat:'assessment',label:'הערכת מטופל'},{cat:'followup',label:'טיפול שוטף'},{cat:'general',label:'כללי'},{cat:'non-clinical',label:'פעולות מנהלתיות'}];
  cats.forEach(({cat,label})=>{
    const cl=logs.filter(l=>getRuleById(l.ruleId).category===cat);
    if(!cl.length)return;
    const rows=logsToRows(cl);
    const tot=cl.reduce((s,l)=>s+parseFloat(l.points||0),0);
    rows.push({'תאריך':'','שם פעילות':'סה"כ','סוג פעילות':label,'נקודות':Math.round(tot*100)/100,'הערות':''});
    styledExcelSheet(rows,wb,label);
  });
  const allPts=logs.reduce((s,l)=>s+parseFloat(l.points||0),0);
  const pay=calcPay(currentUser,allPts,target);
  const ptVal=target>0?(maxPay/target).toFixed(2):0;
  const summaryRows=[
    {'פרט':'שם','ערך':currentUser.name},
    {'פרט':'חודש','ערך':MHE[mo-1]+' '+y},
    {'פרט':'יעד נקודות','ערך':target},
    {'פרט':'נקודות שנצברו','ערך':Math.round(allPts*100)/100},
    {'פרט':'% השגה','ערך':Math.min(100,Math.round(allPts/target*100))+'%'},
    {'פרט':'ערך נקודה (₪)','ערך':ptVal},
    {'פרט':'שכר מחושב (₪)','ערך':pay},
    {'פרט':'שכר מקסימלי (₪)','ערך':maxPay},
  ];
  const wsSum=XLSX.utils.json_to_sheet(summaryRows);
  wsSum['!cols']=[{wch:22},{wch:22}];wsSum['!sheetView']=[{rightToLeft:true}];
  XLSX.utils.book_append_sheet(wb,wsSum,'סיכום');
  downloadXlsx(wb,'PhysioControl_'+currentUser.name+'_'+monthKey+'.xlsx');
}
function exportNutriPeriod(type){
  const now=new Date();
  const curYear=now.getFullYear();
  const yearOpts=[curYear,curYear-1,curYear-2].map(y=>`<option value="${y}">${y}</option>`).join('');
  if(type==='year'){
    openModal(`<div class="modal-title">📊 דוח שנתי</div>
      <div class="form-group"><label>בחר שנה</label>
        <select id="period-year" style="width:100%;padding:10px;border:1.5px solid var(--border);border-radius:10px;font-family:Rubik,sans-serif;font-size:15px">${yearOpts}</select>
      </div>
      <button class="btn btn-excel" onclick="doNutriExportYear()">⬇️ ייצא דוח שנתי</button>`);
    return;
  }
  if(type==='quarter'){
    openModal(`<div class="modal-title">📊 דוח רבעוני</div>
      <div class="form-group"><label>בחר שנה</label>
        <select id="period-year" style="width:100%;padding:10px;border:1.5px solid var(--border);border-radius:10px;font-family:Rubik,sans-serif;font-size:15px">${yearOpts}</select>
      </div>
      <div class="form-group"><label>בחר רבעון</label>
        <select id="period-quarter" style="width:100%;padding:10px;border:1.5px solid var(--border);border-radius:10px;font-family:Rubik,sans-serif;font-size:15px">
          <option value="1">רבעון 1 (ינואר–מרץ)</option>
          <option value="2">רבעון 2 (אפריל–יוני)</option>
          <option value="3">רבעון 3 (יולי–ספטמבר)</option>
          <option value="4">רבעון 4 (אוקטובר–דצמבר)</option>
        </select>
      </div>
      <button class="btn btn-excel" onclick="doNutriExportQuarter()">⬇️ ייצא דוח רבעוני</button>`);
    const curQ=Math.ceil((now.getMonth()+1)/3);
    setTimeout(()=>{const el=document.getElementById('period-quarter');if(el)el.value=curQ;},50);
    return;
  }
}
function doNutriExportYear(){
  const year=parseInt(document.getElementById('period-year')?.value||new Date().getFullYear());
  closeModal();
  let months=[];for(let m=0;m<12;m++)months.push(`${year}-${String(m+1).padStart(2,'0')}`);
  _runNutriPeriod(months,`שנתי_${year}`);
}
function doNutriExportQuarter(){
  const year=parseInt(document.getElementById('period-year')?.value||new Date().getFullYear());
  const quarter=parseInt(document.getElementById('period-quarter')?.value||1);
  closeModal();
  const startM=(quarter-1)*3;
  let months=[];for(let m=startM;m<startM+3;m++)months.push(`${year}-${String(m+1).padStart(2,'0')}`);
  _runNutriPeriod(months,`רבעון${quarter}_${year}`);
}
function _runNutriPeriod(months,label){
  const allLogs=_logs;
  const cats=['assessment','followup','general','non-clinical'];
  const wb=XLSX.utils.book_new();
  const rows=months.map(mk=>{
    const [y,mo]=mk.split('-').map(Number);
    const target=calcMonthlyTarget(y,mo-1);
    const ul=_logs.filter(l=>l.date.startsWith(mk));
    const row={'חודש':MHE[mo-1]+' '+y,'יעד':target};
    let total=0;
    cats.forEach(cat=>{const pts=ul.filter(l=>getRuleById(l.ruleId).category===cat).reduce((s,l)=>s+parseFloat(l.points||0),0);row[catHe(cat)]=Math.round(pts*100)/100;total+=pts;});
    row['סה"כ נקודות']=Math.round(total*100)/100;
    row['% השגה']=Math.min(100,Math.round(total/target*100))+'%';
    row['שכר מחושב (₪)']=calcPay(currentUser,total,target);
    row['שכר מקסימלי (₪)']=getMaxPay(currentUser);
    return row;
  });
  const ws=XLSX.utils.json_to_sheet(rows);
  ws['!sheetView']=[{rightToLeft:true}];
  XLSX.utils.book_append_sheet(wb,ws,currentUser.name.slice(0,31));
  downloadXlsx(wb,`PhysioControl_${currentUser.name}_${label}.xlsx`);
}
function exportAdminExcel(){
  if(!_users.length||!_rules.length){showToast('הנתונים עדיין נטענים, נסה שוב');return;}
  const monthKey=adminCurrentMonth||tmKey();
  const [y,mo]=monthKey.split('-').map(Number);
  const target=calcMonthlyTarget(y,mo-1);
  const users=_users.filter(u=>u.role==='physio');
  if(!users.length){showToast('אין פיזיותרפיסטים במערכת');return;}
  const wb=XLSX.utils.book_new();
  users.forEach(u=>{
    const ul=_logs.filter(l=>l.userId===u.id&&l.date.startsWith(monthKey));
    const rows=ul.length?logsToRows(ul):[{'תאריך':'','שם פעילות':'אין פעילויות','סוג פעילות':'','נקודות':0,'הערות':''}];
    const tot=ul.reduce((s,l)=>s+parseFloat(l.points||0),0);
    if(ul.length)rows.push({'תאריך':'','שם פעילות':'סה"כ','סוג פעילות':'','נקודות':Math.round(tot*100)/100,'הערות':''});
    styledExcelSheet(rows,wb,u.name.slice(0,31));
  });
  const summaryRows=users.map(u=>{
    const ul=_logs.filter(l=>l.userId===u.id&&l.date.startsWith(monthKey));
    const pts=ul.reduce((s,l)=>s+parseFloat(l.points||0),0);
    const pay=calcPay(u,pts,target);
    const maxPay=getMaxPay(u);
    const sv=u.seniority||0;
    return {'שם':u.name,'ותק':sv<2?'0-2':sv<4?'2-4':'4+','נקודות':Math.round(pts*100)/100,'יעד':target,'% השגה':Math.min(100,Math.round(pts/target*100))+'%','שכר מחושב (₪)':pay,'שכר מקסימלי (₪)':maxPay};
  });
  const wsAll=XLSX.utils.json_to_sheet(summaryRows);
  wsAll['!cols']=[{wch:22},{wch:8},{wch:12},{wch:10},{wch:10},{wch:16},{wch:16}];
  wsAll['!sheetView']=[{rightToLeft:true}];
  XLSX.utils.book_append_sheet(wb,wsAll,'סיכום');
  downloadXlsx(wb,'PhysioControl_'+monthKey+'.xlsx');
}
function exportPeriodExcel(type){
  if(!_users.length){showToast('הנתונים עדיין נטענים, נסה שוב');return;}
  const now=new Date();
  const curYear=now.getFullYear();
  // Build year options (last 3 years)
  const yearOpts=[curYear,curYear-1,curYear-2].map(y=>`<option value="${y}">${y}</option>`).join('');
  if(type==='year'){
    const html=`<div class="modal-title">📊 דוח שנתי</div>
      <div class="form-group"><label>בחר שנה</label>
        <select id="period-year" style="width:100%;padding:10px;border:1.5px solid var(--border);border-radius:10px;font-family:Rubik,sans-serif;font-size:15px">${yearOpts}</select>
      </div>
      <button class="btn btn-excel" onclick="doExportYear()">⬇️ ייצא דוח שנתי</button>`;
    openModal(html);
    return;
  }
  if(type==='quarter'){
    const html=`<div class="modal-title">📊 דוח רבעוני</div>
      <div class="form-group"><label>בחר שנה</label>
        <select id="period-year" style="width:100%;padding:10px;border:1.5px solid var(--border);border-radius:10px;font-family:Rubik,sans-serif;font-size:15px">${yearOpts}</select>
      </div>
      <div class="form-group"><label>בחר רבעון</label>
        <select id="period-quarter" style="width:100%;padding:10px;border:1.5px solid var(--border);border-radius:10px;font-family:Rubik,sans-serif;font-size:15px">
          <option value="1">רבעון 1 (ינואר–מרץ)</option>
          <option value="2">רבעון 2 (אפריל–יוני)</option>
          <option value="3">רבעון 3 (יולי–ספטמבר)</option>
          <option value="4">רבעון 4 (אוקטובר–דצמבר)</option>
        </select>
      </div>
      <button class="btn btn-excel" onclick="doExportQuarter()">⬇️ ייצא דוח רבעוני</button>`;
    openModal(html);
    // Default to current quarter
    const curQ=Math.ceil((now.getMonth()+1)/3);
    if(document.getElementById('period-quarter'))document.getElementById('period-quarter').value=curQ;
    return;
  }
}
function doExportYear(){
  const year=parseInt(document.getElementById('period-year')?.value||new Date().getFullYear());
  closeModal();
  _runPeriodExport('year',year,null);
}
function doExportQuarter(){
  const year=parseInt(document.getElementById('period-year')?.value||new Date().getFullYear());
  const quarter=parseInt(document.getElementById('period-quarter')?.value||1);
  closeModal();
  _runPeriodExport('quarter',year,quarter);
}
function _runPeriodExport(type,year,quarter){
  let months=[];
  if(type==='year'){
    for(let m=0;m<12;m++)months.push(`${year}-${String(m+1).padStart(2,'0')}`);
  }else{
    const startM=(quarter-1)*3;
    for(let m=startM;m<startM+3;m++)months.push(`${year}-${String(m+1).padStart(2,'0')}`);
  }
  const users=_users.filter(u=>u.role==='physio');
  const cats=['assessment','followup','general','non-clinical'];
  const wb=XLSX.utils.book_new();
  users.forEach(u=>{
    const rows=months.map(mk=>{
      const [y,mo]=mk.split('-').map(Number);
      const target=calcMonthlyTarget(y,mo-1);
      const ul=_logs.filter(l=>l.userId===u.id&&l.date.startsWith(mk));
      const row={'חודש':MHE[mo-1]+' '+y,'יעד':target};
      let total=0;
      cats.forEach(cat=>{const pts=ul.filter(l=>getRuleById(l.ruleId).category===cat).reduce((s,l)=>s+parseFloat(l.points||0),0);row[catHe(cat)]=Math.round(pts*100)/100;total+=pts;});
      row['סה"כ']=Math.round(total*100)/100;
      row['% השגה']=Math.min(100,Math.round(total/target*100))+'%';
      row['שכר (₪)']=calcPay(u,total,target);
      return row;
    });
    const ws=XLSX.utils.json_to_sheet(rows);
    ws['!sheetView']=[{rightToLeft:true}];
    XLSX.utils.book_append_sheet(wb,ws,u.name.slice(0,31));
  });
  // Cross-user summary
  const summaryRows=users.map(u=>{
    const row={'שם':u.name};
    let grandTotal=0;
    months.forEach(mk=>{
      const [y,mo]=mk.split('-').map(Number);
      const ul=_logs.filter(l=>l.userId===u.id&&l.date.startsWith(mk));
      const pts=ul.reduce((s,l)=>s+parseFloat(l.points||0),0);
      row[MHE[mo-1]+"'"+(String(y).slice(2))]=Math.round(pts*100)/100;
      grandTotal+=pts;
    });
    row['סה"כ']=Math.round(grandTotal*100)/100;
    return row;
  });
  const wsSummary=XLSX.utils.json_to_sheet(summaryRows);
  wsSummary['!sheetView']=[{rightToLeft:true}];
  XLSX.utils.book_append_sheet(wb,wsSummary,'סיכום');
  const label=type==='quarter'?`רבעון${quarter}_${year}`:`שנתי_${year}`;
  downloadXlsx(wb,`PhysioControl_${label}.xlsx`);
}
