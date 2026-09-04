import { euro, num } from "./calculos.js?v=20260903a";
import {
  getWorkbook, findIngresoDate, loadAhorroRow, findAhorroDate,
  nextFreeAhorroRow, contabilidadPeriod
} from "./cloud-data.js?v=20260904c";

const $=id=>document.getElementById(id);
const INCOME_INPUTS=[["b100","100 €",100],["b50","50 €",50],["b20","20 €",20],["b10","10 €",10],["b5","5 €",5]];
const INCOME_OUTPUTS=[["total","Total"],["gasto","Gasto"],["ingresos","Ingresos"],["me_queda","Me queda"],["a_deber","A deber"]];
const CASH_MONTH=[["mes_b100","100 €"],["mes_b50","50 €"],["mes_b20","20 €"],["mes_b10","10 €"],["mes_b5","5 €"],["mes_total_efectivo","Total efectivo"]];
const SAVING_INPUTS=[["date","Fecha"],["b100","100 €"],["b50","50 €"],["b20","20 €"],["b10","10 €"],["b5","5 €"],["gasto","Gasto / nota"],["detalle","Detalle / total"]];
const PERIOD_FIELDS=[
  ["mes_suma_total","Suma total"],["mes_cierre_pidetaxi","Cierre PideTaxi"],["mes_abonados_sin_comision","Abonados sin comisión"],
  ["mes_uber","Uber"],["mes_ubercash","Uber Cash"],["mes_joinup_bruto","JOIN UP bruto"],["mes_imbric_bruto","IMBRIC bruto"],
  ["mes_total_abonados_neto","Total abonados neto"],["mes_total_comisiones","Total comisiones"],["mes_mitad","50 %"],
  ["mes_seguro","Seguro"],["mes_cobro_tarjeta","Cobro tarjeta"],["mes_gasolina_lavado","Gasolina / lavado"],
  ["mes_total_jefe","Total jefe"],["mes_total_jefe_seguro","Total jefe 50 % + seguro"],["mes_me_queda","Me queda"]
];

const isoToDisplay=iso=>{
  const match=String(iso||"").match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return match ? `${match[3]}/${match[2]}/${match[1]}` : "";
};
const displayToIso=value=>{
  const text=String(value||"").trim();
  let match=text.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
  if(match) return `${match[3]}-${match[2].padStart(2,"0")}-${match[1].padStart(2,"0")}`;
  match=text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  return match ? `${match[1]}-${match[2].padStart(2,"0")}-${match[3].padStart(2,"0")}` : "";
};
const count=value=>{
  if(value==="" || value===null || value===undefined || String(value).trim().toLowerCase()==="descanso") return "—";
  return String(Math.round(num(value)));
};
const resultRow=(label,value)=>`<div class="legacy-row"><span>${label}</span><strong class="${num(value)<0?"negative":""}">${typeof value==="string"&&value.trim().toLowerCase()==="descanso"?"Descanso":euro(value)}</strong></div>`;
const periodTone=label=>({
  "Suma total":"period-total",
  "Cierre PideTaxi":"period-pidetaxi",
  "Uber":"period-uber",
  "Cobro tarjeta":"period-card",
  "Total jefe":"period-chief",
  "Me queda":"period-remains"
})[label]||"";
const periodRow=(label,value)=>`<div class="legacy-row ${periodTone(label)}"><span>${label}</span><strong class="${num(value)<0?"negative":""}">${euro(value)}</strong></div>`;
const miniRow=(label,value,money=false)=>`<div class="legacy-row"><span>${label}</span><strong>${money?euro(value):count(value)}</strong></div>`;

function incomeHtml(){
  return `<div class="feature-grid">
    <div class="card feature-card income-card">
      <div class="card-head"><div><span class="eyebrow">EFECTIVO</span><h2>Ingreso del día</h2></div><span class="row-pill">Fila <b id="ingresoRow">—</b></span></div>
      <div class="bill-grid" id="ingresoInputs">${INCOME_INPUTS.map(([key,label])=>`<label><span>${label}</span><input data-income="${key}" inputmode="numeric" placeholder="0"></label>`).join("")}</div>
      <div class="actions"><button class="btn primary" id="loadIngreso">Cargar desde Excel</button><button class="btn primary" id="saveIngresoCloud">Guardar en OneDrive</button><button class="btn" id="loadIngresoDraft">Recuperar borrador</button><button class="btn" id="saveIngresoDraft">Guardar borrador local</button></div>
      <p class="notice">Guardar en OneDrive escribe únicamente los cinco recuentos editables de la fila mostrada.</p>
    </div>
    <div class="card feature-card">
      <div class="card-head"><div><span class="eyebrow">RESULTADOS</span><h2>Resultado del ingreso</h2></div></div>
      <div class="legacy-table" id="ingresoOutputs">${INCOME_OUTPUTS.map(([,label])=>resultRow(label,0)).join("")}</div>
    </div>
  </div>
  <div class="card feature-card monthly-feature">
    <div class="card-head"><div><span class="eyebrow">ACUMULADO</span><h2 id="ingresoMonthTitle">Acumulado del mes</h2></div><span class="row-pill">Días <b id="ingresoMonthDays">—</b></span></div>
    <div class="income-columns">
      <section><h3>Remanente anterior</h3><p id="remainderLabel">Sin remanente detectado</p><div class="legacy-table" id="incomeRemainder"></div></section>
      <section class="current"><h3>Mes en curso</h3><p>Movimiento del mes seleccionado</p><div class="legacy-table" id="incomeMonth"></div></section>
      <section class="combined"><h3>Remanente + mes</h3><p>Total disponible combinado</p><div class="legacy-table" id="incomeCombined"></div></section>
    </div>
  </div>`;
}

function savingHtml(){
  return `<div class="feature-grid">
    <div class="card feature-card saving-card">
      <div class="card-head"><div><span class="eyebrow">AHORRO</span><h2>Registro</h2></div><label class="row-pill">Fila <input id="ahorroRow" type="number" min="2" max="16" value="2"></label></div>
      <div class="inline-tools"><input id="ahorroSearchDate" class="mini-input" placeholder="dd/mm/aaaa"><button class="btn" id="findAhorro">Buscar fecha</button><button class="btn" id="freeAhorro">Siguiente libre</button></div>
      <div class="saving-grid" id="ahorroInputs">${SAVING_INPUTS.map(([key,label])=>`<label><span>${label}</span><input data-saving="${key}" ${key==="date"?'placeholder="dd/mm/aaaa"':'inputmode="decimal"'}></label>`).join("")}</div>
      <div class="single-result"><span>Ingreso calculado</span><strong id="ahorroIngreso">0,00 €</strong></div>
      <div class="actions"><button class="btn primary" id="loadAhorro">Cargar fila</button><button class="btn primary" id="saveAhorroCloud">Guardar en OneDrive</button><button class="btn" id="loadAhorroDraft">Recuperar borrador</button><button class="btn" id="saveAhorroDraft">Guardar borrador local</button></div>
      <p class="notice">Guardar en OneDrive conserva la fórmula de Ingreso y escribe solo fecha, recuentos, gasto y detalle.</p>
    </div>
    <div class="card feature-card">
      <div class="card-head"><div><span class="eyebrow">RESUMEN</span><h2>Ahorro</h2></div></div>
      <h3 class="section-title">Totales</h3><div class="legacy-table" id="ahorroTotals"></div>
      <h3 class="section-title">Restante</h3><div class="legacy-table" id="ahorroRest"></div>
    </div>
  </div>`;
}

function analysisHtml(){
  return `<div class="analysis-intro card"><div><span class="eyebrow">CENTRO DE ANÁLISIS</span><h2>Acumulados, medias e informes</h2><p>Consulta un mes o un rango, compara la media por día trabajado, prepara la liquidación y exporta tus informes PDF.</p></div><span class="analysis-badge">ANÁLISIS</span></div>
  <div class="card feature-card analysis-card">
    <div class="period-controls"><input id="monthView" class="mini-input" placeholder="mm/aaaa"><button class="btn primary" id="viewMonth">Ver mes</button><button class="btn" id="monthOfDate">Mes de la fecha</button><span class="period-divider"></span><input id="rangeStart" type="date" class="mini-input"><input id="rangeEnd" type="date" class="mini-input"><button class="btn accent" id="filterRange">Filtrar rango</button></div>
    <h2 id="periodTitle">Acumulado</h2>
    <div class="pdf-export-zone">
      <div class="pdf-export-card"><div><span class="eyebrow">DETALLADO</span><strong>Días trabajados del período</strong><small>Incluye únicamente las columnas con movimiento.</small></div><button class="btn detail-pdf" id="exportDetailPdf">▤ Exportar PDF detallado</button></div>
      <div class="pdf-export-card"><div><span class="eyebrow">ACUMULADO</span><strong>Totales y medias del período</strong><small>Oculta automáticamente las líneas sin movimiento.</small><label class="pdf-option"><input id="includeSettlementPdf" type="checkbox"> Incluir referencia de liquidación</label></div><button class="btn accumulated-pdf" id="exportPdf">↓ Exportar PDF acumulado</button></div>
    </div>
    <div class="settlement-wrap">
      <div class="settlement-heading"><span class="eyebrow">LIQUIDACIÓN</span><h3>Datos para la liquidación del período</h3><p>La nómina y las cantidades son editables. Total jefe y A percibir se calculan automáticamente.</p></div>
      <div class="settlement-panel">
        <label><span>Importe de la nómina</span><input id="pdfNomina" value="1.365,46" inputmode="decimal"></label>
        <div><span>Total jefe del período</span><strong id="settlementChief">0,00 €</strong><small>Se obtiene del acumulado seleccionado.</small></div>
        <div class="final"><span>A percibir</span><strong id="settlementReceive">1.365,46 €</strong><small>Nómina + Total jefe.</small></div>
      </div>
      <div class="cash-heading"><div><strong>Desglose del efectivo</strong><small>Introduce la cantidad de cada billete o moneda de la liquidación en efectivo.</small></div><span> CANTIDADES EDITABLES </span></div>
    </div>
    <div class="cash-grid">${[50,20,10,5,1,.5,.2,.1].map(v=>`<label><span>${String(v).replace(".",",")} €</span><small>Cantidad</small><input class="cash-qty" data-denom="${v}" value="0" min="0" inputmode="numeric"></label>`).join("")}</div>
    <div class="cash-summary"><span>Total del sobre <strong id="cashTotal">0,00 €</strong></span><span>Diferencia <strong id="cashDiff">—</strong></span></div>
    <div class="stats-row"><div><span>Días con fecha</span><strong id="statDays">—</strong></div><div><span>Días trabajados</span><strong id="statWork">—</strong></div><div><span>Días descanso</span><strong id="statRest">—</strong></div></div>
    <div class="period-grid"><section><span class="eyebrow">TOTAL</span><h3>Total del período</h3><div class="legacy-table" id="monthTotals"></div></section><section><span class="eyebrow">MEDIA</span><h3>Media por día trabajado</h3><div class="legacy-table" id="monthAverages"></div></section></div>
  </div>`;
}

function renderIncomeMonth(target,data){
  $(target).innerHTML=CASH_MONTH.map(([key,label])=>miniRow(label,data?.[key],key==="mes_total_efectivo")).join("");
}

export function initLegacyTabs({ensureWorkbook,toast,getDate,saveIngreso,saveAhorro}){
  $("tab-ingreso").innerHTML=incomeHtml();
  $("tab-ahorro").innerHTML=savingHtml();
  $("tab-analisis").innerHTML=analysisHtml();

  const fillIncome=result=>{
    $("ingresoRow").textContent=result.row;
    document.querySelectorAll("[data-income]").forEach(input=>input.value=result.inputs[input.dataset.income]??"");
    $("ingresoOutputs").innerHTML=INCOME_OUTPUTS.map(([key,label])=>resultRow(label,result.outputs[key])).join("");
    const month=result.month;
    $("ingresoMonthTitle").textContent=month.label;
    $("ingresoMonthDays").textContent=month.days;
    $("remainderLabel").textContent=month.remainder_row?`Remanente anterior · fila ${month.remainder_row}`:"Sin remanente anterior detectado";
    renderIncomeMonth("incomeRemainder",month.remainder);
    renderIncomeMonth("incomeMonth",month.month);
    renderIncomeMonth("incomeCombined",month.combined);
  };
  const loadIncome=async()=>{
    if(!getWorkbook() && !(await ensureWorkbook(false))) return;
    const result=findIngresoDate(getDate());
    if(!result){toast("No se encontró esa fecha en la hoja Ingreso");return;}
    fillIncome(result); toast(`Ingreso cargado desde la fila ${result.row}`);
  };
  $("loadIngreso").onclick=loadIncome;
  $("loadIngresoDraft").onclick=()=>{
    const saved=localStorage.getItem(`taxiIngresoDraft:${getDate()}`);
    if(!saved){toast("No hay borrador de Ingreso para esta fecha");return;}
    try{
      const data=JSON.parse(saved);
      document.querySelectorAll("[data-income]").forEach(i=>i.value=data[i.dataset.income]??"");
      document.querySelector("[data-income]")?.dispatchEvent(new Event("input"));
      toast("Borrador de Ingreso recuperado");
    }catch{toast("El borrador de Ingreso no se pudo leer");}
  };
  $("saveIngresoDraft").onclick=()=>{
    const data={}; document.querySelectorAll("[data-income]").forEach(i=>data[i.dataset.income]=i.value);
    localStorage.setItem(`taxiIngresoDraft:${getDate()}`,JSON.stringify(data)); toast("Borrador de Ingreso guardado en este dispositivo");
  };
  $("saveIngresoCloud").onclick=async()=>{
    if(!saveIngreso){toast("El guardado de Ingreso todavía no está configurado.");return;}
    if(!getWorkbook() && !(await ensureWorkbook(false))) return;
    const result=findIngresoDate(getDate());
    if(!result){toast("No existe una fila para esta fecha en la hoja Ingreso.");return;}
    const values={};
    document.querySelectorAll("[data-income]").forEach(i=>values[i.dataset.income]=num(i.value));
    try{await saveIngreso({fila:result.row,fecha:getDate(),valores:values});toast(`Ingreso guardado en OneDrive · fila ${result.row}`);}
    catch(error){toast(`No se ha guardado Ingreso: ${error.message}`);}
  };
  document.querySelectorAll("[data-income]").forEach(input=>input.addEventListener("input",()=>{
    const total=INCOME_INPUTS.reduce((sum,[key,,denom])=>sum+num(document.querySelector(`[data-income="${key}"]`).value)*denom,0);
    const values={total,gasto:0,ingresos:total,me_queda:0,a_deber:-total};
    $("ingresoOutputs").innerHTML=INCOME_OUTPUTS.map(([key,label])=>resultRow(label,values[key])).join("");
  }));

  const fillSaving=result=>{
    $("ahorroRow").value=result.row;
    document.querySelectorAll("[data-saving]").forEach(input=>{
      let value=result.inputs[input.dataset.saving]??"";
      if(input.dataset.saving==="date") value=isoToDisplay(value);
      input.value=value;
    });
    $("ahorroSearchDate").value=isoToDisplay(result.inputs.date);
    $("ahorroIngreso").textContent=euro(result.ingreso);
    const s=result.summary;
    $("ahorroTotals").innerHTML=[["100 €",s.tot_100],["50 €",s.tot_50],["20 €",s.tot_20],["10 €",s.tot_10],["5 €",s.tot_5],["Ingreso",s.tot_ingreso,true]].map(v=>miniRow(...v)).join("");
    $("ahorroRest").innerHTML=[["100 €",s.rest_100],["50 €",s.rest_50],["20 €",s.rest_20],["10 €",s.rest_10],["5 €",s.rest_5],["Ingreso",s.rest_ingreso,true]].map(v=>miniRow(...v)).join("");
  };
  const loadSaving=async row=>{
    if(!getWorkbook() && !(await ensureWorkbook(false))) return;
    try{const result=loadAhorroRow(row??$("ahorroRow").value);fillSaving(result);toast(`Ahorro cargado desde la fila ${result.row}`);}catch(error){toast(error.message);}
  };
  $("loadAhorro").onclick=()=>loadSaving();
  $("findAhorro").onclick=async()=>{
    if(!getWorkbook() && !(await ensureWorkbook(false))) return;
    const iso=displayToIso($("ahorroSearchDate").value);
    if(!iso){toast("Fecha no válida. Usa dd/mm/aaaa.");return;}
    const result=findAhorroDate(iso);
    if(!result){toast("No se encontró esa fecha en Ahorro");return;}
    fillSaving(result);toast(`Fecha encontrada en la fila ${result.row}`);
  };
  $("freeAhorro").onclick=async()=>{
    if(!getWorkbook() && !(await ensureWorkbook(false))) return;
    const row=nextFreeAhorroRow();
    if(!row){toast("No hay filas libres entre la 2 y la 16");return;}
    $("ahorroRow").value=row;document.querySelectorAll("[data-saving]").forEach(i=>i.value="");
    document.querySelector('[data-saving="date"]').value=isoToDisplay(getDate());toast(`Siguiente fila libre: ${row}`);
  };
  $("saveAhorroDraft").onclick=()=>{
    const data={row:$("ahorroRow").value};document.querySelectorAll("[data-saving]").forEach(i=>data[i.dataset.saving]=i.value);
    localStorage.setItem(`taxiAhorroDraft:${data.row}`,JSON.stringify(data));toast("Borrador de Ahorro guardado en este dispositivo");
  };
  $("saveAhorroCloud").onclick=async()=>{
    if(!saveAhorro){toast("El guardado de Ahorro todavía no está configurado.");return;}
    if(!getWorkbook() && !(await ensureWorkbook(false))) return;
    const row=Number($("ahorroRow").value);
    if(!Number.isInteger(row)||row<2||row>16){toast("La fila de Ahorro debe estar entre la 2 y la 16.");return;}
    const values={}; document.querySelectorAll("[data-saving]").forEach(i=>values[i.dataset.saving]=i.value.trim());
    const iso=displayToIso(values.date);
    if(!iso){toast("Fecha no válida. Usa dd/mm/aaaa.");return;}
    ["b100","b50","b20","b10","b5"].forEach(key=>values[key]=num(values[key]));
    try{await saveAhorro({fila:row,fecha:iso,valores});toast(`Ahorro guardado en OneDrive · fila ${row}`);}
    catch(error){toast(`No se ha guardado Ahorro: ${error.message}`);}
  };
  $("loadAhorroDraft").onclick=()=>{
    const row=$("ahorroRow").value;
    const saved=localStorage.getItem(`taxiAhorroDraft:${row}`);
    if(!saved){toast(`No hay borrador de Ahorro para la fila ${row}`);return;}
    try{
      const data=JSON.parse(saved);
      document.querySelectorAll("[data-saving]").forEach(i=>i.value=data[i.dataset.saving]??"");
      document.querySelector('[data-saving="date"]')?.dispatchEvent(new Event("input"));
      toast("Borrador de Ahorro recuperado");
    }catch{toast("El borrador de Ahorro no se pudo leer");}
  };
  document.querySelectorAll("[data-saving]").forEach(input=>input.addEventListener("input",()=>{
    const denoms={b100:100,b50:50,b20:20,b10:10,b5:5};
    const total=Object.entries(denoms).reduce((sum,[key,value])=>sum+num(document.querySelector(`[data-saving="${key}"]`).value)*value,0);
    $("ahorroIngreso").textContent=euro(total);
  }));

  let chiefTotal=0, currentPeriod=null;
  const updateSettlement=()=>{
    const payroll=num($("pdfNomina").value), receive=payroll+chiefTotal;
    $("settlementChief").textContent=euro(chiefTotal);$("settlementReceive").textContent=euro(receive);
    const cash=[...document.querySelectorAll(".cash-qty")].reduce((sum,i)=>sum+num(i.value)*num(i.dataset.denom),0);
    $("cashTotal").textContent=euro(cash);$("cashDiff").textContent=euro(receive-cash);
  };
  const renderPeriod=period=>{
    currentPeriod=period;
    $("periodTitle").textContent=period.label;
    $("statDays").textContent=period.stats.days;$("statWork").textContent=period.stats.workdays;$("statRest").textContent=period.stats.restdays;
    $("monthTotals").innerHTML=PERIOD_FIELDS.map(([key,label])=>periodRow(label,period.totals[key])).join("");
    $("monthAverages").innerHTML=PERIOD_FIELDS.map(([key,label])=>periodRow(label,period.averages[key]??0)).join("");
    chiefTotal=num(period.totals.mes_total_jefe);updateSettlement();
  };
  const ensureAnalysis=async()=>getWorkbook() || await ensureWorkbook(false);
  $("viewMonth").onclick=async()=>{if(!await ensureAnalysis())return;try{renderPeriod(contabilidadPeriod({month:$("monthView").value}));toast("Análisis mensual cargado");}catch(error){toast(error.message);}};
  $("filterRange").onclick=async()=>{if(!await ensureAnalysis())return;try{renderPeriod(contabilidadPeriod({start:$("rangeStart").value,end:$("rangeEnd").value}));toast("Rango aplicado");}catch(error){toast(error.message);}};
  $("monthOfDate").onclick=()=>{
    const parts=getDate().split("-");
    $("monthView").value=`${parts[1]}/${parts[0]}`;
    $("viewMonth").click();
  };
  const nonEmptyPeriodRows=kind=>{
    if(!currentPeriod) return [];
    const source=kind==="average"?currentPeriod.averages:currentPeriod.totals;
    return PERIOD_FIELDS.filter(([key])=>Math.abs(num(source[key]))>.0001).map(([key,label])=>[label,source[key]]);
  };
  const pdfDate=()=>new Date().toLocaleDateString("es-ES");
  const filePeriod=()=>String(currentPeriod?.label||"periodo").replace(/[^0-9a-záéíóúñ]+/gi,"-").replace(/^-|-$/g,"");
  const loadPdf=()=>new Promise((resolve,reject)=>{
    if(window.jspdf?.jsPDF) return resolve(window.jspdf.jsPDF);
    const existing=document.getElementById("jspdf-cdn");
    if(existing){existing.addEventListener("load",()=>resolve(window.jspdf.jsPDF),{once:true});existing.addEventListener("error",reject,{once:true});return;}
    const script=document.createElement("script");script.id="jspdf-cdn";script.src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js";
    script.onload=()=>window.jspdf?.jsPDF?resolve(window.jspdf.jsPDF):reject(new Error("No se pudo cargar el generador PDF"));script.onerror=()=>reject(new Error("No se pudo cargar el generador PDF"));document.head.appendChild(script);
  });
  const exportPdf=async detailed=>{
    if(!currentPeriod){toast("Primero carga un mes o un rango para exportar");return;}
    try{
      const JsPDF=await loadPdf(),doc=new JsPDF({unit:"mm",format:"a4",orientation:detailed?"landscape":"portrait"});
      const margin=16,pageWidth=detailed?297:210,width=detailed?265:178;let y=18;
      const title=detailed?"Contabilidad Taxi · Informe detallado":"Contabilidad Taxi · Informe acumulado";
      doc.setFillColor(13,35,53);doc.rect(0,0,pageWidth,33,"F");doc.setTextColor(255,255,255);doc.setFontSize(17);doc.text(title,margin,15);doc.setFontSize(10);doc.text(currentPeriod.label,margin,23);doc.text(`Generado: ${pdfDate()}`,margin,28);
      doc.setTextColor(20,42,58);doc.setFontSize(12);y=44;
      const addLine=(label,value)=>{if(y>278){doc.addPage();y=18;}doc.setFontSize(10);doc.text(label,margin,y);doc.text(euro(value),margin+width,y,{align:"right"});doc.setDrawColor(215,225,232);doc.line(margin,y+2,margin+width,y+2);y+=7;};
      if(detailed){
        const fields=(currentPeriod.detailFields||[]).filter(field=>currentPeriod.detailRows.some(row=>Math.abs(num(row.values[field.key]))>.0001));
        const excluded=(currentPeriod.detailFields||[]).filter(field=>!fields.includes(field)).map(field=>field.label);
        const allColumns=[{key:"date",label:"Fecha"},{key:"day",label:"Día"},...fields];
        const columnWidth=width/allColumns.length,headerHeight=10,rowHeight=6;
        const drawDetailHeader=at=>{doc.setFillColor(33,86,123);doc.rect(margin,at,width,headerHeight,"F");doc.setTextColor(255,255,255);doc.setFontSize(5.7);allColumns.forEach((field,index)=>{const x=margin+index*columnWidth;const lines=doc.splitTextToSize(field.label,columnWidth-1);doc.text(lines,x+columnWidth/2,at+3.5,{align:"center"});});};
        y=40;drawDetailHeader(y);y+=headerHeight;
        currentPeriod.detailRows.forEach((row,index)=>{
          if(y>198){doc.addPage();y=16;drawDetailHeader(y);y+=headerHeight;}
          doc.setFillColor(index%2?249:238,index%2?251:246,index%2?253:249);doc.rect(margin,y,width,rowHeight,"F");doc.setDrawColor(215,225,232);doc.rect(margin,y,width,rowHeight,"S");doc.setTextColor(22,48,69);doc.setFontSize(5.9);
          allColumns.forEach((field,col)=>{const value=field.key==="date"?isoToDisplay(row.date):field.key==="day"?row.day:euro(row.values[field.key]);doc.text(String(value),margin+col*columnWidth+columnWidth/2,y+4,{align:"center"});});y+=rowHeight;
        });
        if(y>198){doc.addPage();y=16;drawDetailHeader(y);y+=headerHeight;}
        doc.setFillColor(224,238,247);doc.rect(margin,y,width,rowHeight,"F");doc.setDrawColor(97,139,166);doc.rect(margin,y,width,rowHeight,"S");doc.setTextColor(20,58,82);doc.setFontSize(6.2);doc.text("TOTAL",margin+columnWidth,y+4,{align:"center"});
        fields.forEach((field,index)=>{const total=currentPeriod.detailRows.reduce((sum,row)=>sum+num(row.values[field.key]),0);doc.text(euro(total),margin+(index+2)*columnWidth+columnWidth/2,y+4,{align:"center"});});y+=rowHeight;
        if(excluded.length){doc.setTextColor(84,105,121);doc.setFontSize(7);doc.text(`Sin movimientos en el período · excluidas: ${excluded.join(", ")}`,margin,y+5);y+=10;}
        if(y>180){doc.addPage();y=18;}
        const payroll=num($("pdfNomina").value),cash=[...document.querySelectorAll(".cash-qty")].reduce((s,i)=>s+num(i.value)*num(i.dataset.denom),0),receive=payroll+chiefTotal;
        doc.setFillColor(21,57,82);doc.roundedRect(margin,y,width,25,3,3,"F");doc.setTextColor(255,255,255);doc.setFontSize(9);doc.text("LIQUIDACIÓN",margin+4,y+7);
        [["Nómina",payroll],["Total jefe",chiefTotal],["A percibir",receive],["Efectivo",cash]].forEach(([label,value],index)=>{const x=margin+index*(width/4);doc.setTextColor(193,215,230);doc.setFontSize(7);doc.text(label,x+width/8,y+14,{align:"center"});doc.setTextColor(value<0?255:255,value<0?126:255,value<0?126:255);doc.setFontSize(12);doc.text(euro(value),x+width/8,y+21,{align:"center"});});
      }else{
        /* Maquetación equivalente al informe acumulado de la aplicación de escritorio. */
        const rangeText=currentPeriod.label.replace("Acumulado del mes ","Mes ").replace("Acumulado del rango ","Rango ");
        doc.setFillColor(21,57,82);doc.rect(0,0,210,30,"F");doc.setTextColor(255,255,255);doc.setFontSize(18);doc.text("Contabilidad Taxi - Acumulado",margin,14);doc.setFontSize(10);doc.text(rangeText,margin,22);
        const cards=[["Días con fecha",currentPeriod.stats.days,[230,240,250],[41,125,190]],["Días trabajados",currentPeriod.stats.workdays,[231,248,239],[33,139,91]],["Días de descanso",currentPeriod.stats.restdays,[253,242,230],[194,116,24]]];
        cards.forEach(([label,value,bg,ink],index)=>{const x=margin+index*60;doc.setFillColor(...bg);doc.roundedRect(x,37,55,20,3,3,"F");doc.setDrawColor(...ink);doc.roundedRect(x,37,55,20,3,3,"S");doc.setTextColor(75,99,118);doc.setFontSize(8);doc.text(label,x+3,43);doc.setTextColor(...ink);doc.setFontSize(15);doc.text(String(value),x+51,52,{align:"right"});});
        const rows=nonEmptyPeriodRows("total").filter(([label])=>label!=="Me queda").map(([label,total])=>[label,total,currentPeriod.averages[PERIOD_FIELDS.find(([,name])=>name===label)?.[0]]??0]);
        let tableY=67,rowH=8;doc.setFillColor(33,86,123);doc.roundedRect(margin,tableY,width,9,2,2,"F");doc.setTextColor(255,255,255);doc.setFontSize(9);doc.text("CONCEPTO",margin+3,tableY+6);doc.text("TOTAL DEL PERÍODO",margin+118,tableY+6,{align:"right"});doc.text("MEDIA / DÍA TRABAJADO",margin+width-3,tableY+6,{align:"right"});tableY+=9;
        rows.forEach(([label,total,average],index)=>{if(tableY>257){doc.addPage();tableY=18;}doc.setFillColor(index%2?248:238,index%2?250:245,index%2?252:248);doc.rect(margin,tableY,width,rowH,"F");doc.setDrawColor(211,224,233);doc.rect(margin,tableY,width,rowH,"S");doc.setTextColor(23,49,69);doc.setFontSize(9);doc.text(label,margin+3,tableY+5.4);doc.text(euro(total),margin+118,tableY+5.4,{align:"right"});doc.text(euro(average),margin+width-3,tableY+5.4,{align:"right"});tableY+=rowH;});
        if($("includeSettlementPdf").checked){
          const payroll=num($("pdfNomina").value),receive=payroll+chiefTotal;
          if(tableY>244){doc.addPage();tableY=18;}tableY+=8;doc.setFillColor(21,57,82);doc.roundedRect(margin,tableY,width,25,3,3,"F");doc.setTextColor(255,255,255);doc.setFontSize(10);doc.text("REFERENCIA DE LIQUIDACIÓN",margin+4,tableY+7);
          const values=[["Nómina",payroll],["Total jefe del período",chiefTotal],["A percibir",receive]];
          values.forEach(([label,value],index)=>{const x=margin+index*(width/3);if(index){doc.setDrawColor(95,133,159);doc.line(x,tableY+10,x,tableY+22);}doc.setTextColor(193,215,230);doc.setFontSize(8);doc.text(label,x+width/6,tableY+14,{align:"center"});doc.setTextColor(value<0?255:255,value<0?126:255,value<0?126:255);doc.setFontSize(15);doc.text(euro(value),x+width/6,tableY+21,{align:"center"});});
        }
      }
      doc.save(`contabilidad-taxi-${detailed?"detallado":"acumulado"}-${filePeriod()}.pdf`);
      toast(`PDF ${detailed?"detallado":"acumulado"} descargado`);
    }catch(error){toast(`${error.message}. Puedes usar Imprimir del navegador como alternativa.`);}
  };
  $("exportPdf").onclick=()=>exportPdf(false);
  $("exportDetailPdf").onclick=()=>exportPdf(true);
  $("pdfNomina").oninput=updateSettlement;document.querySelectorAll(".cash-qty").forEach(i=>i.oninput=updateSettlement);
  const now=getDate().split("-");$("monthView").value=`${now[1]}/${now[0]}`;$("rangeStart").value=getDate();$("rangeEnd").value=getDate();

  return {loadIncome,loadSaving,loadAnalysis:()=>$("viewMonth").click()};
}
