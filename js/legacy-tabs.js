import { euro, num } from "./calculos.js";
import {
  getWorkbook, findIngresoDate, loadAhorroRow, findAhorroDate,
  nextFreeAhorroRow, contabilidadPeriod
} from "./cloud-data.js";

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
const miniRow=(label,value,money=false)=>`<div class="legacy-row"><span>${label}</span><strong>${money?euro(value):count(value)}</strong></div>`;

function incomeHtml(){
  return `<div class="feature-grid">
    <div class="card feature-card income-card">
      <div class="card-head"><div><span class="eyebrow">EFECTIVO</span><h2>Ingreso del día</h2></div><span class="row-pill">Fila <b id="ingresoRow">—</b></span></div>
      <div class="bill-grid" id="ingresoInputs">${INCOME_INPUTS.map(([key,label])=>`<label><span>${label}</span><input data-income="${key}" inputmode="numeric" placeholder="0"></label>`).join("")}</div>
      <div class="actions"><button class="btn primary" id="loadIngreso">Cargar desde Excel</button><button class="btn" id="saveIngresoDraft">Guardar borrador local</button></div>
      <p class="notice">La lectura ya usa la hoja Ingreso real. El borrador no modifica todavía el Excel de OneDrive.</p>
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
      <div class="actions"><button class="btn primary" id="loadAhorro">Cargar fila</button><button class="btn" id="saveAhorroDraft">Guardar borrador local</button></div>
      <p class="notice">La lectura y búsqueda usan la hoja Ahorro real. El borrador permanece en este dispositivo.</p>
    </div>
    <div class="card feature-card">
      <div class="card-head"><div><span class="eyebrow">RESUMEN</span><h2>Ahorro</h2></div></div>
      <h3 class="section-title">Totales</h3><div class="legacy-table" id="ahorroTotals"></div>
      <h3 class="section-title">Restante</h3><div class="legacy-table" id="ahorroRest"></div>
    </div>
  </div>`;
}

function analysisHtml(){
  return `<div class="analysis-intro card"><div><span class="eyebrow">CENTRO DE ANÁLISIS</span><h2>Acumulados, medias e informes</h2><p>Consulta un mes o un rango y compara la media por día trabajado.</p></div><span class="analysis-badge">ANÁLISIS</span></div>
  <div class="card feature-card analysis-card">
    <div class="period-controls"><input id="monthView" class="mini-input" placeholder="mm/aaaa"><button class="btn primary" id="viewMonth">Ver mes</button><input id="rangeStart" type="date" class="mini-input"><input id="rangeEnd" type="date" class="mini-input"><button class="btn" id="filterRange">Filtrar rango</button><button class="btn" id="printAnalysis">Imprimir / PDF</button></div>
    <h2 id="periodTitle">Acumulado</h2>
    <div class="settlement-panel">
      <label><span>Nómina editable</span><input id="pdfNomina" value="1.365,46" inputmode="decimal"></label>
      <div><span>Total jefe</span><strong id="settlementChief">0,00 €</strong></div>
      <div class="final"><span>A percibir</span><strong id="settlementReceive">1.365,46 €</strong></div>
    </div>
    <div class="cash-grid">${[50,20,10,5,1,.5,.2,.1].map(v=>`<label><span>${String(v).replace(".",",")} €</span><input class="cash-qty" data-denom="${v}" value="0" inputmode="numeric"></label>`).join("")}</div>
    <div class="cash-summary"><span>Total del sobre <strong id="cashTotal">0,00 €</strong></span><span>Diferencia <strong id="cashDiff">—</strong></span></div>
    <div class="stats-row"><div><span>Días con fecha</span><strong id="statDays">—</strong></div><div><span>Días trabajados</span><strong id="statWork">—</strong></div><div><span>Días descanso</span><strong id="statRest">—</strong></div></div>
    <div class="period-grid"><section><span class="eyebrow">TOTAL</span><h3>Total del período</h3><div class="legacy-table" id="monthTotals"></div></section><section><span class="eyebrow">MEDIA</span><h3>Media por día trabajado</h3><div class="legacy-table" id="monthAverages"></div></section></div>
  </div>`;
}

function renderIncomeMonth(target,data){
  $(target).innerHTML=CASH_MONTH.map(([key,label])=>miniRow(label,data?.[key],key==="mes_total_efectivo")).join("");
}

export function initLegacyTabs({ensureWorkbook,toast,getDate}){
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
  $("saveIngresoDraft").onclick=()=>{
    const data={}; document.querySelectorAll("[data-income]").forEach(i=>data[i.dataset.income]=i.value);
    localStorage.setItem(`taxiIngresoDraft:${getDate()}`,JSON.stringify(data)); toast("Borrador de Ingreso guardado en este dispositivo");
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
  document.querySelectorAll("[data-saving]").forEach(input=>input.addEventListener("input",()=>{
    const denoms={b100:100,b50:50,b20:20,b10:10,b5:5};
    const total=Object.entries(denoms).reduce((sum,[key,value])=>sum+num(document.querySelector(`[data-saving="${key}"]`).value)*value,0);
    $("ahorroIngreso").textContent=euro(total);
  }));

  let chiefTotal=0;
  const updateSettlement=()=>{
    const payroll=num($("pdfNomina").value), receive=payroll+chiefTotal;
    $("settlementChief").textContent=euro(chiefTotal);$("settlementReceive").textContent=euro(receive);
    const cash=[...document.querySelectorAll(".cash-qty")].reduce((sum,i)=>sum+num(i.value)*num(i.dataset.denom),0);
    $("cashTotal").textContent=euro(cash);$("cashDiff").textContent=euro(receive-cash);
  };
  const renderPeriod=period=>{
    $("periodTitle").textContent=period.label;
    $("statDays").textContent=period.stats.days;$("statWork").textContent=period.stats.workdays;$("statRest").textContent=period.stats.restdays;
    $("monthTotals").innerHTML=PERIOD_FIELDS.map(([key,label])=>resultRow(label,period.totals[key])).join("");
    $("monthAverages").innerHTML=PERIOD_FIELDS.map(([key,label])=>resultRow(label,period.averages[key]??0)).join("");
    chiefTotal=num(period.totals.mes_total_jefe);updateSettlement();
  };
  const ensureAnalysis=async()=>getWorkbook() || await ensureWorkbook(false);
  $("viewMonth").onclick=async()=>{if(!await ensureAnalysis())return;try{renderPeriod(contabilidadPeriod({month:$("monthView").value}));toast("Análisis mensual cargado");}catch(error){toast(error.message);}};
  $("filterRange").onclick=async()=>{if(!await ensureAnalysis())return;try{renderPeriod(contabilidadPeriod({start:$("rangeStart").value,end:$("rangeEnd").value}));toast("Rango aplicado");}catch(error){toast(error.message);}};
  $("printAnalysis").onclick=()=>window.print();
  $("pdfNomina").oninput=updateSettlement;document.querySelectorAll(".cash-qty").forEach(i=>i.oninput=updateSettlement);
  const now=getDate().split("-");$("monthView").value=`${now[1]}/${now[0]}`;$("rangeStart").value=getDate();$("rangeEnd").value=getDate();

  return {loadIncome,loadSaving,loadAnalysis:()=>$("viewMonth").click()};
}
