import { calcularDia, euro, num } from "./calculos.js?v=20260903a";
import { saveLocalDay, loadLocalDay, markLocalRest } from "./storage.js?v=20260902c";
import {
  getWebhookUrl, getWorkbook, saveWebhookUrl, clearWebhookUrl, fetchWorkbook,
  workbookSummary, findContabilidadDate, nextPendingContabilidad,
  getWriteWebhookUrl, saveWriteWebhookUrl, clearWriteWebhookUrl, saveContabilidadValues
} from "./cloud-data.js?v=20260903b";
import { initLegacyTabs } from "./legacy-tabs.js?v=20260903a";

const $=id=>document.getElementById(id);
const qsa=s=>[...document.querySelectorAll(s)];

function toast(msg){
  const el=$("toast"); el.textContent=msg; el.classList.add("show");
  clearTimeout(window.__toast); window.__toast=setTimeout(()=>el.classList.remove("show"),3000);
}
function isoToday(){
  const d=new Date(); d.setMinutes(d.getMinutes()-d.getTimezoneOffset());
  return d.toISOString().slice(0,10);
}
function collectConta(){
  const o={}; qsa("[data-field]").forEach(i=>o[i.dataset.field]=i.value); return o;
}
function fillConta(data={}){
  qsa("[data-field]").forEach(i=>i.value=data[i.dataset.field]??"");
  updateCalc(data);
}
function updateCalc(source=null){
  const raw=source || collectConta();
  const r=calcularDia(raw);
  const value=(key)=>source && source[key]!=="" && source[key]!==undefined
    ? source[key]
    : r[key];
  $("rSuma").textContent = euro(value("suma_total"));
  $("rMitad").textContent = euro(value("mitad"));
  $("rComisiones").textContent = euro(value("total_comisiones"));
  $("rJefe").textContent = euro(value("total_jefe"));
}
function setCloudStatus(){
  const url=getWebhookUrl();
  const writeUrl=getWriteWebhookUrl();
  if($("makeWebhookUrl")) $("makeWebhookUrl").value=url;
  if($("makeWriteWebhookUrl")) $("makeWriteWebhookUrl").value=writeUrl;
  if($("makeState")) $("makeState").textContent=url?"Configurada en este dispositivo":"Sin configurar";
  if($("makeWriteState")) $("makeWriteState").textContent=writeUrl?"Preparado para guardar":"Pendiente de configurar";
  $("cloudDot")?.classList.toggle("online",!!url);
  if($("accountText")) $("accountText").textContent=url?(writeUrl?"Lectura y guardado configurados":"Make + OneDrive configurado"):"Nube sin configurar";
}
function renderSheets(list){
  const box=$("sheetSummary");
  if(!box)return;
  if(!list?.length){
    box.className="sheet-summary empty";
    box.textContent="No se detectaron hojas.";
    return;
  }
  box.className="sheet-summary";
  box.innerHTML=list.map(x=>`<div><strong>${x.name}</strong><span>${x.rows} filas · ${x.cols} columnas</span></div>`).join("");
}
async function loadCloudWorkbook(showToast=true){
  $("globalStatus").textContent="Conectando con OneDrive…";
  $("globalStatusDetail").textContent="Make está recuperando el Excel.";
  try{
    await fetchWorkbook();
    const list=workbookSummary()||[];
    $("workbookName").textContent="164_CONTA CALEN TAXI 164.xlsx";
    $("sheetCount").textContent=String(list.length);
    renderSheets(list);
    $("globalStatus").textContent="Excel conectado";
    $("globalStatusDetail").textContent="El libro de OneDrive está disponible en esta sesión.";
    if(showToast) toast("Excel cargado correctamente desde OneDrive");
    return true;
  }catch(e){
    $("globalStatus").textContent="No se pudo cargar el Excel";
    $("globalStatusDetail").textContent=e.message;
    if(showToast) toast(e.message);
    return false;
  }
}

$("workDate").value=isoToday();
const legacyTabs=initLegacyTabs({
  ensureWorkbook:loadCloudWorkbook,
  toast,
  getDate:()=>$("workDate").value
});

function changeWorkDate(days){
  const input=$("workDate"),date=new Date(`${input.value}T12:00:00`);
  if(Number.isNaN(date.getTime())) return;
  date.setDate(date.getDate()+days);
  input.value=date.toISOString().slice(0,10);
}
function addDailyControls(){
  const host=document.querySelector("#tab-contabilidad .card");
  if(!host || $("dailyControls")) return;
  const controls=document.createElement("div");
  controls.id="dailyControls";
  controls.className="daily-controls";
  controls.innerHTML=`<button class="btn" id="todayDate">● Hoy</button><button class="btn" id="previousDate">‹ Día</button><button class="btn" id="nextDate">Día ›</button><button class="btn" id="loadCurrentDay">↻ Cargar día</button><button class="btn primary" id="loadBundle">⇄ Contabilidad + Ingreso</button><button class="btn accent" id="nextPending">⚑ Siguiente pendiente</button>`;
  host.querySelector(".card-head")?.after(controls);
  $("todayDate").onclick=()=>{$("workDate").value=isoToday();toast("Fecha de hoy seleccionada");};
  $("previousDate").onclick=()=>changeWorkDate(-1);
  $("nextDate").onclick=()=>changeWorkDate(1);
  $("loadCurrentDay").onclick=()=>$("loadDay").click();
  $("loadBundle").onclick=async()=>{ $("loadDay").click(); await legacyTabs.loadIncome(); };
  $("nextPending").onclick=async()=>{
    if(!getWebhookUrl()){toast("Primero guarda la URL privada del webhook de Make.");return;}
    if(!getWorkbook() && !await loadCloudWorkbook(false)) return;
    const pending=nextPendingContabilidad($("workDate").value);
    if(!pending){toast("No hay días pendientes a partir de la fecha seleccionada.");return;}
    $("workDate").value=pending.date;
    toast(`Siguiente pendiente: ${pending.date} · fila ${pending.row}`);
  };
}
addDailyControls();

qsa(".tab").forEach(btn=>btn.addEventListener("click",async()=>{
  qsa(".tab").forEach(x=>x.classList.toggle("active",x===btn));
  qsa(".tabpage").forEach(p=>p.classList.toggle("active",p.id==="tab-"+btn.dataset.tab));
  if(btn.dataset.loaded)return;
  if(btn.dataset.tab==="ingreso") await legacyTabs.loadIncome();
  if(btn.dataset.tab==="ahorro") await legacyTabs.loadSaving();
  if(btn.dataset.tab==="analisis") await legacyTabs.loadAnalysis();
  btn.dataset.loaded="1";
}));
qsa("[data-field]").forEach(i=>i.addEventListener("input",()=>updateCalc()));

$("saveDay").addEventListener("click",async()=>{
  const date=$("workDate").value;
  const values=collectConta();
  saveLocalDay(date,{descanso:false,...values});
  if(!getWriteWebhookUrl()){
    toast("Guardado local. Configura el webhook de guardado para escribir en OneDrive.");
    return;
  }
  try{
    let row=findContabilidadDate(date);
    if(!row){
      const loaded=await loadCloudWorkbook(false);
      if(loaded) row=findContabilidadDate(date);
    }
    if(!row) throw new Error("No existe una fila para esta fecha en la hoja Contabilidad.");
    await saveContabilidadValues({
      fecha:date,
      fila:row.row,
      valores:{
        cierre_pidetaxi:num(values.cierre_pidetaxi),
        abonados_sin_comision:num(values.abonados_sin_comision),
        uber:num(values.uber),
        ubercash:num(values.ubercash),
        joinup_bruto:num(values.joinup_bruto),
        imbric_bruto:num(values.imbric_bruto),
        cobro_tarjeta:num(values.cobro_tarjeta),
        gasolina_lavado:num(values.gasolina_lavado)
      }
    });
    toast(`Guardado en OneDrive · fila ${row.row}`);
  }catch(e){
    toast(`No se ha guardado en OneDrive: ${e.message}`);
  }
});

$("loadDay").addEventListener("click",async()=>{
  const date=$("workDate").value;
  try{
    let row=findContabilidadDate(date);
    if(!row && getWebhookUrl()){
      const ok=await loadCloudWorkbook(false);
      if(ok) row=findContabilidadDate(date);
    }
    if(row){
      fillConta(row);
      toast(`Cargado desde Excel · fila ${row.row}`);
      return;
    }
  }catch(e){
    // Si aún no hay Excel Cloud cargado, probamos almacenamiento local.
  }
  const d=loadLocalDay(date);
  if(!d){toast("No hay datos para ese día en el Excel ni en el modo local");return}
  if(d.descanso){fillConta({});toast("Ese día está marcado como descanso");return}
  fillConta(d);toast("Día cargado del modo local");
});

$("markRest").addEventListener("click",()=>{
  markLocalRest($("workDate").value); fillConta({}); toast("Descanso guardado en modo local de prueba");
});

$("saveWebhook")?.addEventListener("click",()=>{
  try{
    saveWebhookUrl($("makeWebhookUrl").value);
    setCloudStatus();
    toast("Conexión privada guardada en este dispositivo");
  }catch(e){toast(e.message)}
});
$("clearWebhook")?.addEventListener("click",()=>{
  clearWebhookUrl(); setCloudStatus(); $("makeWebhookUrl").value="";
  toast("URL privada eliminada de este dispositivo");
});
$("saveWriteWebhook")?.addEventListener("click",()=>{
  try{
    saveWriteWebhookUrl($("makeWriteWebhookUrl").value);
    setCloudStatus();
    toast("Webhook de guardado guardado en este dispositivo");
  }catch(e){toast(e.message)}
});
$("clearWriteWebhook")?.addEventListener("click",()=>{
  clearWriteWebhookUrl(); setCloudStatus(); $("makeWriteWebhookUrl").value="";
  toast("URL de guardado eliminada de este dispositivo");
});
$("testCloud")?.addEventListener("click",()=>loadCloudWorkbook(true));

setCloudStatus();

if("serviceWorker" in navigator){
  navigator.serviceWorker.register("./service-worker.js")
    .catch(()=>{});
}

let deferredPrompt=null;
window.addEventListener("beforeinstallprompt",e=>{
  e.preventDefault(); deferredPrompt=e; $("installBtn").hidden=false;
});
$("installBtn").addEventListener("click",async()=>{
  if(!deferredPrompt)return;
  deferredPrompt.prompt(); await deferredPrompt.userChoice; deferredPrompt=null; $("installBtn").hidden=true;
});
updateCalc();
