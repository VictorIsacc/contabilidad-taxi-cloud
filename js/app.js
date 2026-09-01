import { calcularDia, euro } from "./calculos.js";
import { saveLocalDay, loadLocalDay, markLocalRest } from "./storage.js";
import {
  getWebhookUrl, saveWebhookUrl, clearWebhookUrl, fetchWorkbook,
  workbookSummary, findContabilidadDate
} from "./cloud-data.js";

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
  $("rSuma").textContent = source && source.suma_total!=="" ? euro(source.suma_total) : euro(r.suma_total);
  $("rMitad").textContent = source && source.mitad!=="" ? euro(source.mitad) : euro(r.mitad);
  $("rComisiones").textContent = source && source.total_comisiones!=="" ? euro(source.total_comisiones) : "Pendiente fórmula";
  $("rJefe").textContent = source && source.total_jefe!=="" ? euro(source.total_jefe) : "Pendiente fórmula";
}
function setCloudStatus(){
  const url=getWebhookUrl();
  if($("makeWebhookUrl")) $("makeWebhookUrl").value=url;
  if($("makeState")) $("makeState").textContent=url?"Configurada en este dispositivo":"Sin configurar";
  $("cloudDot")?.classList.toggle("online",!!url);
  if($("accountText")) $("accountText").textContent=url?"Make + OneDrive configurado":"Nube sin configurar";
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

qsa(".tab").forEach(btn=>btn.addEventListener("click",()=>{
  qsa(".tab").forEach(x=>x.classList.toggle("active",x===btn));
  qsa(".tabpage").forEach(p=>p.classList.toggle("active",p.id==="tab-"+btn.dataset.tab));
}));
qsa("[data-field]").forEach(i=>i.addEventListener("input",()=>updateCalc()));

$("workDate").value=isoToday();

$("saveDay").addEventListener("click",()=>{
  saveLocalDay($("workDate").value,{descanso:false,...collectConta()});
  toast("Guardado local de prueba. El guardado en OneDrive será el siguiente paso.");
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
