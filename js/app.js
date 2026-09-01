import { calcularDia, euro } from "./calculos.js";
import { saveLocalDay, loadLocalDay, markLocalRest } from "./storage.js";
import { microsoftConfigured, signInMicrosoft } from "./auth.js";
import { CLOUD_CONFIG } from "./config.js";

const $=id=>document.getElementById(id);
const qsa=s=>[...document.querySelectorAll(s)];

function toast(msg){
  const el=$("toast"); el.textContent=msg; el.classList.add("show");
  clearTimeout(window.__toast); window.__toast=setTimeout(()=>el.classList.remove("show"),2600);
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
  updateCalc();
}
function updateCalc(){
  const r=calcularDia(collectConta());
  $("rSuma").textContent=euro(r.suma_total);
  $("rMitad").textContent=euro(r.mitad);
  $("rComisiones").textContent=r.total_comisiones===null?"Pendiente fórmula":euro(r.total_comisiones);
  $("rJefe").textContent=r.total_jefe===null?"Pendiente fórmula":euro(r.total_jefe);
}
qsa(".tab").forEach(btn=>btn.addEventListener("click",()=>{
  qsa(".tab").forEach(x=>x.classList.toggle("active",x===btn));
  qsa(".tabpage").forEach(p=>p.classList.toggle("active",p.id==="tab-"+btn.dataset.tab));
}));
qsa("[data-field]").forEach(i=>i.addEventListener("input",updateCalc));

$("workDate").value=isoToday();
$("saveDay").addEventListener("click",()=>{
  saveLocalDay($("workDate").value,{descanso:false,...collectConta()});
  toast("Día guardado en modo local de prueba");
});
$("loadDay").addEventListener("click",()=>{
  const d=loadLocalDay($("workDate").value);
  if(!d){toast("No hay datos guardados para ese día");return}
  if(d.descanso){fillConta({});toast("Ese día está marcado como descanso");return}
  fillConta(d);toast("Día cargado");
});
$("markRest").addEventListener("click",()=>{
  markLocalRest($("workDate").value); fillConta({}); toast("Descanso guardado en modo local");
});

$("loginMicrosoft").addEventListener("click",async()=>{
  try{await signInMicrosoft()}catch(e){toast(e.message)}
});

$("clientState").textContent=microsoftConfigured()?"Configurado":"Pendiente";
$("workbookName").textContent=CLOUD_CONFIG.workbook.name || "Ninguno";

if("serviceWorker" in navigator){
  navigator.serviceWorker.register("./service-worker.js")
    .then(()=>{$("pwaState").textContent="Activa"})
    .catch(()=>{$("pwaState").textContent="Error"});
}else $("pwaState").textContent="No compatible";

let deferredPrompt=null;
window.addEventListener("beforeinstallprompt",e=>{
  e.preventDefault(); deferredPrompt=e; $("installBtn").hidden=false;
});
$("installBtn").addEventListener("click",async()=>{
  if(!deferredPrompt)return;
  deferredPrompt.prompt(); await deferredPrompt.userChoice; deferredPrompt=null; $("installBtn").hidden=true;
});
updateCalc();
