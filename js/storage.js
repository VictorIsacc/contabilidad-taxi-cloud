const KEY="contabilidadTaxiCloudDemoDays";

function all(){
  try{return JSON.parse(localStorage.getItem(KEY)||"{}")}catch{return {}}
}
function write(data){localStorage.setItem(KEY,JSON.stringify(data))}
export function saveLocalDay(date,data){
  const db=all(); db[date]=data; write(db);
}
export function loadLocalDay(date){
  return all()[date] || null;
}
export function markLocalRest(date){
  const db=all(); db[date]={descanso:true}; write(db);
}
