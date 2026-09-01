export function num(value){
  if(value === null || value === undefined || value === "") return 0;
  if(typeof value === "number") return Number.isFinite(value) ? value : 0;
  let t = String(value).trim().replace("€","").replace(/\s/g,"");
  if(t.includes(",") && t.includes(".")) t = t.replace(/\./g,"").replace(",",".");
  else t = t.replace(",",".");
  const n = Number(t);
  return Number.isFinite(n) ? n : 0;
}
export function euro(value){
  return new Intl.NumberFormat("es-ES",{style:"currency",currency:"EUR"}).format(num(value));
}

// Primera capa de cálculo independiente de Excel.
// Las comisiones y reglas exactas se trasladarán cuando carguemos el libro real.
export function calcularDia(v){
  const suma =
    num(v.cierre_pidetaxi) +
    num(v.abonados_sin_comision) +
    num(v.uber) +
    num(v.ubercash) +
    num(v.joinup_bruto) +
    num(v.imbric_bruto);

  return {
    suma_total:suma,
    total_comisiones:null,
    mitad:suma/2,
    total_jefe:null
  };
}
