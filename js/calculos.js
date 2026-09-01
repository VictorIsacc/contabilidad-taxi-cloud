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

// Reglas vigentes de la hoja Contabilidad del libro real.
// JOIN UP aplica un 10 % de comisión e IMBRIC un 12 %.
export function calcularDia(v){
  const cierrePideTaxi = num(v.cierre_pidetaxi);
  const abonadosSinComision = num(v.abonados_sin_comision);
  const uber = num(v.uber);
  const uberCash = num(v.ubercash);
  const joinupBruto = num(v.joinup_bruto);
  const imbricBruto = num(v.imbric_bruto);
  const cobroTarjeta = num(v.cobro_tarjeta);
  const gasolinaLavado = num(v.gasolina_lavado);
  const seguro = num(v.seguro);

  const joinupNeto = joinupBruto * 0.90;
  const comisionJoinup = joinupBruto - joinupNeto;
  const imbricNeto = imbricBruto * 0.88;
  const comisionImbric = imbricBruto - imbricNeto;
  const totalAbonadosNeto =
    abonadosSinComision + uber + joinupNeto + imbricNeto - uberCash;
  const totalComisiones = comisionJoinup + comisionImbric;
  const totalPideTaxiSinComisiones = cierrePideTaxi - totalComisiones;
  const sumaTotal = totalPideTaxiSinComisiones + uber;
  const mitad = sumaTotal / 2;
  const totalJefeSeguro = mitad + seguro;
  const totalJefe =
    totalJefeSeguro - totalAbonadosNeto - cobroTarjeta - gasolinaLavado;
  const meQueda = mitad - seguro;

  return {
    suma_total: sumaTotal,
    joinup_neto: joinupNeto,
    comision_joinup: comisionJoinup,
    imbric_neto: imbricNeto,
    comision_imbric: comisionImbric,
    total_abonados_neto: totalAbonadosNeto,
    total_comisiones: totalComisiones,
    total_pidetaxi_sin_comisiones: totalPideTaxiSinComisiones,
    mitad,
    seguro,
    total_jefe_seguro: totalJefeSeguro,
    total_jefe: totalJefe,
    me_queda: meQueda
  };
}
