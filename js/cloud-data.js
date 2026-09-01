const WEBHOOK_KEY = "taxiMakeWebhookUrl";
let workbook = null;
let workbookBytes = null;

export function getWebhookUrl(){
  return localStorage.getItem(WEBHOOK_KEY) || "";
}

export function saveWebhookUrl(url){
  const clean = String(url || "").trim();
  if(!/^https:\/\/hook\.[^/]+\.make\.com\//i.test(clean)){
    throw new Error("La dirección no parece una URL válida de webhook de Make.");
  }
  localStorage.setItem(WEBHOOK_KEY, clean);
  return clean;
}

export function clearWebhookUrl(){
  localStorage.removeItem(WEBHOOK_KEY);
}

export async function fetchWorkbook(){
  const url = getWebhookUrl();
  if(!url) throw new Error("Primero guarda la URL privada del webhook de Make.");

  const response = await fetch(url, {
    method: "GET",
    cache: "no-store"
  });

  if(!response.ok){
    throw new Error(`Make respondió con error ${response.status}.`);
  }

  const text = (await response.text()).trim();
  if(!text) throw new Error("Make no devolvió contenido del Excel.");
  if(text === "Accepted"){
    throw new Error("Make ha puesto la solicitud en cola. Activa el escenario o ejecútalo una vez antes de cargar el Excel.");
  }

  let normalized = null;

  function isZipXlsx(buffer){
    const sig = new Uint8Array(buffer, 0, Math.min(4, buffer.byteLength));
    return sig.length >= 4 && sig[0] === 0x50 && sig[1] === 0x4B;
  }

  function hexTextToBuffer(value){
    // Acepta texto tipo:
    // 50 4b 03 04 ...
    // 504b0304...
    // y también posibles prefijos de Make.
    const pairs = String(value).match(/[0-9a-fA-F]{2}/g);
    if(!pairs || pairs.length < 4) return null;

    // Localizamos la firma ZIP "50 4B 03 04" por si existe un prefijo textual.
    let startIndex = -1;
    for(let i=0; i<=pairs.length-4; i++){
      if(
        pairs[i].toUpperCase()==="50" &&
        pairs[i+1].toUpperCase()==="4B" &&
        pairs[i+2].toUpperCase()==="03" &&
        pairs[i+3].toUpperCase()==="04"
      ){
        startIndex = i;
        break;
      }
    }
    if(startIndex < 0) return null;

    const usable = pairs.slice(startIndex);
    const bytes = new Uint8Array(usable.length);
    for(let i=0; i<usable.length; i++) bytes[i] = parseInt(usable[i], 16);
    return bytes.buffer;
  }

  // CASO 1: Make devuelve Base64.
  try{
    const compact = text.replace(/\s+/g, "");
    const binary = atob(compact);
    const bytes = new Uint8Array(binary.length);
    for(let i=0; i<binary.length; i++) bytes[i] = binary.charCodeAt(i);

    if(isZipXlsx(bytes.buffer)){
      // Base64 contenía directamente el XLSX.
      normalized = bytes.buffer;
    }else{
      // Base64 contenía el texto hexadecimal generado por Make.
      const decodedText = new TextDecoder("utf-8").decode(bytes);
      const fromHex = hexTextToBuffer(decodedText);
      if(fromHex && isZipXlsx(fromHex)) normalized = fromHex;
    }
  }catch(_e){
    // Seguimos con los formatos alternativos.
  }

  // CASO 2: por compatibilidad, Make puede devolver el hexadecimal sin Base64.
  if(!normalized){
    const fromHex = hexTextToBuffer(text);
    if(fromHex && isZipXlsx(fromHex)) normalized = fromHex;
  }

  if(!normalized){
    throw new Error("Make ha respondido, pero no he podido reconstruir el archivo XLSX.");
  }

  workbookBytes = normalized;
  workbook = XLSX.read(normalized, {
    type: "array",
    cellDates: true,
    cellFormula: true,
    cellNF: true
  });

  // No damos por bueno un libro vacío ni una Sheet1 ficticia.
  if(!workbook.SheetNames || !workbook.SheetNames.length){
    throw new Error("El XLSX recibido no contiene hojas reconocibles.");
  }
  const requiredSheets = ["Contabilidad", "Ingreso", "Ahorro"];
  const missingSheets = requiredSheets.filter(name=>!workbook.SheetNames.includes(name));
  if(missingSheets.length){
    throw new Error(`El XLSX recibido no es el libro esperado. Faltan: ${missingSheets.join(", ")}.`);
  }

  return workbook;
}

export function getWorkbook(){
  return workbook;
}

export function getWorkbookBytes(){
  return workbookBytes;
}

export function workbookSummary(){
  if(!workbook) return null;
  return workbook.SheetNames.map(name => {
    const ws = workbook.Sheets[name];
    const range = ws && ws["!ref"] ? XLSX.utils.decode_range(ws["!ref"]) : null;
    return {
      name,
      rows: range ? range.e.r + 1 : 0,
      cols: range ? range.e.c + 1 : 0
    };
  });
}

function excelDateToISO(value){
  if(value instanceof Date && !Number.isNaN(value.valueOf())){
    const y=value.getFullYear();
    const m=String(value.getMonth()+1).padStart(2,"0");
    const d=String(value.getDate()).padStart(2,"0");
    return `${y}-${m}-${d}`;
  }
  if(typeof value === "number"){
    const p=XLSX.SSF.parse_date_code(value);
    if(p) return `${p.y}-${String(p.m).padStart(2,"0")}-${String(p.d).padStart(2,"0")}`;
  }
  if(typeof value === "string"){
    const t=value.trim();
    let m=t.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
    if(m) return `${m[3]}-${m[2].padStart(2,"0")}-${m[1].padStart(2,"0")}`;
    m=t.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if(m) return `${m[1]}-${m[2].padStart(2,"0")}-${m[3].padStart(2,"0")}`;
  }
  return "";
}

function cell(ws, addr){
  return ws[addr]?.v ?? "";
}

function numberValue(value){
  if(value === null || value === undefined || value === "") return 0;
  if(typeof value === "number") return Number.isFinite(value) ? value : 0;
  if(typeof value === "string" && value.trim().toLowerCase() === "descanso") return 0;
  let text=String(value).trim().replace(/€/g, "").replace(/\s/g, "");
  if(text.includes(",") && text.includes(".")) text=text.replace(/\./g, "").replace(",", ".");
  else text=text.replace(",", ".");
  const parsed=Number(text);
  return Number.isFinite(parsed) ? parsed : 0;
}

function isRest(value){
  return typeof value === "string" && value.trim().toLowerCase() === "descanso";
}

function findDateRow(ws, isoDate, startRow=1){
  const range=ws?.["!ref"] ? XLSX.utils.decode_range(ws["!ref"]) : {e:{r:0}};
  for(let row=startRow; row<=range.e.r+1; row++){
    if(excelDateToISO(cell(ws, `A${row}`)) === isoDate) return row;
  }
  return 0;
}

export function findContabilidadDate(isoDate){
  if(!workbook) throw new Error("Primero carga el Excel desde OneDrive.");
  const ws = workbook.Sheets["Contabilidad"];
  if(!ws) throw new Error("El Excel no contiene la hoja Contabilidad.");

  const range = ws["!ref"] ? XLSX.utils.decode_range(ws["!ref"]) : {s:{r:0},e:{r:0}};
  for(let r=0;r<=range.e.r;r++){
    const row=r+1;
    const iso=excelDateToISO(cell(ws, `A${row}`));
    if(iso === isoDate){
      return {
        row,
        fecha: iso,
        dia: cell(ws, `B${row}`),
        descansoLetra: cell(ws, `C${row}`),
        letraDescanso: cell(ws, `D${row}`),
        suma_total: cell(ws, `E${row}`),
        cierre_pidetaxi: cell(ws, `F${row}`),
        abonados_sin_comision: cell(ws, `G${row}`),
        uber: cell(ws, `H${row}`),
        ubercash: cell(ws, `I${row}`),
        joinup_bruto: cell(ws, `J${row}`),
        joinup_neto: cell(ws, `K${row}`),
        comision_joinup: cell(ws, `L${row}`),
        imbric_bruto: cell(ws, `M${row}`),
        imbric_neto: cell(ws, `N${row}`),
        comision_imbric: cell(ws, `O${row}`),
        total_abonados_neto: cell(ws, `P${row}`),
        total_comisiones: cell(ws, `Q${row}`),
        total_pidetaxi_sin_comisiones: cell(ws, `R${row}`),
        mitad: cell(ws, `S${row}`),
        seguro: cell(ws, `T${row}`),
        cobro_tarjeta: cell(ws, `U${row}`),
        gasolina_lavado: cell(ws, `V${row}`),
        total_jefe: cell(ws, `W${row}`),
        total_jefe_seguro: cell(ws, `X${row}`),
        me_queda: cell(ws, `Y${row}`)
      };
    }
  }
  return null;
}

export function findIngresoDate(isoDate){
  if(!workbook) throw new Error("Primero carga el Excel desde OneDrive.");
  const ws=workbook.Sheets["Ingreso"];
  if(!ws) throw new Error("El Excel no contiene la hoja Ingreso.");
  const row=findDateRow(ws, isoDate, 2);
  if(!row) return null;

  const descanso=isRest(cell(ws, `H${row}`));
  const inputs={
    b100: descanso ? "" : cell(ws, `C${row}`),
    b50: descanso ? "" : cell(ws, `D${row}`),
    b20: descanso ? "" : cell(ws, `E${row}`),
    b10: descanso ? "" : cell(ws, `F${row}`),
    b5: descanso ? "" : cell(ws, `G${row}`)
  };
  const outputs={
    total: cell(ws, `H${row}`),
    gasto: cell(ws, `I${row}`),
    ingresos: cell(ws, `J${row}`),
    me_queda: cell(ws, `K${row}`),
    a_deber: cell(ws, `L${row}`)
  };

  const parts=isoDate.split("-").map(Number);
  const monthRows=[];
  const range=ws["!ref"] ? XLSX.utils.decode_range(ws["!ref"]) : {e:{r:0}};
  for(let r=2; r<=range.e.r+1; r++){
    const iso=excelDateToISO(cell(ws, `A${r}`));
    if(iso && Number(iso.slice(0,4))===parts[0] && Number(iso.slice(5,7))===parts[1]) monthRows.push(r);
  }
  let remainderRow=0;
  if(monthRows.length){
    for(let r=monthRows[0]-1; r>=2; r--){
      if(["C","D","E","F","G","H","J"].some(col=>{
        const value=cell(ws, `${col}${r}`);
        return value!=="" && value!==null && value!==undefined && !isRest(value);
      })){ remainderRow=r; break; }
    }
  }
  const keys=["mes_b100","mes_b50","mes_b20","mes_b10","mes_b5"];
  const cols=["C","D","E","F","G"];
  const month={}, remainder={}, combined={};
  keys.forEach((key,index)=>{
    month[key]=monthRows.reduce((sum,r)=>sum+numberValue(cell(ws, `${cols[index]}${r}`)),0);
    remainder[key]=remainderRow ? numberValue(cell(ws, `${cols[index]}${remainderRow}`)) : 0;
    combined[key]=month[key]+remainder[key];
  });
  const cash=data=>data.mes_b100*100+data.mes_b50*50+data.mes_b20*20+data.mes_b10*10+data.mes_b5*5;
  month.mes_total_efectivo=cash(month);
  remainder.mes_total_efectivo=cash(remainder);
  combined.mes_total_efectivo=cash(combined);

  return {row, fecha:isoDate, descanso, inputs, outputs, month:{
    label:`Acumulado del mes ${String(parts[1]).padStart(2,"0")}/${parts[0]}`,
    days:monthRows.length,
    remainder_row:remainderRow || null,
    month,remainder,combined
  }};
}

export function loadAhorroRow(row){
  if(!workbook) throw new Error("Primero carga el Excel desde OneDrive.");
  const ws=workbook.Sheets["Ahorro"];
  if(!ws) throw new Error("El Excel no contiene la hoja Ahorro.");
  row=Number(row);
  if(!Number.isInteger(row) || row<2 || row>16) throw new Error("La fila de Ahorro debe estar entre 2 y 16.");
  const inputs={
    date:excelDateToISO(cell(ws, `A${row}`)),
    b100:cell(ws, `B${row}`), b50:cell(ws, `C${row}`), b20:cell(ws, `D${row}`),
    b10:cell(ws, `E${row}`), b5:cell(ws, `F${row}`),
    gasto:cell(ws, `H${row}`), detalle:cell(ws, `I${row}`)
  };
  return {row,inputs,ingreso:cell(ws, `G${row}`),summary:{
    tot_100:cell(ws,"B17"),tot_50:cell(ws,"C17"),tot_20:cell(ws,"D17"),tot_10:cell(ws,"E17"),tot_5:cell(ws,"F17"),tot_ingreso:cell(ws,"G17"),
    rest_100:cell(ws,"B19"),rest_50:cell(ws,"C19"),rest_20:cell(ws,"D19"),rest_10:cell(ws,"E19"),rest_5:cell(ws,"F19"),rest_ingreso:cell(ws,"G19")
  }};
}

export function findAhorroDate(isoDate){
  if(!workbook) throw new Error("Primero carga el Excel desde OneDrive.");
  const row=findDateRow(workbook.Sheets["Ahorro"], isoDate, 2);
  return row>=2 && row<=16 ? loadAhorroRow(row) : null;
}

export function nextFreeAhorroRow(){
  if(!workbook) throw new Error("Primero carga el Excel desde OneDrive.");
  const ws=workbook.Sheets["Ahorro"];
  for(let row=2; row<=16; row++) if(!excelDateToISO(cell(ws, `A${row}`))) return row;
  return null;
}

const PERIOD_FIELDS=[
  ["mes_suma_total","E"],["mes_cierre_pidetaxi","F"],["mes_abonados_sin_comision","G"],["mes_uber","H"],
  ["mes_ubercash","I"],["mes_joinup_bruto","J"],["mes_imbric_bruto","M"],["mes_total_abonados_neto","P"],
  ["mes_total_comisiones","Q"],["mes_mitad","S"],["mes_seguro","T"],["mes_cobro_tarjeta","U"],
  ["mes_gasolina_lavado","V"],["mes_total_jefe","W"],["mes_total_jefe_seguro","X"],["mes_me_queda","Y"]
];

export function contabilidadPeriod({month="",start="",end=""}={}){
  if(!workbook) throw new Error("Primero carga el Excel desde OneDrive.");
  const ws=workbook.Sheets["Contabilidad"];
  const range=ws["!ref"] ? XLSX.utils.decode_range(ws["!ref"]) : {e:{r:0}};
  let monthYear=0, monthNumber=0;
  if(month){
    const match=String(month).trim().match(/^(\d{1,2})[\/-](\d{4})$/);
    if(!match) throw new Error("Mes no válido. Usa mm/aaaa.");
    monthNumber=Number(match[1]); monthYear=Number(match[2]);
  }
  if(!month && (!start || !end)) throw new Error("Indica un mes o un rango de fechas.");
  if(start && end && start>end) throw new Error("La fecha inicial no puede ser posterior a la final.");
  const rows=[];
  for(let row=5; row<=range.e.r+1; row++){
    const iso=excelDateToISO(cell(ws, `A${row}`));
    if(!iso) continue;
    const include=month
      ? Number(iso.slice(0,4))===monthYear && Number(iso.slice(5,7))===monthNumber
      : iso>=start && iso<=end;
    if(include) rows.push(row);
  }
  const totals={};
  PERIOD_FIELDS.forEach(([key,col])=>totals[key]=rows.reduce((sum,row)=>sum+numberValue(cell(ws, `${col}${row}`)),0));
  const restdays=rows.filter(row=>isRest(cell(ws, `E${row}`))).length;
  const workdays=Math.max(0,rows.length-restdays);
  const averages={};
  PERIOD_FIELDS.forEach(([key])=>averages[key]=workdays ? totals[key]/workdays : null);
  return {
    label:month ? `Acumulado del mes ${String(monthNumber).padStart(2,"0")}/${monthYear}` : `Acumulado del rango ${start} a ${end}`,
    stats:{days:rows.length,workdays,restdays},totals,averages,rows
  };
}
