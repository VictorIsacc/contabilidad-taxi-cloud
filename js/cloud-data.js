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

  let buf = await response.arrayBuffer();
  if(!buf.byteLength) throw new Error("Make no devolvió contenido del Excel.");

  // Make puede devolver el Buffer binario como texto hexadecimal
  // (por ejemplo: "50 4b 03 04 ..."). Si ocurre, lo reconstruimos
  // automáticamente antes de entregarlo a SheetJS.
  const raw = new Uint8Array(buf);
  let normalized = buf;

  try{
    const text = new TextDecoder("utf-8").decode(raw).trim();

    // Acepta "50 4B 03 04", "504B0304", comas o saltos de línea.
    const compact = text
      .replace(/0x/gi, "")
      .replace(/[\s,;:\-]+/g, "");

    const looksHex =
      compact.length >= 8 &&
      compact.length % 2 === 0 &&
      /^[0-9a-f]+$/i.test(compact) &&
      compact.slice(0, 8).toUpperCase() === "504B0304";

    if(looksHex){
      const bytes = new Uint8Array(compact.length / 2);
      for(let i=0;i<bytes.length;i++){
        bytes[i] = parseInt(compact.slice(i*2, i*2+2), 16);
      }
      normalized = bytes.buffer;
    }
  }catch(_e){
    // Si no es texto hexadecimal, se usa el binario original.
  }

  const sig = new Uint8Array(normalized, 0, Math.min(4, normalized.byteLength));
  if(sig.length < 4 || sig[0] !== 0x50 || sig[1] !== 0x4B){
    throw new Error("Make ha devuelto datos, pero todavía no tienen formato XLSX válido.");
  }

  workbookBytes = normalized;
  workbook = XLSX.read(normalized, {
    type: "array",
    cellDates: true,
    cellFormula: true,
    cellNF: true
  });
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
