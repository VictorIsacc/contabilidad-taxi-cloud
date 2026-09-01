import { getAccessToken } from "./auth.js";

const GRAPH = "https://graph.microsoft.com/v1.0";

async function graph(path, options={}){
  const token = await getAccessToken();
  const headers = new Headers(options.headers || {});
  headers.set("Authorization", `Bearer ${token}`);
  const response = await fetch(GRAPH + path, {...options, headers});
  if(!response.ok){
    const text = await response.text();
    throw new Error(`Microsoft Graph ${response.status}: ${text}`);
  }
  return response;
}

export async function getDriveInfo(){
  const r = await graph("/me/drive");
  return r.json();
}

export async function downloadWorkbook(itemId){
  const r = await graph(`/me/drive/items/${encodeURIComponent(itemId)}/content`);
  return r.arrayBuffer();
}

export async function uploadWorkbook(itemId, arrayBuffer){
  const r = await graph(`/me/drive/items/${encodeURIComponent(itemId)}/content`, {
    method:"PUT",
    headers:{"Content-Type":"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"},
    body:arrayBuffer
  });
  return r.json();
}

// La selección del archivo se conectará en el siguiente paso.
// Guardamos el ID estable del archivo, no una ruta fija.
export function rememberWorkbook(item){
  localStorage.setItem("taxiWorkbookItemId", item.id);
  localStorage.setItem("taxiWorkbookName", item.name);
}
