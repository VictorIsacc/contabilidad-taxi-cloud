# Contabilidad Taxi Cloud · GitHub Pages v0.9

## Estado comprobado
La PWA recibe desde Make el Excel de OneDrive y lo lee directamente en el navegador con SheetJS.
Se han reconocido las hojas reales `Contabilidad`, `Ingreso` y `Ahorro`.

La URL del webhook NO se incluye en el repositorio. Se introduce desde la pestaña Nube y
se guarda únicamente en `localStorage` de ese teléfono/PC.

## Lectura desde OneDrive
En el módulo final `Webhook response`, añade un tercer Custom header:

Key:
Access-Control-Allow-Origin

Value:
https://victorisacc.github.io

Los otros dos headers que ya tienes se mantienen:
- Content-Type
- Content-Disposition

Esto permite que la PWA de GitHub Pages lea el Excel mediante `fetch`.

Para que GitHub Pages pueda leer el Excel, el módulo final `Webhook response` del escenario
de lectura debe conservar este encabezado:

Key:
Access-Control-Allow-Origin

Value:
https://victorisacc.github.io

## Cómo comprobar la lectura
1. Sube esta versión al mismo repositorio sustituyendo los archivos anteriores.
2. Abre la web de GitHub Pages.
3. Entra en la pestaña Nube.
4. Pega tu URL privada del webhook de Make.
5. Pulsa Guardar conexión.
6. Pulsa Cargar Excel desde OneDrive.
7. Deben aparecer las hojas del libro.
8. En Contabilidad selecciona una fecha existente y pulsa Cargar día.

## Guardado en OneDrive
La pestaña Nube ya separa dos URLs privadas:

- Lectura: descarga el Excel desde OneDrive.
- Guardado: preparada para enviar a Make la fecha, número de fila validado y solo los ocho
  importes manuales de Contabilidad.

El escenario de guardado todavía debe configurarse en Make con Microsoft Graph. No se debe
usar "Actualizar una fila" porque sobrescribe una fila completa. El flujo seguro es:

1. Verificar la fecha recibida contra la columna A de `Contabilidad`.
2. Actualizar solo F, G, H, I, J, M, U y V de esa fila.
3. Dejar intactas todas las fórmulas, formatos y demás columnas.

## Seguridad
No publiques ninguna URL de webhook en GitHub. La app las guarda solo en el navegador de cada dispositivo.


## v0.3
- Corrige la respuesta de Make cuando el Buffer del Excel llega representado como texto hexadecimal.
- Reconstruye automáticamente los bytes XLSX antes de leer el libro.
- No requiere ningún cambio adicional en el escenario de Make.


## v0.4
- Adaptada al nuevo Webhook response de Make con `base64(9.Data)`.
- La PWA decodifica el Base64 a binario XLSX antes de abrir el libro.
- Verifica que el archivo recibido sea realmente un XLSX.


## v0.5
- Soporta la forma en la que Make representa el Buffer de OneDrive:
  Base64 -> texto hexadecimal -> bytes XLSX.
- Mantiene compatibilidad con Base64 binario directo y con hexadecimal sin Base64.
- No requiere nuevos cambios en Make.
