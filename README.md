# Contabilidad Taxi Cloud · GitHub Pages v0.2

## Novedad principal
La PWA puede llamar al webhook de Make, recibir el Excel de OneDrive y leerlo directamente
en el navegador.

La URL del webhook NO se incluye en el repositorio. Se introduce desde la pestaña Nube y
se guarda únicamente en `localStorage` de ese teléfono/PC.

## Antes de probar desde GitHub Pages: un único ajuste en Make
En el módulo final `Webhook response`, añade un tercer Custom header:

Key:
Access-Control-Allow-Origin

Value:
https://victorisacc.github.io

Los otros dos headers que ya tienes se mantienen:
- Content-Type
- Content-Disposition

Esto permite que la PWA de GitHub Pages lea el Excel mediante `fetch`.

## Cómo probar
1. Sube esta versión al mismo repositorio sustituyendo los archivos anteriores.
2. Abre la web de GitHub Pages.
3. Entra en la pestaña Nube.
4. Pega tu URL privada del webhook de Make.
5. Pulsa Guardar conexión.
6. Pulsa Cargar Excel desde OneDrive.
7. Deben aparecer las hojas del libro.
8. En Contabilidad selecciona una fecha existente y pulsa Cargar día.

## Seguridad
No publiques la URL del webhook en GitHub. La app la guarda solo en el navegador.
Esta versión todavía NO escribe cambios en OneDrive; solo lee y carga datos.
El guardado Cloud será el siguiente paso después de verificar la lectura.


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
