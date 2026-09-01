# Contabilidad Taxi Cloud · Base GitHub Pages v0.1

Esta carpeta está preparada para publicarse directamente con GitHub Pages.

## Qué funciona ya

- Interfaz web responsive para PC y teléfono.
- PWA instalable.
- Manifest e iconos.
- Service Worker.
- Pestañas Contabilidad / Ingreso / Ahorro / Análisis / Nube.
- Modo local de prueba para guardar y cargar días.
- Primera capa de cálculo JavaScript.
- Estructura preparada para Microsoft Graph y OneDrive.
- El libro de OneDrive se identificará por su `itemId`, no por una ruta fija.

## Qué NO está conectado todavía

El acceso Microsoft está deliberadamente pendiente. Antes hay que registrar la aplicación
en Microsoft y obtener un Client ID. No pongas contraseñas ni secretos en estos archivos.

## GitHub Pages

1. Crea un repositorio, por ejemplo `contabilidad-taxi-cloud`.
2. Sube TODOS los archivos y carpetas de este paquete a la raíz del repositorio.
3. GitHub -> Settings -> Pages.
4. En Build and deployment:
   - Source: Deploy from a branch.
   - Branch: main.
   - Folder: /(root).
5. Guarda.
6. GitHub mostrará una URL parecida a:
   `https://TU-USUARIO.github.io/contabilidad-taxi-cloud/`

No hace falta GitHub Actions para esta primera versión.

## Siguiente paso

Con la URL definitiva de GitHub Pages:
1. Registrar Contabilidad Taxi Cloud en Microsoft.
2. Añadir esa URL como Redirect URI de tipo SPA.
3. Obtener el Client ID.
4. Añadirlo en `js/config.js`.
5. Activar MSAL.
6. Conectar OneDrive.
7. Seleccionar el Excel.
8. Trasladar las fórmulas restantes y las hojas Ingreso/Ahorro/Análisis.
