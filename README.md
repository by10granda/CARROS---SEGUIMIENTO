# Distribuidor Punto PAS - Gestión Vehicular

Aplicación web React para consultar, filtrar, registrar y exportar mantenimientos vehiculares conectados a Google Sheets.

## Ejecutar

```bash
npm install
npm run dev
```

## Acceso

- Usuario: `Maria21`
- Contraseña: `PuntoPas2026*`

## Google Sheets

La lectura usa el endpoint público `gviz` del archivo configurado en `VITE_GOOGLE_SHEET_ID`. Para agregar registros desde el navegador se requiere publicar un Google Apps Script como Web App y colocar su URL en `VITE_GOOGLE_SCRIPT_URL`.

## Activar el botón Guardar en Google Sheets

1. Abre https://script.google.com/.
2. Crea un proyecto nuevo.
3. Copia el contenido de `apps-script/Code.gs` y pegalo en el editor de Apps Script.
4. Presiona `Deploy > New deployment`.
5. Selecciona tipo `Web app`.
6. En `Execute as`, selecciona tu usuario.
7. En `Who has access`, selecciona `Anyone with the link`.
8. Presiona `Deploy` y autoriza permisos.
9. Copia la URL que termina en `/exec`.
10. Crea un archivo `.env.local` en la raiz del proyecto con este contenido:

```bash
VITE_GOOGLE_SHEET_ID=1Evr1lpNSwLYXWgcRf5D5NSPUD5nR-YS_EczVK4PVAHI
VITE_GOOGLE_SCRIPT_URL=https://script.google.com/macros/s/TU_URL_AQUI/exec
```

La app usa estas pestañas del mismo archivo:

- Mantenimiento: `gid=0`
- Lavado: `gid=819388144`
- Engrasada: `gid=2024356449`
- Cambio de aceite: `gid=464443967`

11. Reinicia la app con `npm run dev`.

Si no se configura `VITE_GOOGLE_SCRIPT_URL`, el formulario abre pero no puede escribir en Google Sheets.

Script usado:

```javascript
const SHEET_ID = '1Evr1lpNSwLYXWgcRf5D5NSPUD5nR-YS_EczVK4PVAHI';

// Ver archivo apps-script/Code.gs
```
