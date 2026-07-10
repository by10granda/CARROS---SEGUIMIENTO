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

## Cambio de aceite

El modulo `Cambio de aceite` concentra tambien el control de kilometraje y horometro. Usa estas columnas:

`ITEM, FECHA, LUGAR, MAESTRO, TALLER, CHOFER, PLACA, CANTIDAD DE PAGO, DESCRIPCION, KILOMETRAJE INICIAL, KILOMETRAJE FINAL, SUMA KILOMETRAJE, TRANSPORTE, HOROMETRO INICIAL, HOROMETRO FINAL, HORAS, CAMBIO DE ACEITE`

Reglas de alerta:

- Transporte `GRANDE`: alerta al acumular `10000 km` sin cambio de aceite.
- Transporte `PEQUENO`: alerta al acumular `5000 km` sin cambio de aceite.
- Cualquier transporte: alerta al acumular `250 horas` de horometro sin cambio de aceite.

Despues de modificar `apps-script/Code.gs`, publica una nueva version del Web App en Apps Script para que Google Sheets reciba tambien los registros de cambio de aceite actualizados.

## WhatsApp automatico

El Apps Script ya incluye integracion con Meta WhatsApp Cloud API. Para activarla, en Apps Script abre `Project Settings > Script properties` y agrega:

```text
WHATSAPP_PHONE_NUMBER_ID=tu_phone_number_id
WHATSAPP_ACCESS_TOKEN=tu_access_token
WHATSAPP_TEMPLATE_NAME=alerta_horometro
WHATSAPP_LANGUAGE_CODE=es_EC
```

La plantilla aprobada en Meta debe tener 3 variables. La segunda variable recibe kilometraje u horas segun la regla disparada:

```text
ALERTA PUNTO PAS: La placa {{1}} tiene {{2}} acumulados sin cambio de aceite. Ultimo registro: {{3}}.
```

Cuando una placa alcance el limite de kilometraje u horas sin cambio de aceite, Apps Script enviara automaticamente el mensaje a:

- `593939069555`

El script crea una hoja `ALERTAS ACEITE` para evitar repetir la misma alerta dentro del mismo bloque de kilometraje u horas.

11. Reinicia la app con `npm run dev`.

Si no se configura `VITE_GOOGLE_SCRIPT_URL`, el formulario abre pero no puede escribir en Google Sheets.

Script usado:

```javascript
const SHEET_ID = '1Evr1lpNSwLYXWgcRf5D5NSPUD5nR-YS_EczVK4PVAHI';

// Ver archivo apps-script/Code.gs
```
