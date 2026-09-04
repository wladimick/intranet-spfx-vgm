# TIBOX Organigrama SPFx

Aplicación SPFx independiente para mostrar el organigrama corporativo en SharePoint con lupa circular interactiva.

## Objetivo

- Mostrar una imagen de alta resolución del organigrama.
- Convertir el puntero en una lupa circular que amplía el detalle sin hacer llamadas adicionales a SharePoint.
- Permitir ajustar zoom y tamaño de la lupa.
- Permitir pantalla completa.
- Permitir a usuarios con permiso `ManageWeb` (por ejemplo Control total) reemplazar la imagen publicada.

## Almacenamiento

La imagen activa se guarda en:

`SiteAssets/TIBOX-Organigrama/organigrama-current.png`

La carpeta se crea solamente cuando un administrador publica una imagen por primera vez. No hay provisionamiento automático durante la carga normal.

## Rendimiento y seguridad

La navegación normal no realiza barridos, polling ni consultas REST masivas. La imagen se sirve como un archivo estático de SharePoint y la lupa se resuelve completamente en el navegador.

La actualización usa SharePoint REST solo cuando un usuario autorizado selecciona una nueva imagen.

## Formatos de actualización

PNG, JPG o WebP. La aplicación convierte la imagen seleccionada a PNG antes de publicarla y limita el lado mayor a 9000 px para evitar archivos desproporcionados.

## Desarrollo

```bash
cd clientes/tibox/organigrama
npm install
npm start
```

Workbench:

`https://tibox1.sharepoint.com/_layouts/15/workbench.aspx`

## Build

```bash
npm run build
```

Paquete:

`sharepoint/solution/tibox-organigrama.sppkg`
