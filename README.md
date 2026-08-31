# Intranet SPFx VGM

Prueba minima de SharePoint Framework (SPFx) para entender el ciclo de desarrollo, despliegue y actualizacion de un Web Part en SharePoint Online.

## Stack

- SharePoint Framework 1.23.2
- Node.js 22
- TypeScript
- Heft
- Sin React para mantener la primera prueba lo mas simple posible

## Que muestra

El Web Part **VGM - Prueba SPFx** muestra una tarjeta de bienvenida, el nombre del sitio actual y un boton que cambia un mensaje en pantalla.

## Primera ejecucion en macOS

Comprueba primero que estas usando Node 22:

```bash
node -v
```

Luego, desde la carpeta del proyecto:

```bash
npm install
npx heft trust-dev-cert
npm start
```

`trust-dev-cert` solo es necesario la primera vez que preparas SPFx en ese equipo.

El proyecto esta configurado para abrir el SharePoint Workbench de VGM:

```text
https://vgmconsultants.sharepoint.com/_layouts/15/workbench.aspx
```

En el selector de Web Parts busca **VGM - Prueba SPFx** y agregalo a la pagina.

## Generar paquete para SharePoint

```bash
npm run build
```

El paquete queda en:

```text
sharepoint/solution/intranet-spfx-vgm.sppkg
```

Subelo al App Catalog de SharePoint y despliega la solucion.

## Probar una actualizacion

1. Modifica `src/webparts/vgmDemo/VgmDemoWebPart.ts`.
2. Cambia la version en `config/package-solution.json`, por ejemplo de `1.0.0.0` a `1.0.0.1`.
3. Ejecuta `npm run build`.
4. Reemplaza el `.sppkg` existente en el App Catalog.
5. Vuelve a cargar la pagina que ya contiene el Web Part.

No es necesario quitar y volver a agregar el Web Part para cambios normales de codigo.
