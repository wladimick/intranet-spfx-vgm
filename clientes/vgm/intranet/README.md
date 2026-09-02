# VGM Intranet SPFx

Migración y evolución de la portada histórica de VGM a SharePoint Framework (SPFx 1.23.2).

## Objetivo

Mantener las fuentes y accesos existentes de `https://vgmconsultants.sharepoint.com/sites/intranet`, eliminar dependencias antiguas de jQuery/Semantic UI/Slick/Owl Carousel y transformar la portada en una intranet administrable y personalizada por usuario.

## Fuentes existentes reutilizadas

- Lista `Menú Principal`
- Lista `Contactos` (fallback de cumpleaños)
- Lista `Calendario`
- Lista `Noticias`
- Lista `Links de Interés`
- Biblioteca `Galería`
- Lista `Todos los Clientes - Repositorio Principal`
- `mindicador.cl` para indicadores económicos
- `tiboxrssreader.azurewebsites.net/api/rss/df` para Diario Financiero
- Microsoft Graph para cumpleaños Microsoft 365 (`User.Read.All`)

## Modelo de permisos de clientes

Los usuarios pueden tener únicamente acceso limitado al sitio y permiso real sobre una o más carpetas dentro de `Documentos compartidos`:

- Legal
- Tax
- Outsourcing
- Auditoría

La aplicación valida directamente esas carpetas con la sesión del usuario. El resultado se precarga en segundo plano, se mantiene en memoria y se cachea durante la sesión para evitar repetir cientos de comprobaciones al abrir `Mis clientes`.

El popup solo muestra clientes y áreas a los que el usuario tiene acceso efectivo.

## Personalización por usuario

- Clientes favoritos
- Accesos recientes
- Buscador de clientes/documentos
- Documentos recientes
- Clientes con mayor actividad
- Actividad por área

Favoritos y recientes se guardan localmente por usuario/navegador y nunca modifican los permisos reales de SharePoint.

## Administración

Los administradores de la colección de sitios, o miembros del grupo SharePoint `VGM Intranet - Administradores`, ven el panel `Administración` y `Modo edición`.

La configuración global se persiste en la lista `VGM Intranet - Configuracion`, creada al guardar por primera vez si aún no existe.

Configuración disponible:

- colores y apariencia
- ancho máximo, radio de tarjetas y tamaño base
- mostrar/ocultar módulos
- orden de las cards de actividad
- CSS personalizado con scope obligatorio `.vgmApp`
- diagnóstico de fuentes y servicios
- estructura para accesos rápidos administrables

## CSS personalizado

El CSS se guarda en SharePoint y se aplica a todos los usuarios sin recompilar el `.sppkg`.

Por seguridad, cada selector debe comenzar con `.vgmApp`. Se bloquean selectores globales como `html`, `body`, `:root`, `@import`, `javascript:` y cierres de `<style>`.

Ejemplo:

```css
.vgmApp .vgmCard {
  border-radius: 10px;
}

.vgmApp .vgmHeader {
  box-shadow: 0 8px 20px rgba(0,0,0,.12);
}
```

## Desarrollo

```bash
cd clientes/vgm/intranet
npm install
npm start
```

Workbench:

`https://vgmconsultants.sharepoint.com/sites/intranet/_layouts/15/workbench.aspx`

## Build

```bash
npm run build
```

Paquete:

`sharepoint/solution/vgm-intranet.sppkg`
