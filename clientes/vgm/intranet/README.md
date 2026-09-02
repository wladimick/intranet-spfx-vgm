# VGM Intranet SPFx

Migración de la portada histórica de VGM a SharePoint Framework (SPFx 1.23.2).

## Objetivo

Mantener la funcionalidad actual de `https://vgmconsultants.sharepoint.com/sites/intranet` pero eliminar dependencias antiguas de jQuery, Semantic UI, Slick, Owl Carousel y scripts globales.

## Fuentes existentes reutilizadas

- Lista `Menú Principal`
- Lista `Contactos` (cumpleaños)
- Lista `Calendario`
- Lista `Noticias`
- Lista `Links de Interés`
- Biblioteca `Galería`
- Lista `Todos los Clientes - Repositorio Principal`
- `mindicador.cl` para indicadores económicos
- `tiboxrssreader.azurewebsites.net/api/rss/df` para Diario Financiero

La primera versión no crea ni modifica estas listas: las consume desde el sitio histórico para permitir una migración gradual.

## Clientes y permisos

Se mantiene la lógica de carpetas Legal, Tax, Outsourcing y Auditoría según grupos de Microsoft 365. La solución solicita `GroupMember.Read.All`. Si Graph no está aprobado o falla, se utiliza el comportamiento compatible con la intranet anterior: no bloquear la visualización por error de resolución de grupos.

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

## Estado MVP

Conectado: usuario/foto, menú principal, accesos históricos, Diario Financiero, noticias internas, indicadores, calendario/eventos, cumpleaños + Teams, galería, links y listado de clientes con filtros y carpetas.

Siguiente etapa: QA contra datos reales, revisar campos internos de listas, mejorar carruseles/galería y decidir qué contenidos deben seguir consumiendo listas legacy versus nuevas listas modernas.
