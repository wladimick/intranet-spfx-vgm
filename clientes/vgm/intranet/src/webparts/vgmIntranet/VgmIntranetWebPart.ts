import { Version } from '@microsoft/sp-core-library';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import VgmDataService, {
  BirthdayItem, EventItem, FinancialNewsItem, GalleryItem, IndicatorItem, LinkItem, MenuItem, NewsItem
} from './VgmDataService';
import VgmSearchService, {
  VgmArea, VgmAreaActivity, VgmClientAccess, VgmClientActivity, VgmRecentDocument
} from './VgmSearchService';
import VgmFolderAccessService from './VgmFolderAccessService';
import VgmDirectActivityService from './VgmDirectActivityService';
import VgmGlobalSearchService from './VgmGlobalSearchService';
import VgmAdminConfigService, {
  DEFAULT_VGM_CONFIG, VgmActivityModuleKey, VgmConfigModuleKey, VgmPortalConfig
} from './VgmAdminConfigService';
import VgmUserPreferenceService, { VgmRecentAccess } from './VgmUserPreferenceService';
import { VGM_STYLES } from './VgmIntranetStyles';

export interface IVgmIntranetWebPartProps {}

type DashboardData = {
  menu: MenuItem[];
  birthdays: BirthdayItem[];
  events: EventItem[];
  news: NewsItem[];
  links: LinkItem[];
  gallery?: GalleryItem;
  indicators: IndicatorItem[];
  financial: FinancialNewsItem[];
};

type ClientViewMode = 'ALL' | 'FAVORITES' | 'RECENT';

type FolderScanProgress = { completed: number; total: number; found: number };

export default class VgmIntranetWebPart extends BaseClientSideWebPart<IVgmIntranetWebPartProps> {
  private service!: VgmDataService;
  private searchService!: VgmSearchService;
  private folderAccessService!: VgmFolderAccessService;
  private directActivityService!: VgmDirectActivityService;
  private globalSearchService!: VgmGlobalSearchService;
  private adminConfigService!: VgmAdminConfigService;
  private preferenceService!: VgmUserPreferenceService;

  private accessibleClients: VgmClientAccess[] = [];
  private recentDocuments: VgmRecentDocument[] = [];
  private activityClients: VgmClientActivity[] = [];
  private areaActivity: VgmAreaActivity = { LEGAL: 0, TAX: 0, OUTSOURCING: 0, AUDITORIA: 0 };
  private clientSearch: string = '';
  private clientAreaFilter: 'ALL' | VgmArea = 'ALL';
  private clientViewMode: ClientViewMode = 'ALL';
  private folderScanStarted: boolean = false;
  private folderScanCompleted: boolean = false;
  private folderScanProgress: FolderScanProgress = { completed: 0, total: 0, found: 0 };
  private isAdminUser: boolean = false;
  private editMode: boolean = false;
  private graphBirthdaysAvailable: boolean = false;
  private config: VgmPortalConfig = JSON.parse(JSON.stringify(DEFAULT_VGM_CONFIG)) as VgmPortalConfig;
  private data: DashboardData = { menu: [], birthdays: [], events: [], news: [], links: [], indicators: [], financial: [] };

  public render(): void {
    void this.renderAsync();
  }

  private async renderAsync(): Promise<void> {
    this.domElement.innerHTML = `<style>${VGM_STYLES}</style><div class="vgmApp"><div class="vgmLoading">Cargando intranet VGM…</div></div>`;
    const email: string = this.context.pageContext.user.email || '';
    this.service = new VgmDataService(this.context.spHttpClient);
    this.searchService = new VgmSearchService(this.context.spHttpClient);
    this.folderAccessService = new VgmFolderAccessService(this.context.spHttpClient,this.service);
    this.directActivityService = new VgmDirectActivityService(this.context.spHttpClient);
    this.globalSearchService = new VgmGlobalSearchService(this.context.spHttpClient);
    this.adminConfigService = new VgmAdminConfigService(this.context.spHttpClient,email);
    this.preferenceService = new VgmUserPreferenceService(email);

    const [menu, sharePointBirthdays, events, news, links, gallery, indicators, financial, searchClients, recentDocuments, activity, graphBirthdays, portalConfig, isAdmin] = await Promise.all([
      this.service.getMenu(),
      this.service.getBirthdays(),
      this.service.getEvents(),
      this.service.getInternalNews(),
      this.service.getLinks(),
      this.service.getGallery(),
      this.service.getIndicators(),
      this.service.getFinancialNews(),
      this.searchService.getAccessibleClients().catch((): VgmClientAccess[] => []),
      this.searchService.getRecentDocuments(8).catch((): VgmRecentDocument[] => []),
      this.searchService.getActivity(7,1200).catch(() => ({ clients: [] as VgmClientActivity[], areas: { LEGAL:0,TAX:0,OUTSOURCING:0,AUDITORIA:0 } as VgmAreaActivity, documents: [] as VgmRecentDocument[] })),
      this.loadMicrosoft365Birthdays(),
      this.adminConfigService.load(),
      this.adminConfigService.isAdmin()
    ]);

    const cachedClients: VgmClientAccess[] = this.loadAccessCache();
    this.accessibleClients = cachedClients.length ? cachedClients : searchClients;
    this.recentDocuments = recentDocuments;
    this.activityClients = activity.clients;
    this.areaActivity = activity.areas;
    this.config = portalConfig;
    this.isAdminUser = isAdmin;
    this.data = {
      menu,
      birthdays: this.mergeBirthdays(graphBirthdays,sharePointBirthdays),
      events,
      news,
      links,
      gallery,
      indicators,
      financial
    };

    this.paint();
    this.bindEvents();

    if (this.accessibleClients.length) {
      this.folderScanCompleted = true;
      void this.refreshDirectActivity();
    } else {
      void this.preloadFolderAccess();
    }
  }

  private paint(): void {
    const displayName: string = this.context.pageContext.user.displayName || 'Usuario VGM';
    const email: string = this.context.pageContext.user.email || '';
    const photo: string = `https://vgmconsultants.sharepoint.com/_layouts/15/userphoto.aspx?size=L&accountname=${encodeURIComponent(email)}`;
    const today: Date = new Date();
    const nextBirthdays: BirthdayItem[] = this.upcomingBirthdays(this.data.birthdays).slice(0,5);
    const featured: NewsItem | undefined = this.data.news[0];
    const secondary: NewsItem[] = this.data.news.slice(1,5);
    const customCss: string = this.safeStoredCss(this.config.customCss);

    this.domElement.innerHTML = `
      <style>${VGM_STYLES}\n${this.themeCss()}\n${customCss}</style>
      <div class="vgmApp ${this.editMode ? 'vgmEditMode' : ''}">
        <div class="vgmWrap">
          <header class="vgmHeader" data-admin-module data-module-label="Encabezado">
            <div class="vgmHeaderTop">
              <img class="vgmLogo" src="${VgmDataService.sourceWebUrl}/SiteAssets/portada/imagenes/LOGO-VGM-BLANCO.webp" alt="VGM Consultores">
              <div class="vgmSpacer"></div>
              ${this.isAdminUser ? `<div class="vgmAdminActions"><button class="vgmEditBtn ${this.editMode ? 'active' : ''}" data-action="edit-mode">✎ Modo edición</button><button class="vgmAdminBtn" data-action="admin">⚙ Administración</button></div>` : ''}
              <div class="vgmUser"><span><small>Bienvenido</small><strong>${this.esc(displayName)}</strong></span><img src="${photo}" alt="${this.esc(displayName)}"></div>
            </div>
            <nav class="vgmMenu">${this.renderMenu(this.data.menu)}<button class="vgmClientsBtn" data-action="clients">☰ Mis clientes${this.accessibleClients.length ? ` (${this.accessibleClients.length})` : this.folderScanStarted ? ' (…)': ''}</button></nav>
            ${this.enabled('personal') ? `<div class="vgmPortalSearch"><input data-global-search placeholder="Buscar documento o cliente dentro de mis accesos…"><button data-action="global-search">Buscar</button></div>` : ''}
          </header>

          <div class="vgmPortalGrid">
            <section class="vgmLeftArea">
              <div class="vgmLeftTop">
                <div class="vgmTools">
                  <section class="vgmCard"><div class="vgmCardBody"><div class="vgmDateBox">▣ ${this.capitalize(today.toLocaleDateString('es-CL',{weekday:'long',day:'numeric',month:'long'}))}</div></div></section>
                  ${this.renderLegacyQuickLinks()}
                  <a class="vgmTicket" href="https://soporte.tibox.cl/Login/LoginCliente" target="_blank" rel="noopener">◉ Tickets Tibox</a>
                </div>
                <div class="vgmCalendarColumn">
                  ${this.enabled('calendar') ? `<section class="vgmCard" data-admin-module data-module-label="Calendario"><div class="vgmCardHead"><h2>${this.capitalize(today.toLocaleDateString('es-CL',{month:'long'}))} ${today.getFullYear()}</h2></div><div class="vgmCardBody">${this.renderCalendar(today,this.data.events)}${this.renderEvents(this.data.events.slice(0,3))}</div></section>` : ''}
                  ${this.enabled('birthdays') ? `<section class="vgmCard" data-admin-module data-module-label="Cumpleaños"><div class="vgmCardHead"><h2>Próximos Cumpleaños</h2></div><div class="vgmCardBody">${this.renderBirthdays(nextBirthdays)}</div></section>` : ''}
                </div>
              </div>
              ${this.enabled('financial') ? `<section class="vgmCard vgmFinancialCard" data-admin-module data-module-label="Diario Financiero"><div class="vgmCardHead"><h2>Diario Financiero</h2></div><div class="vgmCardBody">${this.renderFinancial(this.data.financial)}</div></section>` : ''}
            </section>

            <section class="vgmRightArea">
              ${this.enabled('news') ? `<section class="vgmCard vgmNewsCard" data-admin-module data-module-label="Noticias"><div class="vgmCardHead"><h2>Noticias</h2></div>${this.renderNews(featured,secondary)}</section>` : ''}
              ${this.enabled('indicators') ? `<section class="vgmCard" data-admin-module data-module-label="Indicadores"><div class="vgmCardHead"><h2>Indicadores Económicos</h2></div><div class="vgmCardBody"><div class="vgmIndicators">${this.renderIndicators(this.data.indicators)}</div></div></section>` : ''}
              ${(this.enabled('gallery') || this.enabled('links')) ? `<div class="vgmBottom">
                ${this.enabled('gallery') ? `<section class="vgmCard" data-admin-module data-module-label="Galerías"><div class="vgmCardHead"><h2>Galerías</h2></div><div class="vgmCardBody">${this.renderGallery(this.data.gallery)}</div></section>` : ''}
                ${this.enabled('links') ? `<section class="vgmCard" data-admin-module data-module-label="Links"><div class="vgmCardHead"><h2>Links de Interés</h2></div><div class="vgmCardBody vgmLinks">${this.renderLinks(this.data.links)}</div></section>` : ''}
              </div>` : ''}
            </section>
          </div>

          ${this.enabled('personal') ? `<section class="vgmPersonalSection" data-admin-module data-module-label="Mi espacio"><div class="vgmWorkTitle"><span>MI ESPACIO</span><small>Favoritos y accesos recientes guardados para tu usuario</small></div><div class="vgmPersonalGrid" data-personal-grid>${this.renderPersonalSpace()}</div></section>` : ''}

          ${this.hasAnyActivityModule() ? `<section class="vgmWorkSection" data-admin-module data-module-label="Mi actividad"><div class="vgmWorkTitle"><span>MI ACTIVIDAD EN VGM</span><small>Información visible según tus permisos en SharePoint · últimos 7 días</small></div><div class="vgmWorkGrid" data-work-grid>${this.renderWorkModules()}</div></section>` : ''}
        </div>

        <div class="vgmModalOverlay" data-modal="clients"><div class="vgmModal"><div class="vgmModalHead"><div><h2>Mis clientes</h2><small data-client-count>${this.accessibleClients.length} clientes visibles según tus permisos</small></div><button class="vgmClose" data-action="close">×</button></div><div class="vgmClientToolbar"><input data-client-search placeholder="Buscar cliente por nombre o código…"><div class="vgmClientModeFilters">${this.renderClientModeButtons()}</div><div class="vgmAreaFilters">${this.renderAreaFilterButtons()}</div></div><div class="vgmClientList" data-client-list>${this.renderClientRows()}</div></div></div>
        <div class="vgmModalOverlay" data-modal="news"><div class="vgmModal"><div class="vgmModalHead"><h2 data-news-title>Noticia</h2><button class="vgmClose" data-action="close">×</button></div><div class="vgmCardBody" data-news-content></div></div></div>
        <div class="vgmModalOverlay" data-modal="search"><div class="vgmModal"><div class="vgmModalHead"><div><h2>Buscar en mis clientes</h2><small data-search-subtitle>Resultados visibles según tus permisos</small></div><button class="vgmClose" data-action="close">×</button></div><div class="vgmCardBody"><div class="vgmSearchResults" data-search-results></div></div></div></div>
        ${this.isAdminUser ? this.renderAdminModal() : ''}
      </div>`;
  }

  private renderMenu(items: MenuItem[]): string {
    const parents: MenuItem[] = items.filter((item: MenuItem) => item.kind === 'menu');
    if (!parents.length) return '<a href="#">Facturación</a><a href="#">RRHH</a><a href="#">Fondo Cliente</a><a href="#">Legal</a>';
    return parents.map((parent: MenuItem) => {
      const children: MenuItem[] = items.filter((item: MenuItem) => item.kind === 'submenu' && item.parent === parent.title);
      if (!children.length) return `<a href="${this.esc(parent.url || '#')}" target="_blank" rel="noopener">${this.esc(parent.title)}</a>`;
      return `<details><summary>${this.esc(parent.title)} ▾</summary><div>${children.map((child: MenuItem) => `<a href="${this.esc(child.url)}" target="_blank" rel="noopener">${this.esc(child.title)}</a>`).join('')}</div></details>`;
    }).join('');
  }

  private renderLegacyQuickLinks(): string {
    return `<section class="vgmCard"><div class="vgmQuickList">
      <a href="${VgmDataService.sourceWebUrl}/Lists/Contactos/AllItems.aspx" target="_blank" rel="noopener">Contactos</a>
      <details class="vgmQuickGroup"><summary>Organigramas</summary><a href="${VgmDataService.sourceWebUrl}/SitePages/Organigrama-VGM-Auditores.aspx">VGM Auditores</a><a href="${VgmDataService.sourceWebUrl}/SitePages/Organigrama-VGM-Outsourcing.aspx">VGM Outsourcing</a><a href="${VgmDataService.sourceWebUrl}/SitePages/Organigrama-VGM-Profesionales.aspx">VGM Profesionales</a></details>
      <details class="vgmQuickGroup"><summary>RIOHS</summary><a href="${VgmDataService.sourceWebUrl}/SitePages/RIOHS-VGM-Auditores.aspx">VGM Auditores</a><a href="${VgmDataService.sourceWebUrl}/SitePages/RIOHS-VGM-Outsourcing.aspx">VGM Outsourcing</a><a href="${VgmDataService.sourceWebUrl}/SitePages/RIOHS-VGM-Profesionales.aspx">VGM Profesionales</a></details>
    </div></section>`;
  }

  private renderFinancial(items: FinancialNewsItem[]): string {
    if (!items.length) return '<div class="vgmEmpty">No fue posible cargar Diario Financiero.</div>';
    return `<div class="vgmDfGrid">${items.slice(0,2).map((item: FinancialNewsItem) => `<a class="vgmDf" href="${this.esc(item.link)}" target="_blank" rel="noopener"><img src="${this.esc(item.image)}" alt=""><strong>${this.esc(item.title)}</strong><small>${this.esc(item.category)}</small><p>${this.esc(item.summary)}</p></a>`).join('')}</div>`;
  }

  private renderNews(featured: NewsItem | undefined, secondary: NewsItem[]): string {
    if (!featured) return '<div class="vgmEmpty">No hay noticias publicadas.</div>';
    return `<button class="vgmNewsHero" data-action="news" data-id="${featured.id}" style="border:0;width:100%;text-align:left;background-image:url('${this.cssUrl(featured.image)}')"><span class="vgmNewsOverlay"><span><h3>${this.esc(featured.title)}</h3><p>${this.esc(featured.summary)}</p></span></span></button><div class="vgmNewsSecondary">${secondary.map((item: NewsItem) => `<button data-action="news" data-id="${item.id}">${this.esc(item.title)}</button>`).join('')}</div>`;
  }

  private renderIndicators(items: IndicatorItem[]): string {
    return items.length ? items.map((item: IndicatorItem) => `<div class="vgmIndicator"><span>${this.esc(item.label)}</span><strong>${this.esc(item.value)}</strong></div>`).join('') : '<div class="vgmEmpty">Indicadores no disponibles.</div>';
  }

  private renderGallery(gallery: GalleryItem | undefined): string {
    if (!gallery || !gallery.images.length) return '<div class="vgmEmpty">No hay imágenes disponibles.</div>';
    return `<div class="vgmGallery"><img src="${this.esc(gallery.images[0])}" alt="${this.esc(gallery.title)}"><div class="vgmGalleryTitle">${this.esc(gallery.title)}</div></div>`;
  }

  private renderLinks(items: LinkItem[]): string {
    return items.length ? items.map((item: LinkItem) => `<a href="${this.esc(item.url)}" target="_blank" rel="noopener">${this.esc(item.title)}</a>`).join('') : '<div class="vgmEmpty">Sin enlaces.</div>';
  }

  private renderCalendar(date: Date, events: EventItem[]): string {
    const year: number = date.getFullYear();
    const month: number = date.getMonth();
    const first: Date = new Date(year,month,1);
    const days: number = new Date(year,month + 1,0).getDate();
    let offset: number = first.getDay() - 1;
    if (offset < 0) offset = 6;
    const eventDays: Set<number> = new Set(events.filter((event: EventItem) => event.start && event.start.getFullYear() === year && event.start.getMonth() === month).map((event: EventItem) => event.start!.getDate()));
    const cells: string[] = [];
    for (let i: number = 0; i < offset; i++) cells.push('<span class="vgmDay"></span>');
    for (let day: number = 1; day <= days; day++) cells.push(`<span class="vgmDay ${day === date.getDate() ? 'today' : ''} ${eventDays.has(day) ? 'event' : ''}">${day}</span>`);
    return `<div class="vgmCalendar">${['L','M','X','J','V','S','D'].map((d: string) => `<span class="vgmDow">${d}</span>`).join('')}${cells.join('')}</div>`;
  }

  private renderEvents(items: EventItem[]): string {
    if (!items.length) return '<ul class="vgmEvents"><li>No hay eventos próximos.</li></ul>';
    return `<ul class="vgmEvents">${items.map((item: EventItem) => `<li><strong>${this.esc(item.title)}</strong><small>${item.start ? this.capitalize(item.start.toLocaleDateString('es-CL',{weekday:'short',day:'numeric',month:'short'})) + ' a las ' + item.start.toLocaleTimeString('es-CL',{hour:'2-digit',minute:'2-digit'}) : ''}</small></li>`).join('')}</ul>`;
  }

  private renderBirthdays(items: BirthdayItem[]): string {
    if (!items.length) return '<div class="vgmEmpty">No hay cumpleaños próximos.</div>';
    return items.map((item: BirthdayItem) => {
      const photo: string = `https://vgmconsultants.sharepoint.com/_layouts/15/userphoto.aspx?size=M&accountname=${encodeURIComponent(item.email)}`;
      const next: Date = this.nextBirthday(item.date!);
      const message: string = encodeURIComponent(`¡Feliz cumpleaños ${item.name.split(' ')[0]}! Espero que tengas un gran día.`);
      return `<div class="vgmBirthday"><img src="${photo}" alt=""><span><strong>${this.esc(item.name)}</strong><small>${this.esc(item.area)} · ${next.toLocaleDateString('es-CL',{day:'numeric',month:'short'})}</small></span>${item.email ? `<a href="https://teams.microsoft.com/l/chat/0/0?users=${encodeURIComponent(item.email)}&message=${message}" target="_blank" rel="noopener">Felicitar</a>` : ''}</div>`;
    }).join('');
  }

  private renderWorkModules(): string {
    const renderers: Record<VgmActivityModuleKey,() => string> = {
      recentDocuments: (): string => `<section class="vgmCard vgmRecentCard" data-admin-module data-module-label="Documentos recientes"><div class="vgmCardHead"><h2>Documentos recientes</h2><span class="vgmSecurityBadge">Permisos aplicados</span></div><div class="vgmCardBody" data-recent-body>${this.renderRecentDocuments(this.recentDocuments)}</div></section>`,
      activityClients: (): string => `<section class="vgmCard" data-admin-module data-module-label="Clientes activos"><div class="vgmCardHead"><h2>Clientes con mayor actividad</h2></div><div class="vgmCardBody" data-activity-clients-body>${this.renderClientActivity(this.activityClients.slice(0,5))}</div></section>`,
      activityAreas: (): string => `<section class="vgmCard" data-admin-module data-module-label="Actividad por área"><div class="vgmCardHead"><h2>Actividad por área</h2></div><div class="vgmCardBody" data-activity-areas-body>${this.renderAreaActivity(this.areaActivity)}</div></section>`
    };
    return this.config.activityOrder.filter((key: VgmActivityModuleKey) => this.enabled(key)).map((key: VgmActivityModuleKey) => renderers[key]()).join('');
  }

  private renderRecentDocuments(items: VgmRecentDocument[]): string {
    if (!items.length) return '<div class="vgmEmpty">Preparando actividad de las carpetas a las que tienes acceso…</div>';
    return `<div class="vgmRecentList">${items.map((item: VgmRecentDocument) => `<a class="vgmRecentDoc" href="${this.esc(item.path)}" target="_blank" rel="noopener"><span class="vgmFileIcon ${this.esc(item.fileType)}">${this.fileAbbr(item.fileType)}</span><span class="vgmRecentText"><strong>${this.esc(item.title)}</strong><small>${this.esc(item.siteTitle || item.clientCode || 'SharePoint')}${item.area ? ` · ${this.areaLabel(item.area)}` : ''}${item.author ? ` · ${this.esc(item.author)}` : ''}</small></span><time>${item.modified ? this.relativeTime(item.modified) : ''}</time></a>`).join('')}</div>`;
  }

  private renderClientActivity(items: VgmClientActivity[]): string {
    if (!items.length) return '<div class="vgmEmpty">Preparando actividad reciente…</div>';
    const max: number = Math.max(...items.map((item: VgmClientActivity) => item.total),1);
    return `<div class="vgmActivityList">${items.map((item: VgmClientActivity,index: number) => `<a href="${this.esc(item.openUrl)}" target="_blank" rel="noopener" class="vgmActivityClient"><span class="vgmActivityRank">${String(index + 1).padStart(2,'0')}</span><span class="vgmActivityData"><strong>${this.esc(item.code)} · ${this.esc(item.name)}</strong><span class="vgmActivityBar"><i style="width:${Math.round((item.total/max)*100)}%"></i></span><small>${this.activityAreasText(item)}</small></span><b>${item.total}</b></a>`).join('')}</div>`;
  }

  private renderAreaActivity(activity: VgmAreaActivity): string {
    const entries: Array<[VgmArea,number]> = (['LEGAL','TAX','OUTSOURCING','AUDITORIA'] as VgmArea[]).map((area: VgmArea) => [area,activity[area]]);
    const total: number = entries.reduce((sum: number,item: [VgmArea,number]) => sum + item[1],0);
    const max: number = Math.max(...entries.map((item: [VgmArea,number]) => item[1]),1);
    return `<div class="vgmAreaSummary"><div class="vgmAreaTotal"><strong>${total}</strong><span>documentos modificados</span></div>${entries.map(([area,count]: [VgmArea,number]) => `<div class="vgmAreaRow"><span>${this.areaLabel(area)}</span><div><i style="width:${Math.round((count/max)*100)}%"></i></div><strong>${count}</strong></div>`).join('')}</div>`;
  }

  private renderPersonalSpace(): string {
    const favorites: VgmClientAccess[] = this.favoriteClients().slice(0,5);
    const recent: VgmRecentAccess[] = this.preferenceService.recent().slice(0,5);
    return `<section class="vgmCard"><div class="vgmCardHead"><h2>Clientes favoritos</h2><button class="vgmSecurityBadge" data-action="clients-favorites">Ver todos</button></div><div class="vgmCardBody"><div class="vgmPersonalList">${favorites.length ? favorites.map((client: VgmClientAccess) => this.renderPersonalClient(client)).join('') : '<div class="vgmEmpty">Marca clientes con ★ desde Mis clientes.</div>'}</div></div></section><section class="vgmCard"><div class="vgmCardHead"><h2>Recientes</h2></div><div class="vgmCardBody"><div class="vgmPersonalList">${recent.length ? recent.map((item: VgmRecentAccess) => `<div class="vgmPersonalClient"><span>↗</span><span><strong>${this.esc(item.code)} · ${this.esc(item.name)}</strong><small>${this.areaLabel(item.area)} · ${this.relativeTime(new Date(item.at))}</small></span><a href="${this.esc(item.url)}" data-client-area data-code="${this.esc(item.code)}" data-name="${this.esc(item.name)}" data-area="${item.area}" target="_blank" rel="noopener">Abrir</a></div>`).join('') : '<div class="vgmEmpty">Tus accesos recientes aparecerán aquí.</div>'}</div></div></section>`;
  }

  private renderPersonalClient(client: VgmClientAccess): string {
    const firstArea: VgmArea | undefined = (['LEGAL','TAX','OUTSOURCING','AUDITORIA'] as VgmArea[]).find((area: VgmArea) => Boolean(client.areas[area]));
    const url: string = firstArea ? client.areas[firstArea] || '#' : '#';
    return `<div class="vgmPersonalClient"><button class="star" data-favorite="${this.esc(client.code)}">★</button><span><strong>${this.esc(client.code)} · ${this.esc(client.name)}</strong><small>${this.clientAreasLabel(client)}</small></span>${firstArea ? `<a href="${this.esc(url)}" data-client-area data-code="${this.esc(client.code)}" data-name="${this.esc(client.name)}" data-area="${firstArea}" target="_blank" rel="noopener">Abrir</a>` : ''}</div>`;
  }

  private renderClientModeButtons(): string {
    const modes: Array<[ClientViewMode,string,number]> = [
      ['ALL','Todos',this.accessibleClients.length],
      ['FAVORITES','Favoritos',this.favoriteClients().length],
      ['RECENT','Recientes',this.recentClients().length]
    ];
    return modes.map(([mode,label,count]: [ClientViewMode,string,number]) => `<button class="vgmChip ${this.clientViewMode === mode ? 'active' : ''}" data-client-mode="${mode}">${label} <b>${count}</b></button>`).join('');
  }

  private renderAreaFilterButtons(): string {
    const filters: Array<'ALL' | VgmArea> = ['ALL','LEGAL','TAX','OUTSOURCING','AUDITORIA'];
    return filters.map((area: 'ALL' | VgmArea) => {
      const count: number = area === 'ALL' ? this.accessibleClients.length : this.accessibleClients.filter((client: VgmClientAccess) => Boolean(client.areas[area])).length;
      return `<button class="vgmChip ${this.clientAreaFilter === area ? 'active' : ''}" data-area-filter="${area}">${area === 'ALL' ? 'Todas las áreas' : this.areaLabel(area)} <b>${count}</b></button>`;
    }).join('');
  }

  private renderClientRows(): string {
    const q: string = this.normalize(this.clientSearch);
    let source: VgmClientAccess[] = this.accessibleClients.slice();
    if (this.clientViewMode === 'FAVORITES') source = this.favoriteClients();
    if (this.clientViewMode === 'RECENT') source = this.recentClients();
    const filtered: VgmClientAccess[] = source
      .filter((client: VgmClientAccess) => this.clientAreaFilter === 'ALL' || Boolean(client.areas[this.clientAreaFilter]))
      .filter((client: VgmClientAccess) => !q || this.normalize(`${client.code} ${client.name}`).includes(q));
    const counter: HTMLElement | null = this.domElement.querySelector<HTMLElement>('[data-client-count]');
    if (counter) counter.textContent = this.folderScanStarted && !this.folderScanCompleted
      ? `Preparando accesos ${this.folderScanProgress.completed}/${this.folderScanProgress.total || '…'} · ${this.folderScanProgress.found} clientes encontrados`
      : `${filtered.length} clientes visibles según tus permisos`;
    if (!filtered.length) {
      if (this.folderScanStarted && !this.folderScanCompleted) return `<div class="vgmEmpty" data-folder-scan>Preparando tus accesos en segundo plano… ${this.folderScanProgress.completed}/${this.folderScanProgress.total || '…'}.</div>`;
      return '<div class="vgmEmpty">No se encontraron clientes para este filtro.</div>';
    }
    return filtered.map((client: VgmClientAccess) => `<div class="vgmClientRow vgmClientRowAccess"><button class="vgmFavoriteBtn ${this.preferenceService.isFavorite(client.code) ? 'active' : ''}" data-favorite="${this.esc(client.code)}" title="Favorito">★</button><strong>${this.esc(client.code)}</strong><span><b>${this.esc(client.name)}</b><small>${this.esc(client.siteUrl)}</small></span><span class="vgmFolderLinks">${this.renderAccessibleAreaLink(client,'LEGAL')}${this.renderAccessibleAreaLink(client,'TAX')}${this.renderAccessibleAreaLink(client,'OUTSOURCING')}${this.renderAccessibleAreaLink(client,'AUDITORIA')}</span></div>`).join('');
  }

  private renderAccessibleAreaLink(client: VgmClientAccess, area: VgmArea): string {
    const url: string | undefined = client.areas[area];
    if (!url) return '';
    return `<a href="${this.esc(url)}" data-client-area data-code="${this.esc(client.code)}" data-name="${this.esc(client.name)}" data-area="${area}" target="_blank" rel="noopener">${this.areaLabel(area)}</a>`;
  }

  private renderAdminModal(): string {
    const modules: Array<[VgmConfigModuleKey,string]> = [
      ['calendar','Calendario'],['birthdays','Cumpleaños'],['financial','Diario Financiero'],['news','Noticias'],['indicators','Indicadores'],['gallery','Galería'],['links','Links de interés'],['personal','Mi espacio / buscador'],['recentDocuments','Documentos recientes'],['activityClients','Clientes con mayor actividad'],['activityAreas','Actividad por área']
    ];
    return `<div class="vgmModalOverlay" data-modal="admin"><div class="vgmModal vgmAdminModal"><div class="vgmModalHead"><div><h2>Administración VGM</h2><small>Configuración global de esta intranet</small></div><button class="vgmClose" data-action="close">×</button></div><div class="vgmAdminTabs"><button class="vgmAdminTab active" data-admin-tab="appearance">Apariencia</button><button class="vgmAdminTab" data-admin-tab="modules">Módulos</button><button class="vgmAdminTab" data-admin-tab="css">CSS avanzado</button><button class="vgmAdminTab" data-admin-tab="status">Estado</button></div><div class="vgmAdminBody">
      <section class="vgmAdminPanel active" data-admin-panel="appearance"><div class="vgmAdminGrid">
        ${this.adminField('Color principal','admin-primary','color',this.config.appearance.primaryColor)}
        ${this.adminField('Color secundario','admin-secondary','color',this.config.appearance.secondaryColor)}
        ${this.adminField('Fondo de portada','admin-background','color',this.config.appearance.pageBackground)}
        ${this.adminField('Radio tarjetas (px)','admin-radius','number',String(this.config.appearance.cardRadius))}
        ${this.adminField('Ancho máximo (px)','admin-width','number',String(this.config.appearance.maxWidth))}
        ${this.adminField('Tamaño base (px)','admin-font','number',String(this.config.appearance.baseFontSize))}
      </div></section>
      <section class="vgmAdminPanel" data-admin-panel="modules"><div class="vgmAdminModuleList">${modules.map(([key,label]: [VgmConfigModuleKey,string]) => `<label class="vgmAdminToggle"><span>${label}</span><input type="checkbox" data-module-toggle="${key}" ${this.enabled(key) ? 'checked' : ''}></label>`).join('')}</div><div class="vgmAdminOrder"><strong>Orden de cards de actividad</strong>${this.config.activityOrder.map((key: VgmActivityModuleKey,index: number) => `<div class="vgmOrderRow"><span>${this.activityModuleLabel(key)}</span><div><button data-order-up="${index}">↑</button><button data-order-down="${index}">↓</button></div></div>`).join('')}</div></section>
      <section class="vgmAdminPanel" data-admin-panel="css"><div class="vgmAdminField"><label>CSS personalizado</label><textarea data-admin-css placeholder=".vgmApp .vgmCard { ... }">${this.esc(this.config.customCss)}</textarea><small>Por seguridad, cada selector debe comenzar con .vgmApp. No se permiten @import, html, body ni :root.</small></div></section>
      <section class="vgmAdminPanel" data-admin-panel="status"><div class="vgmSourceStatus">${this.renderSourceStatus()}</div></section>
    </div><div class="vgmAdminFooter"><span data-admin-message>Los cambios se guardan en SharePoint para todos los usuarios.</span><div><button class="vgmAdminReset" data-action="admin-reset">Restaurar</button><button class="vgmAdminSave" data-action="admin-save">Guardar cambios</button></div></div></div></div>`;
  }

  private adminField(label: string, id: string, type: string, value: string): string {
    return `<div class="vgmAdminField"><label>${label}</label><input type="${type}" data-admin-field="${id}" value="${this.esc(value)}"></div>`;
  }

  private renderSourceStatus(): string {
    const issues: string[] = this.service.getIssues();
    const status = (label: string, ok: boolean, detail: string): string => `<div class="vgmStatusRow"><i class="vgmStatusDot ${ok ? '' : 'warn'}"></i><strong>${label}</strong><small>${this.esc(detail)}</small></div>`;
    return [
      status('Menú Principal',this.data.menu.length > 0,`${this.data.menu.length} elementos`),
      status('Noticias',this.data.news.length > 0,`${this.data.news.length} noticias`),
      status('Calendario',true,`${this.data.events.length} eventos próximos`),
      status('Cumpleaños Microsoft 365',this.graphBirthdaysAvailable,this.graphBirthdaysAvailable ? 'Graph disponible' : 'Usando fallback Contactos'),
      status('Indicadores',this.data.indicators.length > 0,`${this.data.indicators.length} indicadores`),
      status('Diario Financiero',this.data.financial.length > 0,this.data.financial.length ? `${this.data.financial.length} noticias` : 'Servicio sin respuesta'),
      status('Permisos clientes',this.accessibleClients.length > 0,`${this.accessibleClients.length} clientes visibles`),
      status('Actividad directa',this.recentDocuments.length > 0,`${this.recentDocuments.length} documentos recientes`),
      ...issues.map((issue: string) => status('Aviso',false,issue))
    ].join('');
  }

  private bindEvents(): void {
    this.domElement.addEventListener('click',(event: MouseEvent): void => {
      const target: HTMLElement = event.target as HTMLElement;
      const actionTarget: HTMLElement | null = target.closest<HTMLElement>('[data-action]');
      if (actionTarget) {
        const action: string | null = actionTarget.getAttribute('data-action');
        if (action === 'clients') this.openClients();
        if (action === 'clients-favorites') { this.clientViewMode = 'FAVORITES'; this.openClients(); }
        if (action === 'close') this.closeModals();
        if (action === 'news') this.openNews(Number(actionTarget.getAttribute('data-id')));
        if (action === 'global-search') void this.runGlobalSearch();
        if (action === 'admin') this.openModal('admin');
        if (action === 'edit-mode') this.toggleEditMode();
        if (action === 'admin-save') void this.saveAdminConfig();
        if (action === 'admin-reset') void this.resetAdminConfig();
      }

      const areaTarget: HTMLElement | null = target.closest<HTMLElement>('[data-area-filter]');
      if (areaTarget) {
        this.clientAreaFilter = (areaTarget.getAttribute('data-area-filter') || 'ALL') as 'ALL' | VgmArea;
        this.refreshClientModal();
      }
      const modeTarget: HTMLElement | null = target.closest<HTMLElement>('[data-client-mode]');
      if (modeTarget) {
        this.clientViewMode = (modeTarget.getAttribute('data-client-mode') || 'ALL') as ClientViewMode;
        this.refreshClientModal();
      }
      const favoriteTarget: HTMLElement | null = target.closest<HTMLElement>('[data-favorite]');
      if (favoriteTarget) {
        event.preventDefault();
        this.preferenceService.toggleFavorite(favoriteTarget.getAttribute('data-favorite') || '');
        this.refreshClientModal();
        this.refreshPersonalSpace();
      }
      const clientArea: HTMLElement | null = target.closest<HTMLElement>('[data-client-area]');
      if (clientArea) this.recordRecentFromElement(clientArea);
      const adminTab: HTMLElement | null = target.closest<HTMLElement>('[data-admin-tab]');
      if (adminTab) this.activateAdminTab(adminTab.getAttribute('data-admin-tab') || 'appearance');
      const up: HTMLElement | null = target.closest<HTMLElement>('[data-order-up]');
      if (up) this.moveActivityOrder(Number(up.getAttribute('data-order-up')),-1);
      const down: HTMLElement | null = target.closest<HTMLElement>('[data-order-down]');
      if (down) this.moveActivityOrder(Number(down.getAttribute('data-order-down')),1);
    });

    const clientSearch: HTMLInputElement | null = this.domElement.querySelector<HTMLInputElement>('[data-client-search]');
    clientSearch?.addEventListener('input',(): void => { this.clientSearch = clientSearch.value; this.refreshClientModal(); });
    const globalSearch: HTMLInputElement | null = this.domElement.querySelector<HTMLInputElement>('[data-global-search]');
    globalSearch?.addEventListener('keydown',(event: KeyboardEvent): void => { if (event.key === 'Enter') void this.runGlobalSearch(); });
  }

  private openClients(): void {
    this.openModal('clients');
    this.refreshClientModal();
    if (!this.accessibleClients.length && !this.folderScanStarted && !this.folderScanCompleted) void this.preloadFolderAccess();
  }

  private async preloadFolderAccess(): Promise<void> {
    if (this.accessibleClients.length || this.folderScanCompleted || this.folderScanStarted) return;
    this.folderScanStarted = true;
    this.folderScanProgress = { completed: 0, total: 0, found: 0 };
    this.refreshClientsButton();
    try {
      const directClients: VgmClientAccess[] = await this.folderAccessService.scanAccessibleClients((completed: number,total: number,found: number): void => {
        this.folderScanProgress = { completed, total, found };
        const modal: HTMLElement | null = this.domElement.querySelector<HTMLElement>('[data-modal="clients"]');
        if (modal?.classList.contains('open')) this.refreshClientModal();
      },10);
      this.accessibleClients = directClients;
      this.folderScanCompleted = true;
      this.saveAccessCache(directClients);
      this.refreshClientModal();
      this.refreshClientsButton();
      this.refreshPersonalSpace();
      void this.refreshDirectActivity();
    } catch (error) {
      console.warn('VGM: no fue posible precargar los permisos directos de las carpetas.',error);
      this.folderScanCompleted = true;
    } finally {
      this.folderScanStarted = false;
      this.refreshClientsButton();
    }
  }

  private async refreshDirectActivity(): Promise<void> {
    if (!this.accessibleClients.length || !this.hasAnyActivityModule()) return;
    try {
      const activity = await this.directActivityService.getActivity(this.accessibleClients,7,8,8);
      if (activity.documents.length) this.recentDocuments = activity.documents;
      this.activityClients = activity.clients;
      this.areaActivity = activity.areas;
      this.refreshActivityCards();
    } catch (error) {
      console.warn('VGM: no fue posible construir actividad directa.',error);
    }
  }

  private refreshActivityCards(): void {
    const recent: HTMLElement | null = this.domElement.querySelector<HTMLElement>('[data-recent-body]');
    const clients: HTMLElement | null = this.domElement.querySelector<HTMLElement>('[data-activity-clients-body]');
    const areas: HTMLElement | null = this.domElement.querySelector<HTMLElement>('[data-activity-areas-body]');
    if (recent) recent.innerHTML = this.renderRecentDocuments(this.recentDocuments);
    if (clients) clients.innerHTML = this.renderClientActivity(this.activityClients.slice(0,5));
    if (areas) areas.innerHTML = this.renderAreaActivity(this.areaActivity);
  }

  private async runGlobalSearch(): Promise<void> {
    const input: HTMLInputElement | null = this.domElement.querySelector<HTMLInputElement>('[data-global-search]');
    const q: string = (input?.value || '').trim();
    if (q.length < 2) { this.toast('Escribe al menos 2 caracteres.'); return; }
    this.openModal('search');
    const results: HTMLElement | null = this.domElement.querySelector<HTMLElement>('[data-search-results]');
    const subtitle: HTMLElement | null = this.domElement.querySelector<HTMLElement>('[data-search-subtitle]');
    if (results) results.innerHTML = '<div class="vgmEmpty">Buscando documentos y clientes…</div>';
    if (subtitle) subtitle.textContent = `Buscando “${q}” dentro de tus accesos`;
    const matchingClients: VgmClientAccess[] = this.accessibleClients.filter((client: VgmClientAccess) => this.normalize(`${client.code} ${client.name}`).includes(this.normalize(q))).slice(0,10);
    let docs: VgmRecentDocument[] = [];
    try { docs = await this.globalSearchService.searchDocuments(q,this.accessibleClients,30); } catch { /* keep clients */ }
    if (!results) return;
    const clientHtml: string = matchingClients.map((client: VgmClientAccess) => {
      const area: VgmArea | undefined = (['LEGAL','TAX','OUTSOURCING','AUDITORIA'] as VgmArea[]).find((value: VgmArea) => Boolean(client.areas[value]));
      const url: string = area ? client.areas[area] || '#' : '#';
      return `<a class="vgmSearchResult" href="${this.esc(url)}" target="_blank" rel="noopener"><span class="vgmFileIcon">CLI</span><span><strong>${this.esc(client.code)} · ${this.esc(client.name)}</strong><small>Cliente · ${this.clientAreasLabel(client)}</small></span><time>Abrir</time></a>`;
    }).join('');
    const docHtml: string = docs.map((doc: VgmRecentDocument) => `<a class="vgmSearchResult" href="${this.esc(doc.path)}" target="_blank" rel="noopener"><span class="vgmFileIcon ${this.esc(doc.fileType)}">${this.fileAbbr(doc.fileType)}</span><span><strong>${this.esc(doc.title)}</strong><small>${this.esc(doc.siteTitle)}${doc.area ? ` · ${this.areaLabel(doc.area)}` : ''}${doc.author ? ` · ${this.esc(doc.author)}` : ''}</small></span><time>${doc.modified ? this.relativeTime(doc.modified) : ''}</time></a>`).join('');
    results.innerHTML = clientHtml + docHtml || '<div class="vgmEmpty">No se encontraron resultados visibles para tu usuario.</div>';
    if (subtitle) subtitle.textContent = `${matchingClients.length + docs.length} resultados visibles`;
  }

  private async saveAdminConfig(): Promise<void> {
    if (!this.isAdminUser) return;
    const customCss: string = (this.domElement.querySelector<HTMLTextAreaElement>('[data-admin-css]')?.value || '').trim();
    const validation: string | undefined = this.validateCustomCss(customCss);
    if (validation) { this.setAdminMessage(validation); return; }
    const next: VgmPortalConfig = JSON.parse(JSON.stringify(this.config)) as VgmPortalConfig;
    next.appearance.primaryColor = this.adminInput('admin-primary',next.appearance.primaryColor);
    next.appearance.secondaryColor = this.adminInput('admin-secondary',next.appearance.secondaryColor);
    next.appearance.pageBackground = this.adminInput('admin-background',next.appearance.pageBackground);
    next.appearance.cardRadius = this.adminNumber('admin-radius',next.appearance.cardRadius,0,30);
    next.appearance.maxWidth = this.adminNumber('admin-width',next.appearance.maxWidth,900,1900);
    next.appearance.baseFontSize = this.adminNumber('admin-font',next.appearance.baseFontSize,10,20);
    this.domElement.querySelectorAll<HTMLInputElement>('[data-module-toggle]').forEach((input: HTMLInputElement) => {
      const key: VgmConfigModuleKey = input.getAttribute('data-module-toggle') as VgmConfigModuleKey;
      next.modules[key] = input.checked;
    });
    next.activityOrder = this.config.activityOrder.slice();
    next.customCss = customCss;
    this.setAdminMessage('Guardando…');
    try {
      await this.adminConfigService.save(next);
      this.config = next;
      this.toast('Configuración guardada para todos los usuarios.');
      this.paint();
      this.bindEvents();
      this.openModal('admin');
      if (this.accessibleClients.length) void this.refreshDirectActivity();
    } catch (error) {
      this.setAdminMessage(`No fue posible guardar: ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  private async resetAdminConfig(): Promise<void> {
    if (!this.isAdminUser) return;
    const next: VgmPortalConfig = JSON.parse(JSON.stringify(DEFAULT_VGM_CONFIG)) as VgmPortalConfig;
    try {
      await this.adminConfigService.save(next);
      this.config = next;
      this.toast('Configuración restaurada.');
      this.paint();
      this.bindEvents();
      this.openModal('admin');
    } catch (error) {
      this.setAdminMessage(`No fue posible restaurar: ${error instanceof Error ? error.message : String(error)}`);
    }
  }

  private activateAdminTab(tab: string): void {
    this.domElement.querySelectorAll<HTMLElement>('[data-admin-tab]').forEach((button: HTMLElement) => button.classList.toggle('active',button.getAttribute('data-admin-tab') === tab));
    this.domElement.querySelectorAll<HTMLElement>('[data-admin-panel]').forEach((panel: HTMLElement) => panel.classList.toggle('active',panel.getAttribute('data-admin-panel') === tab));
  }

  private moveActivityOrder(index: number, delta: number): void {
    const target: number = index + delta;
    if (index < 0 || target < 0 || index >= this.config.activityOrder.length || target >= this.config.activityOrder.length) return;
    const next: VgmActivityModuleKey[] = this.config.activityOrder.slice();
    const item: VgmActivityModuleKey = next[index];
    next.splice(index,1);
    next.splice(target,0,item);
    this.config.activityOrder = next;
    const panel: HTMLElement | null = this.domElement.querySelector<HTMLElement>('[data-admin-panel="modules"]');
    if (panel) {
      const order: HTMLElement | null = panel.querySelector<HTMLElement>('.vgmAdminOrder');
      if (order) order.innerHTML = `<strong>Orden de cards de actividad</strong>${this.config.activityOrder.map((key: VgmActivityModuleKey,i: number) => `<div class="vgmOrderRow"><span>${this.activityModuleLabel(key)}</span><div><button data-order-up="${i}">↑</button><button data-order-down="${i}">↓</button></div></div>`).join('')}`;
    }
  }

  private toggleEditMode(): void {
    this.editMode = !this.editMode;
    const root: HTMLElement | null = this.domElement.querySelector<HTMLElement>('.vgmApp');
    root?.classList.toggle('vgmEditMode',this.editMode);
    const button: HTMLButtonElement | null = this.domElement.querySelector<HTMLButtonElement>('[data-action="edit-mode"]');
    button?.classList.toggle('active',this.editMode);
  }

  private refreshClientModal(): void {
    const modes: HTMLElement | null = this.domElement.querySelector<HTMLElement>('.vgmClientModeFilters');
    const filters: HTMLElement | null = this.domElement.querySelector<HTMLElement>('.vgmAreaFilters');
    const rows: HTMLElement | null = this.domElement.querySelector<HTMLElement>('[data-client-list]');
    if (modes) modes.innerHTML = this.renderClientModeButtons();
    if (filters) filters.innerHTML = this.renderAreaFilterButtons();
    if (rows) rows.innerHTML = this.renderClientRows();
  }

  private refreshClientsButton(): void {
    const button: HTMLButtonElement | null = this.domElement.querySelector<HTMLButtonElement>('[data-action="clients"]');
    if (!button) return;
    if (this.folderScanStarted && !this.folderScanCompleted && !this.accessibleClients.length) button.textContent = '☰ Mis clientes (…)';
    else button.textContent = `☰ Mis clientes${this.accessibleClients.length ? ` (${this.accessibleClients.length})` : ''}`;
  }

  private refreshPersonalSpace(): void {
    const grid: HTMLElement | null = this.domElement.querySelector<HTMLElement>('[data-personal-grid]');
    if (grid) grid.innerHTML = this.renderPersonalSpace();
  }

  private openNews(id: number): void {
    const item: NewsItem | undefined = this.data.news.find((news: NewsItem) => news.id === id);
    if (!item) return;
    const title: HTMLElement | null = this.domElement.querySelector<HTMLElement>('[data-news-title]');
    const content: HTMLElement | null = this.domElement.querySelector<HTMLElement>('[data-news-content]');
    if (title) title.textContent = item.title;
    if (content) content.innerHTML = `<img src="${this.esc(item.image)}" alt="" style="width:100%;max-height:300px;object-fit:cover;margin-bottom:16px"><div>${item.body}</div>`;
    this.openModal('news');
  }

  private openModal(name: string): void {
    this.domElement.querySelector<HTMLElement>(`[data-modal="${name}"]`)?.classList.add('open');
  }

  private closeModals(): void {
    this.domElement.querySelectorAll<HTMLElement>('.vgmModalOverlay').forEach((element: HTMLElement) => element.classList.remove('open'));
  }

  private recordRecentFromElement(element: HTMLElement): void {
    const code: string = element.getAttribute('data-code') || '';
    const name: string = element.getAttribute('data-name') || '';
    const area: VgmArea = (element.getAttribute('data-area') || 'LEGAL') as VgmArea;
    const url: string = element.getAttribute('href') || '';
    if (!code || !url) return;
    this.preferenceService.addRecent({ code, name, area, url, at: Date.now() });
    this.refreshPersonalSpace();
  }

  private favoriteClients(): VgmClientAccess[] {
    const favorites: string[] = this.preferenceService.favorites();
    return favorites.map((code: string) => this.accessibleClients.find((client: VgmClientAccess) => client.code === code)).filter((client: VgmClientAccess | undefined): client is VgmClientAccess => Boolean(client));
  }

  private recentClients(): VgmClientAccess[] {
    const seen: Set<string> = new Set<string>();
    const result: VgmClientAccess[] = [];
    for (const recent of this.preferenceService.recent()) {
      if (seen.has(recent.code)) continue;
      const client: VgmClientAccess | undefined = this.accessibleClients.find((item: VgmClientAccess) => item.code === recent.code);
      if (client) { result.push(client); seen.add(recent.code); }
    }
    return result;
  }

  private clientAreasLabel(client: VgmClientAccess): string {
    return (['LEGAL','TAX','OUTSOURCING','AUDITORIA'] as VgmArea[]).filter((area: VgmArea) => Boolean(client.areas[area])).map((area: VgmArea) => this.areaLabel(area)).join(' · ');
  }

  private enabled(key: VgmConfigModuleKey): boolean {
    return this.config.modules[key] !== false;
  }

  private hasAnyActivityModule(): boolean {
    return this.enabled('recentDocuments') || this.enabled('activityClients') || this.enabled('activityAreas');
  }

  private themeCss(): string {
    const a = this.config.appearance;
    return `.vgmApp{--vgm-primary:${a.primaryColor};--vgm-secondary:${a.secondaryColor};background:${a.pageBackground}!important;font-size:${a.baseFontSize}px}.vgmApp .vgmWrap{max-width:${a.maxWidth}px!important}.vgmApp .vgmCard{border-radius:${a.cardRadius}px!important}.vgmApp .vgmClientsBtn,.vgmApp .vgmTicket,.vgmApp .vgmChip.active,.vgmApp .vgmAdminSave{background:${a.primaryColor}!important}.vgmApp .vgmAreaTotal strong,.vgmApp .vgmActivityClient>b{color:${a.primaryColor}!important}`;
  }

  private validateCustomCss(css: string): string | undefined {
    if (!css) return undefined;
    if (/@import|javascript:|<\/style|<script/i.test(css)) return 'CSS rechazado: contiene una regla no permitida.';
    if (/(^|[,{\s])(html|body|:root)([\s,{.#[:]|$)/i.test(css)) return 'CSS rechazado: no se permiten selectores globales html, body o :root.';
    const blocks: string[] = css.split('}').map((block: string) => block.trim()).filter(Boolean);
    for (const block of blocks) {
      const selector: string = block.split('{')[0]?.trim() || '';
      if (!selector || selector.startsWith('@')) continue;
      const selectors: string[] = selector.split(',').map((value: string) => value.trim());
      if (selectors.some((value: string) => !value.startsWith('.vgmApp'))) return 'Cada selector del CSS personalizado debe comenzar con .vgmApp';
    }
    return undefined;
  }

  private safeStoredCss(css: string): string {
    return this.validateCustomCss(css || '') ? '' : (css || '');
  }

  private adminInput(name: string, fallback: string): string {
    return this.domElement.querySelector<HTMLInputElement>(`[data-admin-field="${name}"]`)?.value || fallback;
  }

  private adminNumber(name: string, fallback: number, min: number, max: number): number {
    const value: number = Number(this.adminInput(name,String(fallback)));
    return Number.isFinite(value) ? Math.max(min,Math.min(max,value)) : fallback;
  }

  private setAdminMessage(message: string): void {
    const element: HTMLElement | null = this.domElement.querySelector<HTMLElement>('[data-admin-message]');
    if (element) element.textContent = message;
  }

  private activityModuleLabel(key: VgmActivityModuleKey): string {
    if (key === 'recentDocuments') return 'Documentos recientes';
    if (key === 'activityClients') return 'Clientes con mayor actividad';
    return 'Actividad por área';
  }

  private async loadMicrosoft365Birthdays(): Promise<BirthdayItem[]> {
    try {
      const client: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');
      const response: { value?: Array<{ displayName?: string; mail?: string; userPrincipalName?: string; department?: string; birthday?: string }> } = await client.api('/users').select('displayName,mail,userPrincipalName,department,birthday').top(999).get() as { value?: Array<{ displayName?: string; mail?: string; userPrincipalName?: string; department?: string; birthday?: string }> };
      this.graphBirthdaysAvailable = true;
      return (response.value || []).map((user) => ({ name: user.displayName || '', email: user.mail || user.userPrincipalName || '', area: user.department || '', date: user.birthday ? new Date(user.birthday) : undefined })).filter((item: BirthdayItem) => Boolean(item.name && item.date && !Number.isNaN(item.date.getTime())));
    } catch (error) {
      this.graphBirthdaysAvailable = false;
      console.info('VGM: cumpleaños Microsoft 365 no disponibles; se utilizará la lista Contactos.',error);
      return [];
    }
  }

  private mergeBirthdays(graphItems: BirthdayItem[], sharePointItems: BirthdayItem[]): BirthdayItem[] {
    const map: Map<string,BirthdayItem> = new Map<string,BirthdayItem>();
    for (const item of graphItems) map.set(this.personKey(item),item);
    for (const item of sharePointItems) {
      const key: string = this.personKey(item);
      if (!map.has(key)) map.set(key,item);
      else {
        const existing: BirthdayItem = map.get(key)!;
        map.set(key,{ ...existing, email: existing.email || item.email, area: existing.area || item.area });
      }
    }
    return Array.from(map.values()).filter((item: BirthdayItem) => Boolean(item.date));
  }

  private personKey(item: BirthdayItem): string {
    return this.normalize(item.email || item.name);
  }

  private upcomingBirthdays(items: BirthdayItem[]): BirthdayItem[] {
    return items.slice().sort((a: BirthdayItem,b: BirthdayItem) => this.nextBirthday(a.date!).getTime() - this.nextBirthday(b.date!).getTime());
  }

  private nextBirthday(date: Date): Date {
    const now: Date = new Date();
    let next: Date = new Date(now.getFullYear(),date.getMonth(),date.getDate());
    if (next.getTime() < new Date(now.getFullYear(),now.getMonth(),now.getDate()).getTime()) next = new Date(now.getFullYear() + 1,date.getMonth(),date.getDate());
    return next;
  }

  private activityAreasText(item: VgmClientActivity): string {
    return (['LEGAL','TAX','OUTSOURCING','AUDITORIA'] as VgmArea[]).filter((area: VgmArea) => item.areas[area] > 0).map((area: VgmArea) => `${this.areaLabel(area)} ${item.areas[area]}`).join(' · ');
  }

  private areaLabel(area: VgmArea): string {
    if (area === 'LEGAL') return 'Legal';
    if (area === 'TAX') return 'Tax';
    if (area === 'OUTSOURCING') return 'Outsourcing';
    return 'Auditoría';
  }

  private fileAbbr(fileType: string): string {
    const type: string = (fileType || '').toLowerCase();
    if (type === 'doc' || type === 'docx') return 'W';
    if (type === 'xls' || type === 'xlsx' || type === 'csv') return 'X';
    if (type === 'ppt' || type === 'pptx') return 'P';
    if (type === 'pdf') return 'PDF';
    return type ? type.slice(0,3).toUpperCase() : 'DOC';
  }

  private relativeTime(date: Date): string {
    const seconds: number = Math.max(0,Math.floor((Date.now() - date.getTime()) / 1000));
    if (seconds < 60) return 'Ahora';
    const minutes: number = Math.floor(seconds / 60);
    if (minutes < 60) return `Hace ${minutes} min`;
    const hours: number = Math.floor(minutes / 60);
    if (hours < 24) return `Hace ${hours} h`;
    const days: number = Math.floor(hours / 24);
    if (days === 1) return 'Ayer';
    if (days < 7) return `Hace ${days} días`;
    return date.toLocaleDateString('es-CL',{day:'numeric',month:'short'});
  }

  private loadAccessCache(): VgmClientAccess[] {
    try {
      const email: string = (this.context.pageContext.user.email || 'anon').toLowerCase();
      const raw: string | null = window.sessionStorage.getItem(`vgm-intranet:access:${email}`);
      if (!raw) return [];
      const payload: { at: number; clients: VgmClientAccess[] } = JSON.parse(raw) as { at: number; clients: VgmClientAccess[] };
      if (!payload.at || Date.now() - payload.at > 15 * 60 * 1000) return [];
      return Array.isArray(payload.clients) ? payload.clients : [];
    } catch { return []; }
  }

  private saveAccessCache(clients: VgmClientAccess[]): void {
    try {
      const email: string = (this.context.pageContext.user.email || 'anon').toLowerCase();
      window.sessionStorage.setItem(`vgm-intranet:access:${email}`,JSON.stringify({ at: Date.now(), clients }));
    } catch { /* no-op */ }
  }

  private toast(message: string): void {
    const current: HTMLElement | null = this.domElement.querySelector<HTMLElement>('.vgmToast');
    current?.remove();
    const toast: HTMLDivElement = document.createElement('div');
    toast.className = 'vgmToast';
    toast.textContent = message;
    this.domElement.appendChild(toast);
    window.setTimeout(() => toast.remove(),3200);
  }

  private normalize(value: string): string {
    return (value || '').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase();
  }

  private capitalize(value: string): string {
    return value ? value.charAt(0).toUpperCase() + value.slice(1) : value;
  }

  private esc(value: string): string {
    return String(value || '').replace(/[&<>"']/g,(char: string) => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[char] || char));
  }

  private cssUrl(value: string): string {
    return this.esc(value).replace(/'/g,'%27');
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}
