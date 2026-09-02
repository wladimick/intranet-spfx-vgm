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

export default class VgmIntranetWebPart extends BaseClientSideWebPart<IVgmIntranetWebPartProps> {
  private service!: VgmDataService;
  private searchService!: VgmSearchService;
  private folderAccessService!: VgmFolderAccessService;
  private accessibleClients: VgmClientAccess[] = [];
  private recentDocuments: VgmRecentDocument[] = [];
  private activityClients: VgmClientActivity[] = [];
  private areaActivity: VgmAreaActivity = { LEGAL: 0, TAX: 0, OUTSOURCING: 0, AUDITORIA: 0 };
  private clientSearch: string = '';
  private clientAreaFilter: 'ALL' | VgmArea = 'ALL';
  private folderScanStarted: boolean = false;
  private folderScanCompleted: boolean = false;
  private data: DashboardData = { menu: [], birthdays: [], events: [], news: [], links: [], indicators: [], financial: [] };

  public render(): void {
    void this.renderAsync();
  }

  private async renderAsync(): Promise<void> {
    this.domElement.innerHTML = `<style>${VGM_STYLES}</style><div class="vgmApp"><div class="vgmLoading">Cargando intranet VGM…</div></div>`;
    this.service = new VgmDataService(this.context.spHttpClient);
    this.searchService = new VgmSearchService(this.context.spHttpClient);
    this.folderAccessService = new VgmFolderAccessService(this.context.spHttpClient,this.service);

    const [menu, sharePointBirthdays, events, news, links, gallery, indicators, financial, accessibleClients, recentDocuments, activity, graphBirthdays] = await Promise.all([
      this.service.getMenu(),
      this.service.getBirthdays(),
      this.service.getEvents(),
      this.service.getInternalNews(),
      this.service.getLinks(),
      this.service.getGallery(),
      this.service.getIndicators(),
      this.service.getFinancialNews(),
      this.searchService.getAccessibleClients().catch((error: unknown): VgmClientAccess[] => { console.warn('VGM: no se pudieron obtener clientes visibles por Search.',error); return []; }),
      this.searchService.getRecentDocuments(8).catch((error: unknown): VgmRecentDocument[] => { console.warn('VGM: no se pudieron obtener documentos recientes.',error); return []; }),
      this.searchService.getActivity(7,1200).catch((error: unknown) => { console.warn('VGM: no se pudo calcular actividad.',error); return { clients: [] as VgmClientActivity[], areas: { LEGAL:0,TAX:0,OUTSOURCING:0,AUDITORIA:0 } as VgmAreaActivity, documents: [] as VgmRecentDocument[] }; }),
      this.loadMicrosoft365Birthdays()
    ]);

    this.accessibleClients = accessibleClients;
    this.recentDocuments = recentDocuments;
    this.activityClients = activity.clients;
    this.areaActivity = activity.areas;
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
  }

  private paint(): void {
    const displayName: string = this.context.pageContext.user.displayName || 'Usuario VGM';
    const email: string = this.context.pageContext.user.email || '';
    const photo: string = `https://vgmconsultants.sharepoint.com/_layouts/15/userphoto.aspx?size=L&accountname=${encodeURIComponent(email)}`;
    const today: Date = new Date();
    const nextBirthdays: BirthdayItem[] = this.upcomingBirthdays(this.data.birthdays).slice(0,5);
    const featured: NewsItem | undefined = this.data.news[0];
    const secondary: NewsItem[] = this.data.news.slice(1,5);

    this.domElement.innerHTML = `
      <style>${VGM_STYLES}</style>
      <div class="vgmApp">
        <div class="vgmWrap">
          <header class="vgmHeader">
            <div class="vgmHeaderTop">
              <img class="vgmLogo" src="${VgmDataService.sourceWebUrl}/SiteAssets/portada/imagenes/LOGO-VGM-BLANCO.webp" alt="VGM Consultores">
              <div class="vgmSpacer"></div>
              <div class="vgmUser"><span><small>Bienvenido</small><strong>${this.esc(displayName)}</strong></span><img src="${photo}" alt="${this.esc(displayName)}"></div>
            </div>
            <nav class="vgmMenu">${this.renderMenu(this.data.menu)}<button class="vgmClientsBtn" data-action="clients">☰ Mis clientes${this.accessibleClients.length ? ` (${this.accessibleClients.length})` : ''}</button></nav>
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
                  <section class="vgmCard"><div class="vgmCardHead"><h2>${this.capitalize(today.toLocaleDateString('es-CL',{month:'long'}))} ${today.getFullYear()}</h2></div><div class="vgmCardBody">${this.renderCalendar(today,this.data.events)}${this.renderEvents(this.data.events.slice(0,3))}</div></section>
                  <section class="vgmCard"><div class="vgmCardHead"><h2>Próximos Cumpleaños</h2></div><div class="vgmCardBody">${this.renderBirthdays(nextBirthdays)}</div></section>
                </div>
              </div>
              <section class="vgmCard vgmFinancialCard"><div class="vgmCardHead"><h2>Diario Financiero</h2></div><div class="vgmCardBody">${this.renderFinancial(this.data.financial)}</div></section>
            </section>

            <section class="vgmRightArea">
              <section class="vgmCard vgmNewsCard"><div class="vgmCardHead"><h2>Noticias</h2></div>${this.renderNews(featured,secondary)}</section>
              <section class="vgmCard"><div class="vgmCardHead"><h2>Indicadores Económicos</h2></div><div class="vgmCardBody"><div class="vgmIndicators">${this.renderIndicators(this.data.indicators)}</div></div></section>
              <div class="vgmBottom">
                <section class="vgmCard"><div class="vgmCardHead"><h2>Galerías</h2></div><div class="vgmCardBody">${this.renderGallery(this.data.gallery)}</div></section>
                <section class="vgmCard"><div class="vgmCardHead"><h2>Links de Interés</h2></div><div class="vgmCardBody vgmLinks">${this.renderLinks(this.data.links)}</div></section>
              </div>
            </section>
          </div>

          <section class="vgmWorkSection">
            <div class="vgmWorkTitle"><span>MI ACTIVIDAD EN VGM</span><small>Información visible según tus permisos en SharePoint · últimos 7 días</small></div>
            <div class="vgmWorkGrid">
              <section class="vgmCard vgmRecentCard"><div class="vgmCardHead"><h2>Documentos recientes</h2><span class="vgmSecurityBadge">Permisos aplicados</span></div><div class="vgmCardBody">${this.renderRecentDocuments(this.recentDocuments)}</div></section>
              <section class="vgmCard"><div class="vgmCardHead"><h2>Clientes con mayor actividad</h2></div><div class="vgmCardBody">${this.renderClientActivity(this.activityClients.slice(0,5))}</div></section>
              <section class="vgmCard"><div class="vgmCardHead"><h2>Actividad por área</h2></div><div class="vgmCardBody">${this.renderAreaActivity(this.areaActivity)}</div></section>
            </div>
          </section>
        </div>

        <div class="vgmModalOverlay" data-modal="clients"><div class="vgmModal"><div class="vgmModalHead"><div><h2>Mis clientes</h2><small data-client-count>${this.accessibleClients.length} clientes visibles según tus permisos</small></div><button class="vgmClose" data-action="close">×</button></div><div class="vgmClientToolbar"><input data-client-search placeholder="Buscar cliente por nombre o código…"><div class="vgmAreaFilters">${this.renderAreaFilterButtons()}</div></div><div class="vgmClientList" data-client-list>${this.renderClientRows()}</div></div></div>
        <div class="vgmModalOverlay" data-modal="news"><div class="vgmModal"><div class="vgmModalHead"><h2 data-news-title>Noticia</h2><button class="vgmClose" data-action="close">×</button></div><div class="vgmCardBody" data-news-content></div></div></div>
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

  private renderRecentDocuments(items: VgmRecentDocument[]): string {
    if (!items.length) return '<div class="vgmEmpty">No se encontraron documentos recientes visibles para tu usuario.</div>';
    return `<div class="vgmRecentList">${items.map((item: VgmRecentDocument) => `<a class="vgmRecentDoc" href="${this.esc(item.path)}" target="_blank" rel="noopener"><span class="vgmFileIcon ${this.esc(item.fileType)}">${this.fileAbbr(item.fileType)}</span><span class="vgmRecentText"><strong>${this.esc(item.title)}</strong><small>${this.esc(item.siteTitle || item.clientCode || 'SharePoint')}${item.area ? ` · ${this.areaLabel(item.area)}` : ''}${item.author ? ` · ${this.esc(item.author)}` : ''}</small></span><time>${item.modified ? this.relativeTime(item.modified) : ''}</time></a>`).join('')}</div>`;
  }

  private renderClientActivity(items: VgmClientActivity[]): string {
    if (!items.length) return '<div class="vgmEmpty">No se detectó actividad reciente en las carpetas que puedes consultar.</div>';
    const max: number = Math.max(...items.map((item: VgmClientActivity) => item.total),1);
    return `<div class="vgmActivityList">${items.map((item: VgmClientActivity,index: number) => `<a href="${this.esc(item.openUrl)}" target="_blank" rel="noopener" class="vgmActivityClient"><span class="vgmActivityRank">${String(index + 1).padStart(2,'0')}</span><span class="vgmActivityData"><strong>${this.esc(item.code)} · ${this.esc(item.name)}</strong><span class="vgmActivityBar"><i style="width:${Math.round((item.total/max)*100)}%"></i></span><small>${this.activityAreasText(item)}</small></span><b>${item.total}</b></a>`).join('')}</div>`;
  }

  private renderAreaActivity(activity: VgmAreaActivity): string {
    const entries: Array<[VgmArea,number]> = (['LEGAL','TAX','OUTSOURCING','AUDITORIA'] as VgmArea[]).map((area: VgmArea) => [area,activity[area]]);
    const total: number = entries.reduce((sum: number,item: [VgmArea,number]) => sum + item[1],0);
    const max: number = Math.max(...entries.map((item: [VgmArea,number]) => item[1]),1);
    return `<div class="vgmAreaSummary"><div class="vgmAreaTotal"><strong>${total}</strong><span>documentos modificados</span></div>${entries.map(([area,count]: [VgmArea,number]) => `<div class="vgmAreaRow"><span>${this.areaLabel(area)}</span><div><i style="width:${Math.round((count/max)*100)}%"></i></div><strong>${count}</strong></div>`).join('')}</div>`;
  }

  private renderAreaFilterButtons(): string {
    const filters: Array<'ALL' | VgmArea> = ['ALL','LEGAL','TAX','OUTSOURCING','AUDITORIA'];
    return filters.map((area: 'ALL' | VgmArea) => {
      const count: number = area === 'ALL' ? this.accessibleClients.length : this.accessibleClients.filter((client: VgmClientAccess) => Boolean(client.areas[area])).length;
      return `<button class="vgmChip ${this.clientAreaFilter === area ? 'active' : ''}" data-area-filter="${area}">${area === 'ALL' ? 'Todos' : this.areaLabel(area)} <b>${count}</b></button>`;
    }).join('');
  }

  private renderClientRows(): string {
    const q: string = this.normalize(this.clientSearch);
    const filtered: VgmClientAccess[] = this.accessibleClients
      .filter((client: VgmClientAccess) => this.clientAreaFilter === 'ALL' || Boolean(client.areas[this.clientAreaFilter]))
      .filter((client: VgmClientAccess) => !q || this.normalize(`${client.code} ${client.name}`).includes(q));
    const counter: HTMLElement | null = this.domElement.querySelector<HTMLElement>('[data-client-count]');
    if (counter) counter.textContent = `${filtered.length} clientes visibles según tus permisos`;
    if (!filtered.length) return '<div class="vgmEmpty">No se encontraron clientes para este filtro.</div>';
    return filtered.map((client: VgmClientAccess) => `<div class="vgmClientRow vgmClientRowAccess"><strong>${this.esc(client.code)}</strong><span><b>${this.esc(client.name)}</b><small>${this.esc(client.siteUrl)}</small></span><span class="vgmFolderLinks">${this.renderAccessibleAreaLink(client,'LEGAL')}${this.renderAccessibleAreaLink(client,'TAX')}${this.renderAccessibleAreaLink(client,'OUTSOURCING')}${this.renderAccessibleAreaLink(client,'AUDITORIA')}</span></div>`).join('');
  }

  private renderAccessibleAreaLink(client: VgmClientAccess, area: VgmArea): string {
    const url: string | undefined = client.areas[area];
    if (!url) return '';
    return `<a href="${this.esc(url)}" target="_blank" rel="noopener">${this.areaLabel(area)}</a>`;
  }

  private bindEvents(): void {
    this.domElement.addEventListener('click',(event: MouseEvent): void => {
      const actionTarget: HTMLElement | null = (event.target as HTMLElement).closest<HTMLElement>('[data-action]');
      if (actionTarget) {
        const action: string | null = actionTarget.getAttribute('data-action');
        if (action === 'clients') void this.openClients();
        if (action === 'close') this.closeModals();
        if (action === 'news') this.openNews(Number(actionTarget.getAttribute('data-id')));
      }
      const areaTarget: HTMLElement | null = (event.target as HTMLElement).closest<HTMLElement>('[data-area-filter]');
      if (areaTarget) {
        this.clientAreaFilter = (areaTarget.getAttribute('data-area-filter') || 'ALL') as 'ALL' | VgmArea;
        this.refreshClientModal();
      }
    });
    const search: HTMLInputElement | null = this.domElement.querySelector<HTMLInputElement>('[data-client-search]');
    search?.addEventListener('input',(): void => { this.clientSearch = search.value; this.refreshClientModal(); });
  }

  private async openClients(): Promise<void> {
    this.openModal('clients');
    if (this.accessibleClients.length || this.folderScanCompleted || this.folderScanStarted) return;
    this.folderScanStarted = true;
    const rows: HTMLElement | null = this.domElement.querySelector<HTMLElement>('[data-client-list]');
    const counter: HTMLElement | null = this.domElement.querySelector<HTMLElement>('[data-client-count]');
    if (rows) rows.innerHTML = '<div class="vgmEmpty" data-folder-scan>Search no encontró accesos. Verificando permisos directamente en las carpetas…</div>';
    try {
      const directClients: VgmClientAccess[] = await this.folderAccessService.scanAccessibleClients((completed: number,total: number,found: number): void => {
        if (counter) counter.textContent = `Verificando carpetas ${completed}/${total} · ${found} clientes encontrados`;
        const status: HTMLElement | null = this.domElement.querySelector<HTMLElement>('[data-folder-scan]');
        if (status) status.textContent = `Comprobando permisos reales de SharePoint… ${completed}/${total}. Clientes encontrados: ${found}.`;
      },10);
      this.accessibleClients = directClients;
      this.folderScanCompleted = true;
      this.refreshClientModal();
      this.refreshClientsButton();
    } catch (error) {
      console.warn('VGM: no fue posible verificar los permisos directos de las carpetas.',error);
      this.folderScanCompleted = true;
      if (rows) rows.innerHTML = '<div class="vgmEmpty">No fue posible validar los accesos directos a las carpetas.</div>';
    } finally {
      this.folderScanStarted = false;
    }
  }

  private refreshClientsButton(): void {
    const button: HTMLButtonElement | null = this.domElement.querySelector<HTMLButtonElement>('[data-action="clients"]');
    if (button) button.textContent = `☰ Mis clientes${this.accessibleClients.length ? ` (${this.accessibleClients.length})` : ''}`;
  }

  private refreshClientModal(): void {
    const filters: HTMLElement | null = this.domElement.querySelector<HTMLElement>('.vgmAreaFilters');
    const rows: HTMLElement | null = this.domElement.querySelector<HTMLElement>('[data-client-list]');
    if (filters) filters.innerHTML = this.renderAreaFilterButtons();
    if (rows) rows.innerHTML = this.renderClientRows();
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

  private async loadMicrosoft365Birthdays(): Promise<BirthdayItem[]> {
    try {
      const client: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');
      const response: { value?: Array<{ displayName?: string; mail?: string; userPrincipalName?: string; department?: string; birthday?: string }> } = await client
        .api('/users')
        .select('displayName,mail,userPrincipalName,department,birthday')
        .top(999)
        .get() as { value?: Array<{ displayName?: string; mail?: string; userPrincipalName?: string; department?: string; birthday?: string }> };
      return (response.value || []).map((user) => ({
        name: user.displayName || '',
        email: user.mail || user.userPrincipalName || '',
        area: user.department || '',
        date: user.birthday ? new Date(user.birthday) : undefined
      })).filter((item: BirthdayItem) => Boolean(item.name && item.date && !Number.isNaN(item.date.getTime())));
    } catch (error) {
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
    return (['LEGAL','TAX','OUTSOURCING','AUDITORIA'] as VgmArea[])
      .filter((area: VgmArea) => item.areas[area] > 0)
      .map((area: VgmArea) => `${this.areaLabel(area)} ${item.areas[area]}`)
      .join(' · ');
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
