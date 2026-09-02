import { Version } from '@microsoft/sp-core-library';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import VgmDataService, {
  BirthdayItem, ClientItem, EventItem, FinancialNewsItem, GalleryItem, IndicatorItem, LinkItem, MenuItem, NewsItem
} from './VgmDataService';
import { VGM_STYLES } from './VgmIntranetStyles';

export interface IVgmIntranetWebPartProps {}

type FolderKey = 'LEGAL' | 'TAX' | 'OUTSOURCING' | 'AUDITORIA';
type Permissions = { admin: boolean; folders: Record<FolderKey, Set<string>> } | undefined;

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
  private clients: ClientItem[] = [];
  private filteredClients: ClientItem[] = [];
  private clientsLoaded: boolean = false;
  private partnerFilter: string = 'ALL';
  private clientSearch: string = '';
  private permissions: Permissions;
  private data: DashboardData = { menu: [], birthdays: [], events: [], news: [], links: [], indicators: [], financial: [] };

  private readonly groupPermissions: Record<string, { admin?: boolean; partners?: string[] | '*'; folders?: FolderKey[] }> = {
    'SP- Administradores Globales': { admin: true },
    'Equipo Legal - Jaime Rosso': { partners: '*', folders: ['LEGAL'] },
    'Equipo TAX - Pablo Vera': { partners: ['PV'], folders: ['TAX'] },
    'Equipo 2 TAX - Pablo Vera': { partners: ['PV'], folders: ['TAX'] },
    'Equipo TAX - Claudia Gómez': { partners: ['CG'], folders: ['TAX'] },
    'Equipo TAX - Jaime Rosso': { partners: ['JR'], folders: ['TAX'] },
    'Equipo Tax - Álvaro Mecklenburg': { partners: ['AM'], folders: ['TAX'] },
    'Equipo OUT - Pablo Vera': { partners: ['PV'], folders: ['OUTSOURCING'] },
    'Equipo OUT - Claudia Gómez': { partners: ['CG'], folders: ['OUTSOURCING'] },
    'Equipo OUT- Álvaro Mecklenburg': { partners: ['AM'], folders: ['OUTSOURCING'] },
    'Equipo OUT - TODOS': { partners: '*', folders: ['OUTSOURCING'] },
    'Equipo Auditores - César Cavieres': { partners: ['CC'], folders: ['AUDITORIA'] },
    'Equipo Auditores - Jorge Belloni': { partners: ['JB'], folders: ['AUDITORIA'] }
  };

  public render(): void {
    void this.renderAsync();
  }

  private async renderAsync(): Promise<void> {
    this.domElement.innerHTML = `<style>${VGM_STYLES}</style><div class="vgmApp"><div class="vgmLoading">Cargando intranet VGM…</div></div>`;
    this.service = new VgmDataService(this.context.spHttpClient);

    const [menu, birthdays, events, news, links, gallery, indicators, financial, permissions] = await Promise.all([
      this.service.getMenu(),
      this.service.getBirthdays(),
      this.service.getEvents(),
      this.service.getInternalNews(),
      this.service.getLinks(),
      this.service.getGallery(),
      this.service.getIndicators(),
      this.service.getFinancialNews(),
      this.loadPermissions()
    ]);

    this.data = { menu, birthdays, events, news, links, gallery, indicators, financial };
    this.permissions = permissions;
    this.paint();
    this.bindEvents();
  }

  private paint(): void {
    const displayName: string = this.context.pageContext.user.displayName || 'Usuario VGM';
    const email: string = this.context.pageContext.user.email || '';
    const photo: string = `https://vgmconsultants.sharepoint.com/_layouts/15/userphoto.aspx?size=L&accountname=${encodeURIComponent(email)}`;
    const today: Date = new Date();
    const nextBirthdays: BirthdayItem[] = this.upcomingBirthdays(this.data.birthdays).slice(0, 4);
    const featured: NewsItem | undefined = this.data.news[0];
    const secondary: NewsItem[] = this.data.news.slice(1, 5);

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
            <nav class="vgmMenu">${this.renderMenu(this.data.menu)}<button class="vgmClientsBtn" data-action="clients">☰ Clientes</button></nav>
          </header>

          <div class="vgmMain">
            <aside class="vgmSide">
              <section class="vgmCard"><div class="vgmCardBody"><div class="vgmDateBox">▣ ${this.capitalize(today.toLocaleDateString('es-CL',{weekday:'long',day:'numeric',month:'long'}))}</div></div></section>
              ${this.renderLegacyQuickLinks()}
              <a class="vgmTicket" href="https://soporte.tibox.cl/Login/LoginCliente" target="_blank" rel="noopener">◉ Tickets Tibox</a>
            </aside>

            <main class="vgmContent">
              <div class="vgmGrid">
                <div class="vgmCol">
                  <section class="vgmCard">
                    <div class="vgmCardHead"><h2>Diario Financiero</h2></div>
                    <div class="vgmCardBody">${this.renderFinancial(this.data.financial)}</div>
                  </section>
                  <section class="vgmCard">
                    <div class="vgmCardHead"><h2>Noticias VGM</h2></div>
                    ${this.renderNews(featured, secondary)}
                  </section>
                  <section class="vgmCard"><div class="vgmCardHead"><h2>Indicadores Económicos</h2></div><div class="vgmCardBody"><div class="vgmIndicators">${this.renderIndicators(this.data.indicators)}</div></div></section>
                  <div class="vgmBottom">
                    <section class="vgmCard"><div class="vgmCardHead"><h2>Galerías</h2></div><div class="vgmCardBody">${this.renderGallery(this.data.gallery)}</div></section>
                    <section class="vgmCard"><div class="vgmCardHead"><h2>Links de Interés</h2></div><div class="vgmCardBody vgmLinks">${this.renderLinks(this.data.links)}</div></section>
                  </div>
                </div>

                <div class="vgmCol">
                  <section class="vgmCard"><div class="vgmCardHead"><h2>${this.capitalize(today.toLocaleDateString('es-CL',{month:'long'}))} ${today.getFullYear()}</h2></div><div class="vgmCardBody">${this.renderCalendar(today,this.data.events)}${this.renderEvents(this.data.events.slice(0,3))}</div></section>
                  <section class="vgmCard"><div class="vgmCardHead"><h2>Próximos Cumpleaños</h2></div><div class="vgmCardBody">${this.renderBirthdays(nextBirthdays)}</div></section>
                </div>
              </div>
            </main>
          </div>
        </div>

        <div class="vgmModalOverlay" data-modal="clients"><div class="vgmModal"><div class="vgmModalHead"><h2>Listado de Clientes</h2><button class="vgmClose" data-action="close">×</button></div><div class="vgmClientToolbar"><input data-client-search placeholder="Buscar por cliente, código o socio…"><div data-partner-chips></div></div><div class="vgmClientList" data-client-list><div class="vgmLoading">Cargando clientes…</div></div></div></div>
        <div class="vgmModalOverlay" data-modal="news"><div class="vgmModal"><div class="vgmModalHead"><h2 data-news-title>Noticia</h2><button class="vgmClose" data-action="close">×</button></div><div class="vgmCardBody" data-news-content></div></div></div>
      </div>`;
  }

  private renderMenu(items: MenuItem[]): string {
    const parents: MenuItem[] = items.filter((item: MenuItem) => item.kind === 'menu');
    if (!parents.length) return '<a href="#">Facturación</a><a href="#">RRHH</a><a href="#">Fondo Cliente</a><a href="#">Legal</a>';
    return parents.map((parent: MenuItem) => {
      const children: MenuItem[] = items.filter((item: MenuItem) => item.kind === 'submenu' && item.parent === parent.title);
      if (!children.length) return `<a href="${this.esc(parent.url || '#')}" target="_blank" rel="noopener">${this.esc(parent.title)}</a>`;
      return `<details><summary>${this.esc(parent.title)}</summary><div>${children.map((child: MenuItem) => `<a href="${this.esc(child.url)}" target="_blank" rel="noopener">${this.esc(child.title)}</a>`).join('')}</div></details>`;
    }).join('');
  }

  private renderLegacyQuickLinks(): string {
    return `<section class="vgmCard"><div class="vgmQuickList">
      <a href="${VgmDataService.sourceWebUrl}/Lists/Contactos/AllItems.aspx" target="_blank" rel="noopener">◉ Contactos</a>
      <details class="vgmQuickGroup"><summary>◉ Organigramas</summary><a href="${VgmDataService.sourceWebUrl}/SitePages/Organigrama-VGM-Auditores.aspx">VGM Auditores</a><a href="${VgmDataService.sourceWebUrl}/SitePages/Organigrama-VGM-Outsourcing.aspx">VGM Outsourcing</a><a href="${VgmDataService.sourceWebUrl}/SitePages/Organigrama-VGM-Profesionales.aspx">VGM Profesionales</a></details>
      <details class="vgmQuickGroup"><summary>▣ RIOHS</summary><a href="${VgmDataService.sourceWebUrl}/SitePages/RIOHS-VGM-Auditores.aspx">VGM Auditores</a><a href="${VgmDataService.sourceWebUrl}/SitePages/RIOHS-VGM-Outsourcing.aspx">VGM Outsourcing</a><a href="${VgmDataService.sourceWebUrl}/SitePages/RIOHS-VGM-Profesionales.aspx">VGM Profesionales</a></details>
    </div></section>`;
  }

  private renderFinancial(items: FinancialNewsItem[]): string {
    if (!items.length) return '<div class="vgmEmpty">No fue posible cargar Diario Financiero.</div>';
    return `<div class="vgmDfGrid">${items.map((item: FinancialNewsItem) => `<a class="vgmDf" href="${this.esc(item.link)}" target="_blank" rel="noopener"><img src="${this.esc(item.image)}" alt=""><div><strong>${this.esc(item.title)}</strong><small>${this.esc(item.category)}</small><p>${this.esc(item.summary)}</p></div></a>`).join('')}</div>`;
  }

  private renderNews(featured: NewsItem | undefined, secondary: NewsItem[]): string {
    if (!featured) return '<div class="vgmEmpty">No hay noticias publicadas.</div>';
    return `<button class="vgmNewsHero" data-action="news" data-id="${featured.id}" style="border:0;width:100%;text-align:left;background-image:url('${this.cssUrl(featured.image)}')"><span class="vgmNewsOverlay"><span><h3>${this.esc(featured.title)}</h3><p>${this.esc(featured.summary)}</p></span></span></button><div class="vgmCardBody"><div class="vgmNewsSecondary">${secondary.map((item: NewsItem) => `<button class="vgmNewsMini" data-action="news" data-id="${item.id}" style="text-align:left"><strong>${this.esc(item.title)}</strong><p>${this.esc(item.summary)}</p></button>`).join('')}</div></div>`;
  }

  private renderIndicators(items: IndicatorItem[]): string {
    return items.length ? items.map((item: IndicatorItem) => `<div class="vgmIndicator"><span>${this.esc(item.label)}</span><strong>${this.esc(item.value)}</strong></div>`).join('') : '<div class="vgmEmpty">Indicadores no disponibles.</div>';
  }

  private renderGallery(gallery: GalleryItem | undefined): string {
    if (!gallery || !gallery.images.length) return '<div class="vgmEmpty">No hay imágenes disponibles.</div>';
    return `<div class="vgmGallery"><img src="${this.esc(gallery.images[0])}" alt="${this.esc(gallery.title)}"><div class="vgmGalleryTitle">${this.esc(gallery.title)}</div></div>`;
  }

  private renderLinks(items: LinkItem[]): string {
    return items.length ? items.map((item: LinkItem) => `<a href="${this.esc(item.url)}" target="_blank" rel="noopener">• ${this.esc(item.title)}</a>`).join('') : '<div class="vgmEmpty">Sin enlaces.</div>';
  }

  private renderCalendar(date: Date, events: EventItem[]): string {
    const year: number = date.getFullYear(); const month: number = date.getMonth();
    const first: Date = new Date(year, month, 1); const days: number = new Date(year, month + 1, 0).getDate();
    let offset: number = first.getDay() - 1; if (offset < 0) offset = 6;
    const eventDays: Set<number> = new Set(events.filter((event: EventItem) => event.start && event.start.getFullYear() === year && event.start.getMonth() === month).map((event: EventItem) => event.start!.getDate()));
    const cells: string[] = []; for (let i=0;i<offset;i++) cells.push('<span class="vgmDay"></span>');
    for (let day=1;day<=days;day++) cells.push(`<span class="vgmDay ${day===date.getDate()?'today':''} ${eventDays.has(day)?'event':''}">${day}</span>`);
    return `<div class="vgmCalendar">${['L','M','X','J','V','S','D'].map((d:string)=>`<span class="vgmDow">${d}</span>`).join('')}${cells.join('')}</div>`;
  }

  private renderEvents(items: EventItem[]): string {
    if (!items.length) return '<ul class="vgmEvents"><li>No hay eventos próximos.</li></ul>';
    return `<ul class="vgmEvents">${items.map((item: EventItem) => `<li><strong>${this.esc(item.title)}</strong><small>${item.start ? this.capitalize(item.start.toLocaleDateString('es-CL',{weekday:'short',day:'numeric',month:'short'}))+' · '+item.start.toLocaleTimeString('es-CL',{hour:'2-digit',minute:'2-digit'}) : ''}</small></li>`).join('')}</ul>`;
  }

  private renderBirthdays(items: BirthdayItem[]): string {
    if (!items.length) return '<div class="vgmEmpty">No hay cumpleaños próximos.</div>';
    return items.map((item: BirthdayItem) => {
      const photo: string = `https://vgmconsultants.sharepoint.com/_layouts/15/userphoto.aspx?size=M&accountname=${encodeURIComponent(item.email)}`;
      const next: Date = this.nextBirthday(item.date!);
      const message: string = encodeURIComponent(`¡Feliz cumpleaños ${item.name.split(' ')[0]}! Espero que tengas un gran día.`);
      return `<div class="vgmBirthday"><img src="${photo}" alt=""><span><strong>${this.esc(item.name)}</strong><small>${this.esc(item.area)} · ${next.toLocaleDateString('es-CL',{day:'numeric',month:'short'})}</small></span><a href="https://teams.microsoft.com/l/chat/0/0?users=${encodeURIComponent(item.email)}&message=${message}" target="_blank" rel="noopener">Felicitar</a></div>`;
    }).join('');
  }

  private bindEvents(): void {
    this.domElement.addEventListener('click', (event: MouseEvent): void => {
      const target: HTMLElement | null = (event.target as HTMLElement).closest<HTMLElement>('[data-action]');
      if (!target) return;
      const action: string | null = target.getAttribute('data-action');
      if (action === 'clients') void this.openClients();
      if (action === 'close') this.closeModals();
      if (action === 'news') this.openNews(Number(target.getAttribute('data-id')));
    });
    const search: HTMLInputElement | null = this.domElement.querySelector<HTMLInputElement>('[data-client-search]');
    search?.addEventListener('input', (): void => { this.clientSearch = search.value; this.applyClientFilters(); });
  }

  private async openClients(): Promise<void> {
    this.openModal('clients');
    if (!this.clientsLoaded) {
      this.clients = await this.service.getClients();
      this.clientsLoaded = true;
      this.applyClientFilters();
    }
  }

  private applyClientFilters(): void {
    const q: string = this.normalize(this.clientSearch);
    this.filteredClients = this.clients.filter((client: ClientItem) => this.hasAnyClientAccess(client.partner))
      .filter((client: ClientItem) => this.partnerFilter === 'ALL' || this.parsePartners(client.partner).includes(this.partnerFilter))
      .filter((client: ClientItem) => !q || [client.code,client.name,client.partner].some((value:string)=>this.normalize(value).includes(q)))
      .sort((a:ClientItem,b:ClientItem)=>Number(a.code)-Number(b.code));
    this.renderClientChips(); this.renderClientRows();
  }

  private renderClientChips(): void {
    const holder: HTMLElement | null = this.domElement.querySelector<HTMLElement>('[data-partner-chips]'); if (!holder) return;
    const partners: string[] = ['ALL','AM','CG','PV','CC','JB','EH','JR'];
    holder.innerHTML = partners.map((partner:string)=>`<button class="vgmChip ${this.partnerFilter===partner?'active':''}" data-partner="${partner}">${partner==='ALL'?'Todos':partner}</button>`).join('');
    holder.querySelectorAll<HTMLElement>('[data-partner]').forEach((button:HTMLElement)=>button.addEventListener('click',():void=>{this.partnerFilter=button.getAttribute('data-partner')||'ALL';this.applyClientFilters();}));
  }

  private renderClientRows(): void {
    const holder: HTMLElement | null = this.domElement.querySelector<HTMLElement>('[data-client-list]'); if (!holder) return;
    if (!this.filteredClients.length) { holder.innerHTML='<div class="vgmEmpty">No se encontraron clientes con los filtros actuales.</div>'; return; }
    holder.innerHTML = this.filteredClients.slice(0,800).map((client:ClientItem)=>`<div class="vgmClientRow"><strong>${this.esc(client.code)}</strong><span>${this.esc(client.name)}</span><span>${this.esc(client.partner)}</span><span class="vgmFolderLinks">${this.folderLink(client,'LEGAL','Legal')}${this.folderLink(client,'TAX','Tax')}${this.folderLink(client,'OUTSOURCING','Outsourcing')}${this.folderLink(client,'AUDITORIA','Auditoría')}</span></div>`).join('');
  }

  private folderLink(client: ClientItem, folder: FolderKey, label: string): string {
    const allowed: boolean = this.hasFolderAccess(client.partner, folder);
    if (!allowed) return `<span class="disabled">${label}</span>`;
    return `<a href="${this.esc(this.service.buildClientFolderUrl(client,label === 'Auditoría' ? 'Auditoria' : label))}" target="_blank" rel="noopener">${label}</a>`;
  }

  private openNews(id: number): void {
    const item: NewsItem | undefined = this.data.news.find((news:NewsItem)=>news.id===id); if (!item) return;
    const title: HTMLElement | null = this.domElement.querySelector<HTMLElement>('[data-news-title]');
    const content: HTMLElement | null = this.domElement.querySelector<HTMLElement>('[data-news-content]');
    if (title) title.textContent=item.title;
    if (content) content.innerHTML=`<img src="${this.esc(item.image)}" alt="" style="width:100%;max-height:280px;object-fit:cover;border-radius:14px;margin-bottom:16px"><div>${item.body}</div>`;
    this.openModal('news');
  }

  private openModal(name: string): void { this.domElement.querySelector<HTMLElement>(`[data-modal="${name}"]`)?.classList.add('open'); }
  private closeModals(): void { this.domElement.querySelectorAll<HTMLElement>('.vgmModalOverlay').forEach((element:HTMLElement)=>element.classList.remove('open')); }

  private async loadPermissions(): Promise<Permissions> {
    try {
      const client: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');
      const response: { value?: Array<{ displayName?: string }> } = await client.api('/me/memberOf').select('displayName').top(999).get() as { value?: Array<{ displayName?: string }> };
      const names: string[] = (response.value || []).map((group:{displayName?:string})=>group.displayName || '').filter(Boolean);
      if (!names.length) return undefined;
      const result: Permissions = { admin:false, folders:{ LEGAL:new Set<string>(), TAX:new Set<string>(), OUTSOURCING:new Set<string>(), AUDITORIA:new Set<string>() } };
      for (const name of names) {
        const definition = this.groupPermissions[name]; if (!definition) continue;
        if (definition.admin) { result.admin=true; return result; }
        for (const folder of definition.folders || []) {
          if (definition.partners === '*') result.folders[folder].add('*');
          else for (const partner of definition.partners || []) result.folders[folder].add(partner);
        }
      }
      return result;
    } catch (error) {
      console.warn('VGM: no fue posible resolver grupos por Graph; se mantiene fallback compatible con intranet anterior.', error);
      return undefined;
    }
  }

  private hasFolderAccess(partnerText: string, folder: FolderKey): boolean {
    if (!this.permissions || this.permissions.admin) return true;
    const allowed: Set<string> = this.permissions.folders[folder];
    return this.parsePartners(partnerText).some((partner:string)=>allowed.has('*')||allowed.has(partner));
  }

  private hasAnyClientAccess(partnerText: string): boolean {
    return (['LEGAL','TAX','OUTSOURCING','AUDITORIA'] as FolderKey[]).some((folder:FolderKey)=>this.hasFolderAccess(partnerText,folder));
  }

  private parsePartners(value: string): string[] { return this.normalize(value).toUpperCase().split(/[^A-Z]+/).filter(Boolean); }
  private normalize(value: string): string { return (value||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase(); }
  private upcomingBirthdays(items: BirthdayItem[]): BirthdayItem[] { return items.slice().sort((a:BirthdayItem,b:BirthdayItem)=>this.nextBirthday(a.date!).getTime()-this.nextBirthday(b.date!).getTime()); }
  private nextBirthday(date: Date): Date { const now=new Date(); let next=new Date(now.getFullYear(),date.getMonth(),date.getDate()); if(next.getTime()<new Date(now.getFullYear(),now.getMonth(),now.getDate()).getTime()) next=new Date(now.getFullYear()+1,date.getMonth(),date.getDate()); return next; }
  private capitalize(value: string): string { return value ? value.charAt(0).toUpperCase()+value.slice(1) : value; }
  private esc(value: string): string { return String(value||'').replace(/[&<>"']/g,(char:string)=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[char]||char)); }
  private cssUrl(value: string): string { return this.esc(value).replace(/'/g,'%27'); }

  protected get dataVersion(): Version { return Version.parse('1.0'); }
}
