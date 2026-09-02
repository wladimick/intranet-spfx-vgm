import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export type LegacyItem = Record<string, unknown>;
export type MenuItem = { title: string; url: string; parent?: string; kind: 'menu' | 'submenu'; order: number };
export type BirthdayItem = { name: string; email: string; area: string; date?: Date };
export type EventItem = { title: string; start?: Date };
export type NewsItem = { id: number; title: string; summary: string; body: string; image: string };
export type LinkItem = { title: string; url: string };
export type GalleryItem = { title: string; images: string[] };
export type ClientItem = { raw: LegacyItem; code: string; name: string; partner: string };
export type IndicatorItem = { key: string; label: string; value: string };
export type FinancialNewsItem = { title: string; summary: string; category: string; link: string; image: string };

export default class VgmDataService {
  public static readonly sourceWebUrl: string = 'https://vgmconsultants.sharepoint.com/sites/intranet';
  private readonly issues: string[] = [];

  public constructor(private readonly spHttpClient: SPHttpClient) {}

  public getIssues(): string[] {
    return this.issues.slice();
  }

  public async getMenu(): Promise<MenuItem[]> {
    const items: LegacyItem[] = await this.getListItems('Menú Principal', '$orderby=Ubicacion asc&$top=100');
    const menu: MenuItem[] = items.map((item: LegacyItem): MenuItem => ({
      title: this.text(item.Title),
      url: this.urlValue(item['Hiperv_x00ed_nculo']),
      parent: this.text(item.Superior) || undefined,
      kind: this.text(item.Menu).toLowerCase() === 'submenu' ? 'submenu' : 'menu',
      order: Number(item.Ubicacion || 0)
    }));
    return menu.filter((item: MenuItem) => Boolean(item.title));
  }

  public async getBirthdays(): Promise<BirthdayItem[]> {
    // La portada histórica consume la lista Contactos. No ordenamos por un campo interno
    // específico para evitar que un nombre interno distinto deje el bloque completo vacío.
    const items: LegacyItem[] = await this.getListItems('Contactos', '$top=1000');
    return items.map((item: LegacyItem): BirthdayItem => ({
      name: this.firstText(item, ['Title', 'Nombre', 'NombreCompleto']),
      email: this.firstText(item, ['field_1', 'Email', 'EMail', 'Correo', 'CorreoElectronico']),
      area: this.firstText(item, ['field_2', 'Area', 'Área', 'Departamento', 'Unidad']),
      date: this.date(this.firstValue(item, [
        'field_3',
        'Cumpleanos',
        'Cumple_x00f1_os',
        'FechaCumpleanos',
        'Fecha_x0020_Cumpleanos',
        'Fecha_x0020_de_x0020_cumpleanos',
        'FechaNacimiento',
        'Birthday'
      ]))
    })).filter((item: BirthdayItem) => Boolean(item.name && item.date));
  }

  public async getEvents(): Promise<EventItem[]> {
    // Misma fuente/consulta usada por la intranet histórica.
    const now: string = new Date().toISOString();
    const items: LegacyItem[] = await this.getListItems('Calendario', `$filter=Start ge datetime'${now}'&$orderby=Start asc&$top=10`);
    return items.map((item: LegacyItem) => ({
      title: this.text(item.Title),
      start: this.date(item.Start)
    })).filter((item: EventItem) => Boolean(item.title && item.start));
  }

  public async getInternalNews(): Promise<NewsItem[]> {
    const items: LegacyItem[] = await this.getListItems('Noticias', '$top=8&$orderby=Modified desc');
    return items.map((item: LegacyItem) => {
      const id: number = Number(item.Id || item.ID || 0);
      const bodyHtml: string = this.text(item.Contenido || item.Body);
      return {
        id,
        title: this.text(item.Title),
        summary: this.stripHtml(this.text(item.Body || item.Contenido)).slice(0, 180),
        body: bodyHtml,
        image: this.newsImage(item, id)
      };
    }).filter((item: NewsItem) => Boolean(item.title));
  }

  public async getLinks(): Promise<LinkItem[]> {
    const items: LegacyItem[] = await this.getListItems('Links de Interés', '$top=20&$orderby=Ubicacion asc');
    return items.map((item: LegacyItem) => ({ title: this.text(item.Title), url: this.urlValue(item.Link) }))
      .filter((item: LinkItem) => Boolean(item.title && item.url));
  }

  public async getGallery(): Promise<GalleryItem | undefined> {
    try {
      const foldersPayload: { value?: LegacyItem[] } = await this.getJson(`${VgmDataService.sourceWebUrl}/_api/web/lists/getbytitle('Galería')/rootfolder/folders?$select=Name,ServerRelativeUrl,TimeCreated&$orderby=TimeCreated desc`);
      const folders: LegacyItem[] = (foldersPayload.value || []).filter((folder: LegacyItem) => !this.text(folder.Name).startsWith('Forms'));
      const folder: LegacyItem | undefined = folders[0];
      if (!folder) return undefined;
      const relative: string = this.text(folder.ServerRelativeUrl);
      const encoded: string = relative.replace(/'/g, "''");
      const filesPayload: { value?: LegacyItem[] } = await this.getJson(`${VgmDataService.sourceWebUrl}/_api/web/GetFolderByServerRelativeUrl('${encoded}')/Files?$select=Name,ServerRelativeUrl`);
      const images: string[] = (filesPayload.value || [])
        .filter((file: LegacyItem) => /\.(jpg|jpeg|png|gif|webp)$/i.test(this.text(file.Name)))
        .map((file: LegacyItem) => this.text(file.ServerRelativeUrl));
      return { title: this.text(folder.Name), images };
    } catch (error) {
      this.recordIssue('Galería', error);
      return undefined;
    }
  }

  public async getClients(): Promise<ClientItem[]> {
    const items: LegacyItem[] = await this.getListItems('Todos los Clientes - Repositorio Principal', '$top=8000');
    return items.map((raw: LegacyItem) => ({
      raw,
      code: this.clientField(raw, ['Codigo','codigo','CODIGO','C_x00f3_digo','Código'], 'codigo'),
      name: this.clientField(raw, ['Title','Nombre'], 'title'),
      partner: this.text(raw.Socio || raw.socio || raw['Regi_x00f3_n'] || raw.Region)
    })).filter((item: ClientItem) => Boolean(item.code && item.name));
  }

  public async getIndicators(): Promise<IndicatorItem[]> {
    try {
      const response: Response = await fetch('https://mindicador.cl/api');
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      const data: Record<string, unknown> = await response.json() as Record<string, unknown>;
      const defs: Array<[string,string,string]> = [
        ['uf','UF','$'], ['utm','UTM','$'], ['dolar','Dólar','$'], ['euro','Euro','$'],
        ['imacec','Imacec','%'], ['ipc','IPC','%'], ['tpm','TPM','%'], ['libra_cobre','Cobre','USD/lb']
      ];
      return defs.map(([key,label,suffix]: [string,string,string]) => {
        const obj: Record<string, unknown> = (data[key] || {}) as Record<string, unknown>;
        const value: unknown = obj.valor;
        const formatted: string = typeof value === 'number'
          ? (suffix === '%' ? `${value.toLocaleString('es-CL')} %` : suffix === 'USD/lb' ? `${value.toLocaleString('es-CL')} USD/lb` : `$${Math.round(value).toLocaleString('es-CL')}`)
          : '–';
        return { key, label, value: formatted };
      });
    } catch (error) {
      this.recordIssue('Indicadores', error);
      return [];
    }
  }

  public async getFinancialNews(): Promise<FinancialNewsItem[]> {
    try {
      // Replica la llamada que usa funciones-intranet.js de la portada histórica.
      const response: Response = await fetch('https://tiboxrssreader.azurewebsites.net/api/rss/df', {
        method: 'POST',
        headers: { Accept: 'application/json;odata=verbose' },
        cache: 'no-store'
      });
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      const raw: string = await response.text();
      const parsed: unknown = JSON.parse(raw);
      if (!Array.isArray(parsed)) throw new Error('La respuesta no contiene un arreglo de noticias');
      return parsed.slice(0, 4).map((item: LegacyItem) => ({
        title: this.text(item.title),
        summary: this.stripHtml(this.text(item.description)).slice(0, 140),
        category: this.text(item.category),
        link: this.text(item.link),
        image: this.text(item.enclosure)
      })).filter((item: FinancialNewsItem) => Boolean(item.title && item.link));
    } catch (error) {
      this.recordIssue('Diario Financiero', error);
      return [];
    }
  }

  public buildClientSiteBase(client: ClientItem): string {
    const suffix: string = this.clientSuffix(client.code);
    return `https://vgmconsultants.sharepoint.com/sites/${client.code}-${this.slug(client.name)}${suffix}`;
  }

  public buildClientFolderUrl(client: ClientItem, folderName: string): string {
    const suffix: string = this.clientSuffix(client.code);
    const base: string = this.buildClientSiteBase(client);
    const path: string = `/sites/${client.code}-${this.slug(client.name)}${suffix}/Documentos compartidos/${folderName}`;
    return `${base}/Documentos%20compartidos/Forms/AllItems.aspx?RootFolder=${encodeURIComponent(path)}`;
  }

  private clientSuffix(code: string): string {
    const codeNumber: number = Number.parseInt(code, 10);
    return Number.isFinite(codeNumber) && codeNumber >= 105 && codeNumber <= 193 ? '2' : '';
  }

  private async getListItems(title: string, query: string): Promise<LegacyItem[]> {
    try {
      const escaped: string = title.replace(/'/g, "''");
      const payload: { value?: LegacyItem[] } = await this.getJson(`${VgmDataService.sourceWebUrl}/_api/web/lists/getbytitle('${escaped}')/items?${query}`);
      return payload.value || [];
    } catch (error) {
      this.recordIssue(`Lista ${title}`, error);
      return [];
    }
  }

  private async getJson(url: string): Promise<{ value?: LegacyItem[] }> {
    const response: SPHttpClientResponse = await this.spHttpClient.get(url, SPHttpClient.configurations.v1, {
      headers: { Accept: 'application/json;odata=nometadata' }
    });
    if (!response.ok) throw new Error(`SharePoint ${response.status}: ${url}`);
    return response.json() as Promise<{ value?: LegacyItem[] }>;
  }

  private recordIssue(source: string, error: unknown): void {
    const detail: string = error instanceof Error ? error.message : String(error || 'Error desconocido');
    const message: string = `${source}: ${detail}`;
    this.issues.push(message);
    console.warn(`VGM Intranet - ${message}`);
  }

  private firstValue(item: LegacyItem, candidates: string[]): unknown {
    for (const candidate of candidates) {
      const value: unknown = item[candidate];
      if (value !== undefined && value !== null && value !== '') return value;
    }
    const normalizedCandidates: string[] = candidates.map((candidate: string) => this.simplify(candidate));
    const key: string | undefined = Object.keys(item).find((itemKey: string) => normalizedCandidates.indexOf(this.simplify(itemKey)) !== -1);
    return key ? item[key] : undefined;
  }

  private firstText(item: LegacyItem, candidates: string[]): string {
    return this.text(this.firstValue(item, candidates));
  }

  private newsImage(item: LegacyItem, id: number): string {
    const raw: string = this.text(item.Imagen);
    if (raw) {
      try {
        const data: Record<string, unknown> = JSON.parse(raw) as Record<string, unknown>;
        const fileName: string = this.text(data.fileName);
        if (fileName && id) return `${VgmDataService.sourceWebUrl}/Lists/Noticias/Attachments/${id}/${encodeURIComponent(fileName)}`;
      } catch { /* fallback */ }
    }
    return `${VgmDataService.sourceWebUrl}/SiteAssets/img/imagen-no-disponible.jpg`;
  }

  private clientField(item: LegacyItem, candidates: string[], keyword: string): string {
    for (const candidate of candidates) {
      if (item[candidate] !== undefined && item[candidate] !== null && item[candidate] !== '') return this.text(item[candidate]);
    }
    const hit: string | undefined = Object.keys(item).find((key: string) => this.simplify(key).indexOf(keyword) !== -1);
    return hit ? this.text(item[hit]) : '';
  }

  private simplify(value: string): string {
    return value.toLowerCase().replace(/_x[0-9a-f]{4}_/gi, '').normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  }

  private slug(value: string): string {
    return value.replace(/[ÑñÁÉÍÓÚáéíóú]/g, '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^A-Za-z0-9]/g, '').toUpperCase();
  }

  private stripHtml(value: string): string {
    const div: HTMLDivElement = document.createElement('div');
    div.innerHTML = value;
    return (div.textContent || div.innerText || '').trim();
  }

  private text(value: unknown): string {
    if (value === null || value === undefined) return '';
    if (typeof value === 'string') return value;
    if (typeof value === 'number' || typeof value === 'boolean') return String(value);
    return '';
  }

  private date(value: unknown): Date | undefined {
    const text: string = this.text(value);
    if (!text) return undefined;
    const date: Date = new Date(text);
    return Number.isNaN(date.getTime()) ? undefined : date;
  }

  private urlValue(value: unknown): string {
    if (typeof value === 'string') return value;
    if (value && typeof value === 'object') {
      const record: Record<string, unknown> = value as Record<string, unknown>;
      return this.text(record.Url || record.url);
    }
    return '';
  }
}
