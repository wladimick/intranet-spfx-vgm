import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export type VgmArea = 'LEGAL' | 'TAX' | 'OUTSOURCING' | 'AUDITORIA';

export type VgmRecentDocument = {
  title: string;
  path: string;
  siteUrl: string;
  siteTitle: string;
  modified?: Date;
  author: string;
  fileType: string;
  area?: VgmArea;
  clientCode: string;
};

export type VgmClientAccess = {
  code: string;
  name: string;
  siteUrl: string;
  areas: Partial<Record<VgmArea, string>>;
};

export type VgmClientActivity = {
  code: string;
  name: string;
  siteUrl: string;
  openUrl: string;
  total: number;
  lastModified?: Date;
  areas: Record<VgmArea, number>;
};

export type VgmAreaActivity = Record<VgmArea, number>;

type SearchRow = Record<string, string>;

export default class VgmSearchService {
  public static readonly sourceWebUrl: string = 'https://vgmconsultants.sharepoint.com/sites/intranet';
  private static readonly tenantSitesScope: string = 'Path:"https://vgmconsultants.sharepoint.com/sites/"';

  public constructor(private readonly spHttpClient: SPHttpClient) {}

  /**
   * Obtiene clientes/áreas que SharePoint Search devuelve para el usuario actual.
   * Se buscan primero las carpetas y luego se complementa con documentos visibles.
   * Esto es importante en VGM porque el usuario puede tener solamente Limited Access
   * al sitio y permiso real sobre una carpeta concreta de Documentos compartidos.
   */
  public async getAccessibleClients(): Promise<VgmClientAccess[]> {
    const map: Map<string,VgmClientAccess> = new Map<string,VgmClientAccess>();

    const folderRows: SearchRow[] = await this.searchAll(
      `${VgmSearchService.tenantSitesScope} contentclass:STS_ListItem_Folder`,
      ['Title','Path','SPSiteUrl','SiteTitle'],
      undefined,
      5000
    );
    this.mergeAccessRows(map,folderRows);

    // Algunas bibliotecas no exponen la carpeta como resultado de búsqueda,
    // pero sí los documentos que el usuario puede leer dentro de ella.
    const documentRows: SearchRow[] = await this.searchAll(
      `${VgmSearchService.tenantSitesScope} IsDocument:1`,
      ['Title','Path','SPSiteUrl','SiteTitle'],
      'LastModifiedTime:descending',
      5000
    );
    this.mergeAccessRows(map,documentRows);

    return Array.from(map.values())
      .filter((client: VgmClientAccess) => Object.keys(client.areas).length > 0)
      .sort((a: VgmClientAccess,b: VgmClientAccess) => Number(a.code) - Number(b.code));
  }

  public async getRecentDocuments(limit: number = 8): Promise<VgmRecentDocument[]> {
    const rows: SearchRow[] = await this.searchAll(
      `${VgmSearchService.tenantSitesScope} IsDocument:1`,
      ['Title','Path','SPSiteUrl','SiteTitle','LastModifiedTime','Author','FileType'],
      'LastModifiedTime:descending',
      Math.max(limit * 8,80)
    );
    return rows
      .map((row: SearchRow) => this.toDocument(row))
      .filter((item: VgmRecentDocument) => Boolean(item.path && item.clientCode && item.area))
      .slice(0,limit);
  }

  public async getActivity(days: number = 7, maxRows: number = 1200): Promise<{ clients: VgmClientActivity[]; areas: VgmAreaActivity; documents: VgmRecentDocument[] }> {
    const since: Date = new Date();
    since.setDate(since.getDate() - days);
    const dateValue: string = since.toISOString();
    const rows: SearchRow[] = await this.searchAll(
      `${VgmSearchService.tenantSitesScope} IsDocument:1 LastModifiedTime>=${dateValue}`,
      ['Title','Path','SPSiteUrl','SiteTitle','LastModifiedTime','Author','FileType'],
      'LastModifiedTime:descending',
      maxRows
    );
    const documents: VgmRecentDocument[] = rows
      .map((row: SearchRow) => this.toDocument(row))
      .filter((item: VgmRecentDocument) => Boolean(item.path && item.clientCode && item.area));
    const areas: VgmAreaActivity = { LEGAL: 0, TAX: 0, OUTSOURCING: 0, AUDITORIA: 0 };
    const clientMap: Map<string,VgmClientActivity> = new Map<string,VgmClientActivity>();

    for (const doc of documents) {
      if (doc.area) areas[doc.area] += 1;
      if (!doc.siteUrl || !doc.clientCode) continue;
      const current: VgmClientActivity = clientMap.get(doc.siteUrl) || {
        code: doc.clientCode,
        name: this.clientName(doc.siteTitle,doc.siteUrl),
        siteUrl: doc.siteUrl,
        openUrl: doc.area ? this.areaBrowseUrl(doc.siteUrl,doc.area) : doc.path,
        total: 0,
        areas: { LEGAL: 0, TAX: 0, OUTSOURCING: 0, AUDITORIA: 0 }
      };
      current.total += 1;
      if (doc.area) current.areas[doc.area] += 1;
      if (!current.lastModified || (doc.modified && doc.modified > current.lastModified)) {
        current.lastModified = doc.modified;
        current.openUrl = doc.area ? this.areaBrowseUrl(doc.siteUrl,doc.area) : doc.path;
      }
      clientMap.set(doc.siteUrl,current);
    }

    const clients: VgmClientActivity[] = Array.from(clientMap.values())
      .sort((a: VgmClientActivity,b: VgmClientActivity) => b.total - a.total || Number(a.code) - Number(b.code));
    return { clients, areas, documents };
  }

  private mergeAccessRows(map: Map<string,VgmClientAccess>, rows: SearchRow[]): void {
    for (const row of rows) {
      const path: string = row.Path || '';
      const area: VgmArea | undefined = this.inferArea(path || row.Title || '');
      if (!area) continue;
      const siteUrl: string = row.SPSiteUrl || this.siteFromPath(path);
      const code: string = this.clientCode(siteUrl);
      if (!siteUrl || !code) continue;
      const current: VgmClientAccess = map.get(siteUrl) || {
        code,
        name: this.clientName(row.SiteTitle || '',siteUrl),
        siteUrl,
        areas: {}
      };
      // Nunca enviamos al usuario al inicio del sitio: el permiso puede existir solo
      // sobre la carpeta. El enlace siempre apunta directo al área autorizada.
      current.areas[area] = this.areaBrowseUrl(siteUrl,area);
      map.set(siteUrl,current);
    }
  }

  private async searchAll(kql: string, select: string[], sort?: string, maxRows: number = 500): Promise<SearchRow[]> {
    const results: SearchRow[] = [];
    const pageSize: number = Math.min(500,maxRows);
    let startRow: number = 0;
    while (results.length < maxRows) {
      const page: SearchRow[] = await this.search(kql,select,sort,pageSize,startRow);
      results.push(...page);
      if (page.length < pageSize) break;
      startRow += pageSize;
    }
    return results.slice(0,maxRows);
  }

  private async search(kql: string, select: string[], sort: string | undefined, rowLimit: number, startRow: number): Promise<SearchRow[]> {
    const query: string[] = [
      `querytext=${encodeURIComponent(`'${kql.replace(/'/g,"''")}'`)}`,
      `selectproperties=${encodeURIComponent(`'${select.join(',')}'`)}`,
      `rowlimit=${rowLimit}`,
      `startrow=${startRow}`,
      'trimduplicates=false'
    ];
    if (sort) query.push(`sortlist=${encodeURIComponent(`'${sort}'`)}`);
    const url: string = `${VgmSearchService.sourceWebUrl}/_api/search/query?${query.join('&')}`;
    const response: SPHttpClientResponse = await this.spHttpClient.get(url,SPHttpClient.configurations.v1,{ headers: { Accept: 'application/json;odata=nometadata' } });
    if (!response.ok) throw new Error(`SharePoint Search ${response.status}: ${kql}`);
    const payload: any = await response.json();
    const rows: any[] = payload?.PrimaryQueryResult?.RelevantResults?.Table?.Rows || [];
    return rows.map((row: any): SearchRow => {
      const record: SearchRow = {};
      const cells: any[] = row.Cells || [];
      for (const cell of cells) record[String(cell.Key || '')] = String(cell.Value || '');
      return record;
    });
  }

  private toDocument(row: SearchRow): VgmRecentDocument {
    const siteUrl: string = row.SPSiteUrl || this.siteFromPath(row.Path || '');
    const modified: Date | undefined = row.LastModifiedTime ? new Date(row.LastModifiedTime) : undefined;
    return {
      title: row.Title || this.fileName(row.Path || ''),
      path: row.Path || '',
      siteUrl,
      siteTitle: row.SiteTitle || '',
      modified: modified && !Number.isNaN(modified.getTime()) ? modified : undefined,
      author: row.Author || '',
      fileType: (row.FileType || '').toLowerCase(),
      area: this.inferArea(row.Path || ''),
      clientCode: this.clientCode(siteUrl)
    };
  }

  private inferArea(value: string): VgmArea | undefined {
    let clean: string = this.normalize(value);
    try { clean = this.normalize(decodeURIComponent(value)); } catch { /* keep original */ }
    clean = clean.replace(/%20/g,' ');
    if (/(^|\/|\s)legal($|\/|\s|\?)/.test(clean)) return 'LEGAL';
    if (/(^|\/|\s)tax($|\/|\s|\?)/.test(clean)) return 'TAX';
    if (/(^|\/|\s)outsourcing($|\/|\s|\?)/.test(clean)) return 'OUTSOURCING';
    if (/(^|\/|\s)auditoria($|\/|\s|\?)/.test(clean)) return 'AUDITORIA';
    return undefined;
  }

  private areaBrowseUrl(siteUrl: string, area: VgmArea): string {
    const label: string = area === 'AUDITORIA' ? 'Auditoria' : area === 'OUTSOURCING' ? 'Outsourcing' : area === 'LEGAL' ? 'Legal' : 'Tax';
    const sitePath: string = new URL(siteUrl).pathname;
    const folder: string = `${sitePath}/Documentos compartidos/${label}`;
    return `${siteUrl}/Documentos%20compartidos/Forms/AllItems.aspx?RootFolder=${encodeURIComponent(folder)}`;
  }

  private siteFromPath(path: string): string {
    const match: RegExpMatchArray | null = path.match(/^(https:\/\/[^/]+\/sites\/[^/]+)/i);
    return match ? match[1] : '';
  }

  private clientCode(siteUrl: string): string {
    const match: RegExpMatchArray | null = siteUrl.match(/\/sites\/(\d+)-/i);
    return match ? match[1] : '';
  }

  private clientName(siteTitle: string, siteUrl: string): string {
    if (siteTitle && !/^\d+[-\s]/.test(siteTitle)) return siteTitle;
    const segment: string = decodeURIComponent(siteUrl.split('/').pop() || '');
    return segment.replace(/^\d+-/,'').replace(/2$/,'').replace(/([a-z])([A-Z])/g,'$1 $2') || siteTitle || 'Cliente';
  }

  private fileName(path: string): string {
    try { return decodeURIComponent(path.split('/').pop() || 'Documento'); } catch { return path.split('/').pop() || 'Documento'; }
  }

  private normalize(value: string): string {
    return (value || '').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase();
  }
}
