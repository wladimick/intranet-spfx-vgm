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
  total: number;
  lastModified?: Date;
  areas: Record<VgmArea, number>;
};

export type VgmAreaActivity = Record<VgmArea, number>;

type SearchRow = Record<string, string>;

export default class VgmSearchService {
  public static readonly sourceWebUrl: string = 'https://vgmconsultants.sharepoint.com/sites/intranet';
  private hubId?: string;

  public constructor(private readonly spHttpClient: SPHttpClient) {}

  public async getAccessibleClients(): Promise<VgmClientAccess[]> {
    const hubId: string = await this.getHubId();
    const scoped: string = this.hubScope(hubId);
    const kql: string = `${scoped} contentclass:STS_ListItem_Folder (Title:Legal OR Title:Tax OR Title:Outsourcing OR Title:Auditoria OR Title:Auditoría)`;
    const rows: SearchRow[] = await this.searchAll(kql, ['Title','Path','SPSiteUrl','SiteTitle'], undefined, 2000);
    const map: Map<string,VgmClientAccess> = new Map<string,VgmClientAccess>();

    for (const row of rows) {
      const area: VgmArea | undefined = this.inferArea(row.Title || row.Path || '');
      if (!area) continue;
      const siteUrl: string = row.SPSiteUrl || this.siteFromPath(row.Path || '');
      const code: string = this.clientCode(siteUrl);
      if (!siteUrl || !code) continue;
      const current: VgmClientAccess = map.get(siteUrl) || {
        code,
        name: this.clientName(row.SiteTitle || '', siteUrl),
        siteUrl,
        areas: {}
      };
      current.areas[area] = row.Path || this.areaFallbackUrl(siteUrl, area);
      map.set(siteUrl,current);
    }

    return Array.from(map.values()).sort((a: VgmClientAccess,b: VgmClientAccess) => Number(a.code) - Number(b.code));
  }

  public async getRecentDocuments(limit: number = 8): Promise<VgmRecentDocument[]> {
    const hubId: string = await this.getHubId();
    const rows: SearchRow[] = await this.searchAll(
      `${this.hubScope(hubId)} IsDocument:1`,
      ['Title','Path','SPSiteUrl','SiteTitle','LastModifiedTime','Author','FileType'],
      'LastModifiedTime:descending',
      Math.max(limit, 20)
    );
    return rows.map((row: SearchRow) => this.toDocument(row)).filter((item: VgmRecentDocument) => Boolean(item.path)).slice(0,limit);
  }

  public async getActivity(days: number = 7, maxRows: number = 1200): Promise<{ clients: VgmClientActivity[]; areas: VgmAreaActivity; documents: VgmRecentDocument[] }> {
    const hubId: string = await this.getHubId();
    const since: Date = new Date();
    since.setDate(since.getDate() - days);
    const dateValue: string = since.toISOString();
    const rows: SearchRow[] = await this.searchAll(
      `${this.hubScope(hubId)} IsDocument:1 LastModifiedTime>=${dateValue}`,
      ['Title','Path','SPSiteUrl','SiteTitle','LastModifiedTime','Author','FileType'],
      'LastModifiedTime:descending',
      maxRows
    );
    const documents: VgmRecentDocument[] = rows.map((row: SearchRow) => this.toDocument(row)).filter((item: VgmRecentDocument) => Boolean(item.path));
    const areas: VgmAreaActivity = { LEGAL: 0, TAX: 0, OUTSOURCING: 0, AUDITORIA: 0 };
    const clientMap: Map<string,VgmClientActivity> = new Map<string,VgmClientActivity>();

    for (const doc of documents) {
      if (doc.area) areas[doc.area] += 1;
      if (!doc.siteUrl || !doc.clientCode) continue;
      const current: VgmClientActivity = clientMap.get(doc.siteUrl) || {
        code: doc.clientCode,
        name: this.clientName(doc.siteTitle,doc.siteUrl),
        siteUrl: doc.siteUrl,
        total: 0,
        areas: { LEGAL: 0, TAX: 0, OUTSOURCING: 0, AUDITORIA: 0 }
      };
      current.total += 1;
      if (doc.area) current.areas[doc.area] += 1;
      if (!current.lastModified || (doc.modified && doc.modified > current.lastModified)) current.lastModified = doc.modified;
      clientMap.set(doc.siteUrl,current);
    }

    const clients: VgmClientActivity[] = Array.from(clientMap.values())
      .sort((a: VgmClientActivity,b: VgmClientActivity) => b.total - a.total || Number(a.code) - Number(b.code));
    return { clients, areas, documents };
  }

  private async getHubId(): Promise<string> {
    if (this.hubId) return this.hubId;
    try {
      const response: SPHttpClientResponse = await this.spHttpClient.get(
        `${VgmSearchService.sourceWebUrl}/_api/site?$select=Id`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: 'application/json;odata=nometadata' } }
      );
      if (!response.ok) throw new Error(`SharePoint ${response.status}`);
      const payload: { Id?: string } = await response.json() as { Id?: string };
      this.hubId = payload.Id || '';
    } catch (error) {
      console.warn('VGM Search: no se pudo obtener el ID del Hub; se usará scope por URL.',error);
      this.hubId = '';
    }
    return this.hubId;
  }

  private hubScope(hubId: string): string {
    return hubId ? `DepartmentId:${hubId}` : 'Path:https://vgmconsultants.sharepoint.com/sites/';
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
    const clean: string = this.normalize(value);
    if (/(^|\/|\s)legal($|\/|\s)/.test(clean)) return 'LEGAL';
    if (/(^|\/|\s)tax($|\/|\s)/.test(clean)) return 'TAX';
    if (/(^|\/|\s)outsourcing($|\/|\s)/.test(clean)) return 'OUTSOURCING';
    if (/(^|\/|\s)auditoria($|\/|\s)/.test(clean)) return 'AUDITORIA';
    return undefined;
  }

  private areaFallbackUrl(siteUrl: string, area: VgmArea): string {
    const label: string = area === 'AUDITORIA' ? 'Auditoria' : area === 'OUTSOURCING' ? 'Outsourcing' : area === 'LEGAL' ? 'Legal' : 'Tax';
    return `${siteUrl}/Documentos%20compartidos/${encodeURIComponent(label)}`;
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
