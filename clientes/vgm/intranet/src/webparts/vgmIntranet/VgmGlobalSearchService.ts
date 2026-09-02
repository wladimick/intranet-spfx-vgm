import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { VgmArea, VgmClientAccess, VgmRecentDocument } from './VgmSearchService';

export default class VgmGlobalSearchService {
  private static readonly sourceWebUrl: string = 'https://vgmconsultants.sharepoint.com/sites/intranet';

  public constructor(private readonly spHttpClient: SPHttpClient) {}

  public async searchDocuments(queryText: string, access: VgmClientAccess[], limit: number = 30): Promise<VgmRecentDocument[]> {
    const q: string = (queryText || '').trim();
    if (q.length < 2) return [];
    const escaped: string = q.replace(/"/g,'').replace(/'/g,"''");
    const kql: string = `Path:"https://vgmconsultants.sharepoint.com/sites/" IsDocument:1 (${escaped})`;
    const select: string[] = ['Title','Path','SPSiteUrl','SiteTitle','LastModifiedTime','Author','FileType'];
    const params: string[] = [
      `querytext=${encodeURIComponent(`'${kql}'`)}`,
      `selectproperties=${encodeURIComponent(`'${select.join(',')}'`)}`,
      `rowlimit=${Math.max(limit * 8,100)}`,
      `sortlist=${encodeURIComponent("'LastModifiedTime:descending'")}`,
      'trimduplicates=false'
    ];
    const response: SPHttpClientResponse = await this.spHttpClient.get(
      `${VgmGlobalSearchService.sourceWebUrl}/_api/search/query?${params.join('&')}`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: 'application/json;odata=nometadata' } }
    );
    if (!response.ok) throw new Error(`SharePoint Search ${response.status}`);
    const payload: any = await response.json();
    const rows: any[] = payload?.PrimaryQueryResult?.RelevantResults?.Table?.Rows || [];
    const allowed: Map<string,VgmClientAccess> = new Map<string,VgmClientAccess>();
    for (const client of access) allowed.set(client.code,client);
    const results: VgmRecentDocument[] = [];
    for (const row of rows) {
      const record: Record<string,string> = {};
      for (const cell of row.Cells || []) record[String(cell.Key || '')] = String(cell.Value || '');
      const siteUrl: string = record.SPSiteUrl || this.siteFromPath(record.Path || '');
      const code: string = this.clientCode(siteUrl);
      const area: VgmArea | undefined = this.inferArea(record.Path || '');
      const client: VgmClientAccess | undefined = allowed.get(code);
      if (!client || !area || !client.areas[area]) continue;
      const modified: Date | undefined = record.LastModifiedTime ? new Date(record.LastModifiedTime) : undefined;
      results.push({
        title: record.Title || this.fileName(record.Path || ''),
        path: record.Path || '',
        siteUrl,
        siteTitle: client.name || record.SiteTitle || '',
        modified: modified && !Number.isNaN(modified.getTime()) ? modified : undefined,
        author: record.Author || '',
        fileType: (record.FileType || '').toLowerCase(),
        area,
        clientCode: code
      });
      if (results.length >= limit) break;
    }
    return results;
  }

  private inferArea(value: string): VgmArea | undefined {
    let clean: string = value;
    try { clean = decodeURIComponent(value); } catch { /* keep */ }
    clean = clean.normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/%20/g,' ');
    if (/(^|\/|\s)legal($|\/|\s|\?)/.test(clean)) return 'LEGAL';
    if (/(^|\/|\s)tax($|\/|\s|\?)/.test(clean)) return 'TAX';
    if (/(^|\/|\s)outsourcing($|\/|\s|\?)/.test(clean)) return 'OUTSOURCING';
    if (/(^|\/|\s)auditoria($|\/|\s|\?)/.test(clean)) return 'AUDITORIA';
    return undefined;
  }

  private siteFromPath(path: string): string {
    const match: RegExpMatchArray | null = path.match(/^(https:\/\/[^/]+\/sites\/[^/]+)/i);
    return match ? match[1] : '';
  }

  private clientCode(siteUrl: string): string {
    const match: RegExpMatchArray | null = siteUrl.match(/\/sites\/(\d+)-/i);
    return match ? match[1] : '';
  }

  private fileName(path: string): string {
    try { return decodeURIComponent(path.split('/').pop() || 'Documento'); } catch { return path.split('/').pop() || 'Documento'; }
  }
}
