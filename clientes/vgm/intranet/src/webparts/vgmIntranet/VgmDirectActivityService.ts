import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { VgmArea, VgmAreaActivity, VgmClientAccess, VgmClientActivity, VgmRecentDocument } from './VgmSearchService';

export default class VgmDirectActivityService {
  public constructor(private readonly spHttpClient: SPHttpClient) {}

  public async getActivity(
    clients: VgmClientAccess[],
    days: number = 7,
    recentLimit: number = 8,
    concurrency: number = 8
  ): Promise<{ clients: VgmClientActivity[]; areas: VgmAreaActivity; documents: VgmRecentDocument[] }> {
    const documents: VgmRecentDocument[] = [];
    let cursor: number = 0;
    const since: number = Date.now() - (days * 24 * 60 * 60 * 1000);

    const worker = async (): Promise<void> => {
      while (true) {
        const index: number = cursor++;
        if (index >= clients.length) return;
        const client: VgmClientAccess = clients[index];
        const docs: VgmRecentDocument[] = await this.readClient(client,since);
        documents.push(...docs);
      }
    };

    const workers: Promise<void>[] = [];
    const count: number = Math.max(1,Math.min(concurrency,16));
    for (let i: number = 0; i < count; i++) workers.push(worker());
    await Promise.all(workers);

    documents.sort((a: VgmRecentDocument,b: VgmRecentDocument) => (b.modified?.getTime() || 0) - (a.modified?.getTime() || 0));
    const areas: VgmAreaActivity = { LEGAL: 0, TAX: 0, OUTSOURCING: 0, AUDITORIA: 0 };
    const map: Map<string,VgmClientActivity> = new Map<string,VgmClientActivity>();

    for (const doc of documents) {
      if (!doc.area) continue;
      areas[doc.area] += 1;
      const access: VgmClientAccess | undefined = clients.find((item: VgmClientAccess) => item.code === doc.clientCode);
      if (!access) continue;
      const current: VgmClientActivity = map.get(access.code) || {
        code: access.code,
        name: access.name,
        siteUrl: access.siteUrl,
        openUrl: access.areas[doc.area] || doc.path,
        total: 0,
        areas: { LEGAL: 0, TAX: 0, OUTSOURCING: 0, AUDITORIA: 0 }
      };
      current.total += 1;
      current.areas[doc.area] += 1;
      if (!current.lastModified || (doc.modified && doc.modified > current.lastModified)) {
        current.lastModified = doc.modified;
        current.openUrl = access.areas[doc.area] || doc.path;
      }
      map.set(access.code,current);
    }

    return {
      clients: Array.from(map.values()).sort((a: VgmClientActivity,b: VgmClientActivity) => b.total - a.total),
      areas,
      documents: documents.slice(0,recentLimit)
    };
  }

  private async readClient(client: VgmClientAccess, since: number): Promise<VgmRecentDocument[]> {
    const docs: VgmRecentDocument[] = [];
    const areas: VgmArea[] = ['LEGAL','TAX','OUTSOURCING','AUDITORIA'];
    await Promise.all(areas.map(async (area: VgmArea): Promise<void> => {
      if (!client.areas[area]) return;
      try {
        const sitePath: string = new URL(client.siteUrl).pathname;
        const folderName: string = this.folderName(area);
        const relative: string = `${sitePath}/Documentos compartidos/${folderName}`.replace(/'/g,"''");
        const url: string = `${client.siteUrl}/_api/web/GetFolderByServerRelativeUrl('${encodeURIComponent(relative).replace(/%2F/gi,'/')}')/Files?$select=Name,ServerRelativeUrl,TimeLastModified,ListItemAllFields/Editor/Title&$expand=ListItemAllFields/Editor&$orderby=TimeLastModified desc&$top=30`;
        const response: SPHttpClientResponse = await this.spHttpClient.get(url,SPHttpClient.configurations.v1,{ headers: { Accept: 'application/json;odata=nometadata' } });
        if (!response.ok) return;
        const payload: { value?: Array<{ Name?: string; ServerRelativeUrl?: string; TimeLastModified?: string; ListItemAllFields?: { Editor?: { Title?: string } } }> } = await response.json() as { value?: Array<{ Name?: string; ServerRelativeUrl?: string; TimeLastModified?: string; ListItemAllFields?: { Editor?: { Title?: string } } }> };
        for (const file of payload.value || []) {
          const modified: Date | undefined = file.TimeLastModified ? new Date(file.TimeLastModified) : undefined;
          if (!modified || Number.isNaN(modified.getTime()) || modified.getTime() < since) continue;
          const serverRelative: string = file.ServerRelativeUrl || '';
          docs.push({
            title: file.Name || 'Documento',
            path: serverRelative ? `${new URL(client.siteUrl).origin}${serverRelative}` : '',
            siteUrl: client.siteUrl,
            siteTitle: client.name,
            modified,
            author: file.ListItemAllFields?.Editor?.Title || '',
            fileType: this.extension(file.Name || ''),
            area,
            clientCode: client.code
          });
        }
      } catch { /* inaccessible/empty folder */ }
    }));
    return docs;
  }

  private folderName(area: VgmArea): string {
    if (area === 'LEGAL') return 'Legal';
    if (area === 'TAX') return 'Tax';
    if (area === 'OUTSOURCING') return 'Outsourcing';
    return 'Auditoria';
  }

  private extension(name: string): string {
    const parts: string[] = name.split('.');
    return parts.length > 1 ? parts.pop()!.toLowerCase() : '';
  }
}
