import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import VgmDataService, { ClientItem } from './VgmDataService';
import { VgmArea, VgmClientAccess } from './VgmSearchService';

const AREAS: VgmArea[] = ['LEGAL','TAX','OUTSOURCING','AUDITORIA'];

export default class VgmFolderAccessService {
  public constructor(
    private readonly spHttpClient: SPHttpClient,
    private readonly dataService: VgmDataService
  ) {}

  /**
   * Fallback exacto para VGM: prueba la carpeta concreta con la sesión del usuario.
   * Un 200 significa que SharePoint le permite consultar esa carpeta; 401/403/404
   * significa que no debe mostrarse. Se ejecuta solo cuando Search no entrega accesos.
   */
  public async scanAccessibleClients(
    onProgress?: (completed: number,total: number,found: number) => void,
    concurrency: number = 10
  ): Promise<VgmClientAccess[]> {
    const clients: ClientItem[] = await this.dataService.getClients();
    const result: VgmClientAccess[] = [];
    let cursor: number = 0;
    let completed: number = 0;

    const worker = async (): Promise<void> => {
      while (true) {
        const index: number = cursor++;
        if (index >= clients.length) return;
        const access: VgmClientAccess | undefined = await this.probeClient(clients[index]);
        if (access) result.push(access);
        completed += 1;
        if (onProgress) onProgress(completed,clients.length,result.length);
      }
    };

    const workers: Promise<void>[] = [];
    const count: number = Math.max(1,Math.min(concurrency,20));
    for (let i: number = 0; i < count; i++) workers.push(worker());
    await Promise.all(workers);

    return result.sort((a: VgmClientAccess,b: VgmClientAccess) => Number(a.code) - Number(b.code));
  }

  private async probeClient(client: ClientItem): Promise<VgmClientAccess | undefined> {
    const siteUrl: string = this.dataService.buildClientSiteBase(client);
    const checks: Array<Promise<[VgmArea,boolean]>> = AREAS.map(async (area: VgmArea): Promise<[VgmArea,boolean]> => {
      return [area,await this.canReadFolder(siteUrl,area)];
    });
    const resolved: Array<[VgmArea,boolean]> = await Promise.all(checks);
    const areas: Partial<Record<VgmArea,string>> = {};
    for (const [area,allowed] of resolved) {
      if (!allowed) continue;
      areas[area] = this.dataService.buildClientFolderUrl(client,this.folderName(area));
    }
    if (!Object.keys(areas).length) return undefined;
    return { code: client.code, name: client.name, siteUrl, areas };
  }

  private async canReadFolder(siteUrl: string, area: VgmArea): Promise<boolean> {
    try {
      const sitePath: string = new URL(siteUrl).pathname;
      const folderPath: string = `${sitePath}/Documentos compartidos/${this.folderName(area)}`;
      const escaped: string = folderPath.replace(/'/g,"''");
      const encoded: string = encodeURIComponent(escaped).replace(/%2F/gi,'/');
      const url: string = `${siteUrl}/_api/web/GetFolderByServerRelativeUrl('${encoded}')?$select=Name,ServerRelativeUrl,Exists`;
      const response: SPHttpClientResponse = await this.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1,
        { headers: { Accept: 'application/json;odata=nometadata' } }
      );
      if (!response.ok) return false;
      const payload: { Exists?: boolean } = await response.json() as { Exists?: boolean };
      return payload.Exists !== false;
    } catch {
      return false;
    }
  }

  private folderName(area: VgmArea): string {
    if (area === 'LEGAL') return 'Legal';
    if (area === 'TAX') return 'Tax';
    if (area === 'OUTSOURCING') return 'Outsourcing';
    return 'Auditoria';
  }
}
