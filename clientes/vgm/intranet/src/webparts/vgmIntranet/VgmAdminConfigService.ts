import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export type VgmConfigModuleKey = 'calendar' | 'birthdays' | 'financial' | 'news' | 'indicators' | 'gallery' | 'links' | 'personal' | 'recentDocuments' | 'activityClients' | 'activityAreas';
export type VgmActivityModuleKey = 'recentDocuments' | 'activityClients' | 'activityAreas';
export type VgmQuickLinkConfig = { title: string; url: string };

export type VgmAppearanceConfig = {
  primaryColor: string;
  secondaryColor: string;
  pageBackground: string;
  cardRadius: number;
  maxWidth: number;
  baseFontSize: number;
};

export type VgmPortalConfig = {
  appearance: VgmAppearanceConfig;
  modules: Record<VgmConfigModuleKey, boolean>;
  activityOrder: VgmActivityModuleKey[];
  quickLinks: VgmQuickLinkConfig[];
  customCss: string;
};

const LIST_TITLE: string = 'VGM Intranet - Configuracion';
const ADMIN_GROUP: string = 'VGM Intranet - Administradores';

export const DEFAULT_VGM_CONFIG: VgmPortalConfig = {
  appearance: {
    primaryColor: '#334F82',
    secondaryColor: '#0A5E95',
    pageBackground: '#2F405F',
    cardRadius: 6,
    maxWidth: 1500,
    baseFontSize: 14
  },
  modules: {
    calendar: true,
    birthdays: true,
    financial: true,
    news: true,
    indicators: true,
    gallery: true,
    links: true,
    personal: true,
    recentDocuments: true,
    activityClients: true,
    activityAreas: true
  },
  activityOrder: ['recentDocuments','activityClients','activityAreas'],
  quickLinks: [],
  customCss: ''
};

export default class VgmAdminConfigService {
  public static readonly sourceWebUrl: string = 'https://vgmconsultants.sharepoint.com/sites/intranet';

  public constructor(private readonly spHttpClient: SPHttpClient, private readonly userEmail: string) {}

  public async isAdmin(): Promise<boolean> {
    try {
      const response: SPHttpClientResponse = await this.spHttpClient.get(
        `${VgmAdminConfigService.sourceWebUrl}/_api/web/currentuser?$select=IsSiteAdmin,Email`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: 'application/json;odata=nometadata' } }
      );
      if (response.ok) {
        const current: { IsSiteAdmin?: boolean; Email?: string } = await response.json() as { IsSiteAdmin?: boolean; Email?: string };
        if (current.IsSiteAdmin) return true;
      }
    } catch { /* fallback group */ }

    try {
      const group: string = ADMIN_GROUP.replace(/'/g,"''");
      const response: SPHttpClientResponse = await this.spHttpClient.get(
        `${VgmAdminConfigService.sourceWebUrl}/_api/web/sitegroups/getbyname('${group}')/users?$select=Email&$top=200`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: 'application/json;odata=nometadata' } }
      );
      if (!response.ok) return false;
      const payload: { value?: Array<{ Email?: string }> } = await response.json() as { value?: Array<{ Email?: string }> };
      const email: string = (this.userEmail || '').toLowerCase();
      return (payload.value || []).some((item: { Email?: string }) => (item.Email || '').toLowerCase() === email);
    } catch {
      return false;
    }
  }

  public async load(): Promise<VgmPortalConfig> {
    try {
      const escaped: string = LIST_TITLE.replace(/'/g,"''");
      const url: string = `${VgmAdminConfigService.sourceWebUrl}/_api/web/lists/getbytitle('${escaped}')/items?$select=Title,ConfigValue&$top=20`;
      const response: SPHttpClientResponse = await this.spHttpClient.get(url,SPHttpClient.configurations.v1,{ headers: { Accept: 'application/json;odata=nometadata' } });
      if (!response.ok) return this.cloneDefault();
      const payload: { value?: Array<{ Title?: string; ConfigValue?: string }> } = await response.json() as { value?: Array<{ Title?: string; ConfigValue?: string }> };
      const map: Map<string,string> = new Map<string,string>();
      for (const item of payload.value || []) if (item.Title) map.set(item.Title,item.ConfigValue || '');
      const config: VgmPortalConfig = this.cloneDefault();
      const appearance: string | undefined = map.get('appearance');
      const modules: string | undefined = map.get('modules');
      const activityOrder: string | undefined = map.get('activityOrder');
      const quickLinks: string | undefined = map.get('quickLinks');
      if (appearance) config.appearance = { ...config.appearance, ...JSON.parse(appearance) };
      if (modules) config.modules = { ...config.modules, ...JSON.parse(modules) };
      if (activityOrder) config.activityOrder = JSON.parse(activityOrder) as VgmActivityModuleKey[];
      if (quickLinks) config.quickLinks = JSON.parse(quickLinks) as VgmQuickLinkConfig[];
      config.customCss = map.get('customCss') || '';
      return config;
    } catch {
      return this.cloneDefault();
    }
  }

  public async save(config: VgmPortalConfig): Promise<void> {
    await this.ensureList();
    await this.upsert('appearance',JSON.stringify(config.appearance));
    await this.upsert('modules',JSON.stringify(config.modules));
    await this.upsert('activityOrder',JSON.stringify(config.activityOrder));
    await this.upsert('quickLinks',JSON.stringify(config.quickLinks || []));
    await this.upsert('customCss',config.customCss || '');
  }

  private async ensureList(): Promise<void> {
    const escaped: string = LIST_TITLE.replace(/'/g,"''");
    const check: SPHttpClientResponse = await this.spHttpClient.get(
      `${VgmAdminConfigService.sourceWebUrl}/_api/web/lists/getbytitle('${escaped}')?$select=Id`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: 'application/json;odata=nometadata' } }
    );
    if (check.ok) return;

    const create: SPHttpClientResponse = await this.spHttpClient.post(
      `${VgmAdminConfigService.sourceWebUrl}/_api/web/lists`,
      SPHttpClient.configurations.v1,
      {
        headers: { Accept: 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=nometadata' },
        body: JSON.stringify({ Title: LIST_TITLE, BaseTemplate: 100, Description: 'Configuración administrable de VGM Intranet SPFx' })
      }
    );
    if (!create.ok) throw new Error(`No fue posible crear ${LIST_TITLE} (${create.status})`);

    const fieldXml: string = `<Field DisplayName="ConfigValue" Name="ConfigValue" StaticName="ConfigValue" Type="Note" NumLines="30" RichText="FALSE" />`;
    const field: SPHttpClientResponse = await this.spHttpClient.post(
      `${VgmAdminConfigService.sourceWebUrl}/_api/web/lists/getbytitle('${escaped}')/fields/CreateFieldAsXml`,
      SPHttpClient.configurations.v1,
      {
        headers: { Accept: 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=nometadata' },
        body: JSON.stringify({ parameters: { SchemaXml: fieldXml, Options: 0 } })
      }
    );
    if (!field.ok) throw new Error(`No fue posible crear ConfigValue (${field.status})`);
  }

  private async upsert(key: string, value: string): Promise<void> {
    const escapedTitle: string = LIST_TITLE.replace(/'/g,"''");
    const escapedKey: string = key.replace(/'/g,"''");
    const lookup: SPHttpClientResponse = await this.spHttpClient.get(
      `${VgmAdminConfigService.sourceWebUrl}/_api/web/lists/getbytitle('${escapedTitle}')/items?$select=Id&$filter=Title eq '${escapedKey}'&$top=1`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: 'application/json;odata=nometadata' } }
    );
    const payload: { value?: Array<{ Id?: number }> } = lookup.ok ? await lookup.json() as { value?: Array<{ Id?: number }> } : {};
    const id: number | undefined = payload.value?.[0]?.Id;
    if (id) {
      const response: SPHttpClientResponse = await this.spHttpClient.post(
        `${VgmAdminConfigService.sourceWebUrl}/_api/web/lists/getbytitle('${escapedTitle}')/items(${id})`,
        SPHttpClient.configurations.v1,
        {
          headers: { Accept: 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=nometadata', 'IF-MATCH': '*', 'X-HTTP-Method': 'MERGE' },
          body: JSON.stringify({ ConfigValue: value })
        }
      );
      if (!response.ok) throw new Error(`No fue posible actualizar ${key} (${response.status})`);
      return;
    }
    const response: SPHttpClientResponse = await this.spHttpClient.post(
      `${VgmAdminConfigService.sourceWebUrl}/_api/web/lists/getbytitle('${escapedTitle}')/items`,
      SPHttpClient.configurations.v1,
      {
        headers: { Accept: 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=nometadata' },
        body: JSON.stringify({ Title: key, ConfigValue: value })
      }
    );
    if (!response.ok) throw new Error(`No fue posible guardar ${key} (${response.status})`);
  }

  private cloneDefault(): VgmPortalConfig {
    return JSON.parse(JSON.stringify(DEFAULT_VGM_CONFIG)) as VgmPortalConfig;
  }
}
