import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

export interface IVgmDemoWebPartProps {}

export default class VgmDemoWebPart extends BaseClientSideWebPart<IVgmDemoWebPartProps> {
  public render(): void {
    this.domElement.innerHTML = `
      <section style="font-family: Segoe UI, Arial, sans-serif; padding: 28px; border: 1px solid #e1e1e1; border-radius: 12px; background: #ffffff; box-shadow: 0 4px 16px rgba(0,0,0,.06);">
        <div style="font-size: 13px; font-weight: 600; color: #666; margin-bottom: 8px;">PRUEBA SHAREPOINT FRAMEWORK</div>
        <h2 style="font-size: 28px; line-height: 1.2; margin: 0 0 10px; color: #1b1b1b;">Hola desde SPFx 👋</h2>
        <p style="font-size: 16px; line-height: 1.6; margin: 0 0 6px; color: #444;">Este Web Part esta funcionando dentro de SharePoint.</p>
        <p style="font-size: 14px; margin: 0 0 22px; color: #666;">Sitio actual: <strong data-site-title></strong></p>
        <button data-action="test" type="button" style="border: 0; border-radius: 8px; padding: 10px 16px; background: #0078d4; color: #fff; font-weight: 600; cursor: pointer;">Probar boton</button>
        <span data-status style="display: inline-block; margin-left: 12px; color: #107c10; font-weight: 600;"></span>
      </section>
    `;

    const siteTitleElement: HTMLElement | null = this.domElement.querySelector('[data-site-title]');
    if (siteTitleElement) {
      siteTitleElement.textContent = this.context.pageContext.web.title || 'SharePoint';
    }

    const button: HTMLButtonElement | null = this.domElement.querySelector('[data-action="test"]');
    const status: HTMLElement | null = this.domElement.querySelector('[data-status]');

    button?.addEventListener('click', () => {
      if (status) {
        status.textContent = '¡Funciona! SPFx esta ejecutando JavaScript.';
      }
    });
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}
