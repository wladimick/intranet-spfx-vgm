import { Version } from '@microsoft/sp-core-library';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { SPPermission } from '@microsoft/sp-page-context';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

export interface IOrganigramaWebPartProps {}

export default class OrganigramaWebPart extends BaseClientSideWebPart<IOrganigramaWebPartProps> {
  private readonly folderName: string = 'TIBOX-Organigrama';
  private readonly fileName: string = 'organigrama-current.png';
  private zoomFactor: number = 2.5;
  private lensSize: number = 230;
  private isAdmin: boolean = false;
  private imageVersion: number = Date.now();

  public render(): void {
    this.isAdmin = this.context.pageContext.web.permissions.hasPermission(SPPermission.manageWeb);
    this.paint();
    this.bindEvents();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private paint(): void {
    const imageUrl: string = this.currentImageUrl();
    this.domElement.innerHTML = `
      <style>${this.styles()}</style>
      <section class="tbxOrg" data-tbx-organigrama>
        <header class="tbxOrgHead">
          <div>
            <span class="tbxOrgEyebrow">TIBOX · PERSONAS</span>
            <h2>Organigrama</h2>
            <p>Muévete sobre la imagen para ampliar los detalles.</p>
          </div>
          <div class="tbxOrgActions">
            <div class="tbxOrgControl" aria-label="Zoom de lupa">
              <span>Zoom</span>
              <button type="button" data-zoom-down aria-label="Disminuir zoom">−</button>
              <strong data-zoom-label>${this.zoomFactor.toFixed(1)}×</strong>
              <button type="button" data-zoom-up aria-label="Aumentar zoom">+</button>
            </div>
            <div class="tbxOrgControl" aria-label="Tamaño de lupa">
              <span>Lupa</span>
              <button type="button" data-lens-down aria-label="Reducir lupa">−</button>
              <strong data-lens-label>${this.lensSize}px</strong>
              <button type="button" data-lens-up aria-label="Aumentar lupa">+</button>
            </div>
            <button type="button" class="tbxOrgSecondary" data-fullscreen>⛶ Pantalla completa</button>
            ${this.isAdmin ? '<button type="button" class="tbxOrgPrimary" data-update>Actualizar organigrama</button>' : ''}
          </div>
        </header>

        <div class="tbxOrgInfo">
          <span class="tbxOrgDot"></span>
          <span>La lupa funciona localmente en tu navegador y no genera consultas a SharePoint mientras la mueves.</span>
        </div>

        <div class="tbxOrgViewport" data-viewport>
          <div class="tbxOrgStage" data-stage>
            <img class="tbxOrgImage" data-org-image src="${this.escapeAttr(imageUrl)}" alt="Organigrama TIBOX" draggable="false">
            <div class="tbxOrgLens" data-lens aria-hidden="true"></div>
            <div class="tbxOrgEmpty" data-empty hidden>
              <div class="tbxOrgEmptyIcon">⌕</div>
              <h3>No hay un organigrama publicado todavía</h3>
              <p>${this.isAdmin ? 'Usa “Actualizar organigrama” para cargar la primera imagen.' : 'Un administrador del sitio debe publicar el organigrama.'}</p>
            </div>
          </div>
        </div>

        ${this.isAdmin ? `
          <input type="file" data-file-input accept="image/png,image/jpeg,image/webp" hidden>
          <div class="tbxOrgOverlay" data-upload-modal>
            <div class="tbxOrgModal" role="dialog" aria-modal="true" aria-label="Actualizar organigrama">
              <div class="tbxOrgModalHead">
                <div><h3>Actualizar organigrama</h3><p>Solo usuarios con Control total / Manage Web pueden publicar una nueva imagen.</p></div>
                <button type="button" data-close-upload aria-label="Cerrar">×</button>
              </div>
              <div class="tbxOrgDrop" data-choose-file>
                <strong>Seleccionar imagen</strong>
                <span>PNG, JPG o WebP · idealmente de alta resolución</span>
              </div>
              <div class="tbxOrgUploadStatus" data-upload-status></div>
            </div>
          </div>` : ''}
      </section>`;
  }

  private bindEvents(): void {
    const image: HTMLImageElement | null = this.domElement.querySelector<HTMLImageElement>('[data-org-image]');
    const stage: HTMLElement | null = this.domElement.querySelector<HTMLElement>('[data-stage]');
    const lens: HTMLElement | null = this.domElement.querySelector<HTMLElement>('[data-lens]');
    const empty: HTMLElement | null = this.domElement.querySelector<HTMLElement>('[data-empty]');
    if (!image || !stage || !lens || !empty) return;

    image.addEventListener('load', (): void => {
      empty.hidden = true;
      image.hidden = false;
      this.prepareLens(image, lens);
    });

    image.addEventListener('error', (): void => {
      image.hidden = true;
      lens.classList.remove('is-visible');
      empty.hidden = false;
    });

    stage.addEventListener('pointermove', (event: PointerEvent): void => this.moveLens(event, image, stage, lens));
    stage.addEventListener('pointerleave', (): void => lens.classList.remove('is-visible'));
    stage.addEventListener('pointerenter', (): void => { if (!image.hidden && image.complete) lens.classList.add('is-visible'); });

    this.domElement.querySelector('[data-zoom-down]')?.addEventListener('click', (): void => {
      this.zoomFactor = Math.max(1.5, Number((this.zoomFactor - 0.5).toFixed(1)));
      this.updateControls(image, lens);
    });
    this.domElement.querySelector('[data-zoom-up]')?.addEventListener('click', (): void => {
      this.zoomFactor = Math.min(5, Number((this.zoomFactor + 0.5).toFixed(1)));
      this.updateControls(image, lens);
    });
    this.domElement.querySelector('[data-lens-down]')?.addEventListener('click', (): void => {
      this.lensSize = Math.max(150, this.lensSize - 20);
      this.updateControls(image, lens);
    });
    this.domElement.querySelector('[data-lens-up]')?.addEventListener('click', (): void => {
      this.lensSize = Math.min(360, this.lensSize + 20);
      this.updateControls(image, lens);
    });

    this.domElement.querySelector('[data-fullscreen]')?.addEventListener('click', (): void => {
      const viewport: HTMLElement | null = this.domElement.querySelector<HTMLElement>('[data-viewport]');
      if (!viewport) return;
      if (document.fullscreenElement) void document.exitFullscreen();
      else void viewport.requestFullscreen();
    });

    if (this.isAdmin) this.bindAdminEvents(image, lens, empty);
  }

  private bindAdminEvents(image: HTMLImageElement, lens: HTMLElement, empty: HTMLElement): void {
    const modal: HTMLElement | null = this.domElement.querySelector<HTMLElement>('[data-upload-modal]');
    const input: HTMLInputElement | null = this.domElement.querySelector<HTMLInputElement>('[data-file-input]');
    const status: HTMLElement | null = this.domElement.querySelector<HTMLElement>('[data-upload-status]');
    if (!modal || !input || !status) return;

    const open: () => void = (): void => modal.classList.add('is-open');
    const close: () => void = (): void => modal.classList.remove('is-open');

    this.domElement.querySelector('[data-update]')?.addEventListener('click', open);
    this.domElement.querySelector('[data-close-upload]')?.addEventListener('click', close);
    this.domElement.querySelector('[data-choose-file]')?.addEventListener('click', (): void => input.click());
    modal.addEventListener('click', (event: MouseEvent): void => { if (event.target === modal) close(); });

    input.addEventListener('change', async (): Promise<void> => {
      const file: File | undefined = input.files?.[0];
      if (!file) return;
      status.className = 'tbxOrgUploadStatus is-working';
      status.textContent = 'Preparando imagen…';
      try {
        const png: Blob = await this.toPng(file);
        status.textContent = 'Publicando en SharePoint…';
        await this.uploadCurrentImage(png);
        this.imageVersion = Date.now();
        image.hidden = false;
        empty.hidden = true;
        image.src = this.currentImageUrl();
        this.prepareLens(image, lens);
        status.className = 'tbxOrgUploadStatus is-success';
        status.textContent = 'Organigrama actualizado correctamente.';
        window.setTimeout(close, 900);
      } catch (error) {
        console.error('TIBOX Organigrama upload failed', error);
        status.className = 'tbxOrgUploadStatus is-error';
        status.textContent = 'No fue posible actualizar el organigrama. Revisa permisos y vuelve a intentar.';
      } finally {
        input.value = '';
      }
    });
  }

  private prepareLens(image: HTMLImageElement, lens: HTMLElement): void {
    lens.style.width = `${this.lensSize}px`;
    lens.style.height = `${this.lensSize}px`;
    lens.style.backgroundImage = `url("${image.src}")`;
    lens.style.backgroundRepeat = 'no-repeat';
  }

  private moveLens(event: PointerEvent, image: HTMLImageElement, stage: HTMLElement, lens: HTMLElement): void {
    if (image.hidden || !image.complete || !image.naturalWidth) return;
    const rect: DOMRect = image.getBoundingClientRect();
    const x: number = event.clientX - rect.left;
    const y: number = event.clientY - rect.top;
    if (x < 0 || y < 0 || x > rect.width || y > rect.height) {
      lens.classList.remove('is-visible');
      return;
    }

    const stageRect: DOMRect = stage.getBoundingClientRect();
    const localLeft: number = event.clientX - stageRect.left - this.lensSize / 2;
    const localTop: number = event.clientY - stageRect.top - this.lensSize / 2;
    lens.style.left = `${localLeft}px`;
    lens.style.top = `${localTop}px`;
    lens.style.width = `${this.lensSize}px`;
    lens.style.height = `${this.lensSize}px`;
    lens.style.backgroundSize = `${rect.width * this.zoomFactor}px ${rect.height * this.zoomFactor}px`;
    lens.style.backgroundPosition = `${-(x * this.zoomFactor - this.lensSize / 2)}px ${-(y * this.zoomFactor - this.lensSize / 2)}px`;
    lens.classList.add('is-visible');
  }

  private updateControls(image: HTMLImageElement, lens: HTMLElement): void {
    const zoomLabel: HTMLElement | null = this.domElement.querySelector<HTMLElement>('[data-zoom-label]');
    const lensLabel: HTMLElement | null = this.domElement.querySelector<HTMLElement>('[data-lens-label]');
    if (zoomLabel) zoomLabel.textContent = `${this.zoomFactor.toFixed(1)}×`;
    if (lensLabel) lensLabel.textContent = `${this.lensSize}px`;
    this.prepareLens(image, lens);
  }

  private async toPng(file: File): Promise<Blob> {
    if (!/^image\/(png|jpeg|webp)$/i.test(file.type)) throw new Error('Unsupported image type');
    const objectUrl: string = URL.createObjectURL(file);
    try {
      const image: HTMLImageElement = new Image();
      image.src = objectUrl;
      await new Promise<void>((resolve: () => void, reject: (reason?: unknown) => void): void => {
        image.onload = (): void => resolve();
        image.onerror = (): void => reject(new Error('Invalid image'));
      });

      const maxDimension: number = 9000;
      const scale: number = Math.min(1, maxDimension / Math.max(image.naturalWidth, image.naturalHeight));
      const canvas: HTMLCanvasElement = document.createElement('canvas');
      canvas.width = Math.max(1, Math.round(image.naturalWidth * scale));
      canvas.height = Math.max(1, Math.round(image.naturalHeight * scale));
      const context: CanvasRenderingContext2D | null = canvas.getContext('2d', { alpha: false });
      if (!context) throw new Error('Canvas unavailable');
      context.fillStyle = '#ffffff';
      context.fillRect(0, 0, canvas.width, canvas.height);
      context.drawImage(image, 0, 0, canvas.width, canvas.height);

      return await new Promise<Blob>((resolve: (value: Blob) => void, reject: (reason?: unknown) => void): void => {
        canvas.toBlob((blob: Blob | null): void => blob ? resolve(blob) : reject(new Error('PNG conversion failed')), 'image/png', 0.96);
      });
    } finally {
      URL.revokeObjectURL(objectUrl);
    }
  }

  private async uploadCurrentImage(blob: Blob): Promise<void> {
    await this.ensureFolder();
    const webUrl: string = this.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
    const folderRelative: string = this.folderServerRelativeUrl();
    const endpoint: string = `${webUrl}/_api/web/GetFolderByServerRelativeUrl('${this.odata(folderRelative)}')/Files/Add(url='${this.fileName}',overwrite=true)`;
    const response: SPHttpClientResponse = await this.context.spHttpClient.post(
      endpoint,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata',
          'Content-Type': 'image/png'
        },
        body: blob as unknown as string
      }
    );
    if (!response.ok) throw new Error(`Upload failed (${response.status})`);
  }

  private async ensureFolder(): Promise<void> {
    const webUrl: string = this.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
    const folderRelative: string = this.folderServerRelativeUrl();
    const check: SPHttpClientResponse = await this.context.spHttpClient.get(
      `${webUrl}/_api/web/GetFolderByServerRelativeUrl('${this.odata(folderRelative)}')?$select=Name`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: 'application/json;odata=nometadata' } }
    );
    if (check.ok) return;

    const parentRelative: string = this.joinServerRelative(this.context.pageContext.web.serverRelativeUrl, 'SiteAssets');
    const create: SPHttpClientResponse = await this.context.spHttpClient.post(
      `${webUrl}/_api/web/GetFolderByServerRelativeUrl('${this.odata(parentRelative)}')/folders`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose'
        },
        body: JSON.stringify({
          __metadata: { type: 'SP.Folder' },
          ServerRelativeUrl: folderRelative
        })
      }
    );
    if (!create.ok && create.status !== 409) throw new Error(`Folder creation failed (${create.status})`);
  }

  private currentImageUrl(): string {
    const webUrl: string = this.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
    return `${webUrl}/SiteAssets/${this.folderName}/${this.fileName}?v=${this.imageVersion}`;
  }

  private folderServerRelativeUrl(): string {
    return this.joinServerRelative(this.context.pageContext.web.serverRelativeUrl, `SiteAssets/${this.folderName}`);
  }

  private joinServerRelative(base: string, suffix: string): string {
    const cleanBase: string = (base || '').replace(/\/$/, '');
    const cleanSuffix: string = suffix.replace(/^\//, '');
    return `${cleanBase}/${cleanSuffix}`.replace(/^\/\//, '/');
  }

  private odata(value: string): string {
    return value.replace(/'/g, "''");
  }

  private escapeAttr(value: string): string {
    return value.replace(/[&<>\"']/g, (char: string): string => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[char] || char));
  }

  private styles(): string {
    return `
      .tbxOrg,.tbxOrg *{box-sizing:border-box}
      .tbxOrg{--ink:#001233;--muted:#6f7787;--line:#e7e9ee;--support:#0e9cdc;--dark:#000310;width:100%;font-family:'Segoe UI',Arial,sans-serif;color:var(--ink);background:#fff;border:1px solid var(--line);border-radius:22px;overflow:hidden;box-shadow:0 12px 34px rgba(0,18,51,.07)}
      .tbxOrgHead{display:flex;align-items:flex-start;justify-content:space-between;gap:24px;padding:22px 24px;border-bottom:1px solid var(--line);background:linear-gradient(180deg,#fff,#fbfcfe)}
      .tbxOrgHead h2{margin:4px 0 3px;font-size:24px;line-height:1.1}.tbxOrgHead p{margin:0;color:var(--muted);font-size:12px}.tbxOrgEyebrow{display:block;color:var(--support);font-size:9px;font-weight:800;letter-spacing:.16em}
      .tbxOrgActions{display:flex;align-items:center;justify-content:flex-end;gap:8px;flex-wrap:wrap}.tbxOrgControl{height:38px;display:flex;align-items:center;gap:5px;padding:0 7px;border:1px solid var(--line);border-radius:11px;background:#fff}.tbxOrgControl>span{padding:0 4px;color:var(--muted);font-size:9px;font-weight:700;text-transform:uppercase}.tbxOrgControl strong{min-width:43px;text-align:center;font-size:10px}.tbxOrgControl button{width:26px;height:26px;border:0;border-radius:7px;background:#f1f4f8;color:var(--ink);cursor:pointer;font-weight:800}
      .tbxOrgPrimary,.tbxOrgSecondary{height:38px;padding:0 13px;border-radius:10px;font-size:10.5px;font-weight:750;cursor:pointer}.tbxOrgPrimary{border:0;background:var(--dark);color:#fff}.tbxOrgSecondary{border:1px solid var(--line);background:#fff;color:var(--ink)}
      .tbxOrgInfo{display:flex;align-items:center;gap:8px;padding:9px 24px;border-bottom:1px solid var(--line);background:#f9fbfd;color:#697386;font-size:9.5px}.tbxOrgDot{width:7px;height:7px;border-radius:50%;background:#55d9a6;box-shadow:0 0 0 4px rgba(85,217,166,.10)}
      .tbxOrgViewport{position:relative;width:100%;overflow:auto;background:#eef2f6}.tbxOrgViewport:fullscreen{padding:20px;background:#111827}.tbxOrgStage{position:relative;min-width:720px;background:#fff}.tbxOrgImage{display:block;width:100%;height:auto;user-select:none;-webkit-user-drag:none;cursor:none}
      .tbxOrgLens{position:absolute;z-index:5;display:none;border-radius:50%;pointer-events:none;border:5px solid rgba(255,255,255,.98);box-shadow:0 14px 44px rgba(0,18,51,.35),0 0 0 1px rgba(0,18,51,.18);background-color:#fff}.tbxOrgLens.is-visible{display:block}.tbxOrgLens:after{content:'';position:absolute;inset:50% auto auto 50%;width:10px;height:10px;transform:translate(-50%,-50%);border:1px solid rgba(0,18,51,.26);border-radius:50%}
      .tbxOrgEmpty{min-height:380px;padding:40px;display:flex;flex-direction:column;align-items:center;justify-content:center;text-align:center;color:var(--muted)}.tbxOrgEmpty[hidden]{display:none}.tbxOrgEmptyIcon{width:64px;height:64px;display:grid;place-items:center;border-radius:20px;background:#eef8fc;color:var(--support);font-size:28px}.tbxOrgEmpty h3{margin:16px 0 5px;color:var(--ink);font-size:17px}.tbxOrgEmpty p{margin:0;max-width:420px;font-size:11px;line-height:1.5}
      .tbxOrgOverlay{position:fixed;inset:0;z-index:1800;display:none;align-items:center;justify-content:center;padding:20px;background:rgba(0,3,16,.62);backdrop-filter:blur(5px)}.tbxOrgOverlay.is-open{display:flex}.tbxOrgModal{width:min(520px,95vw);border-radius:20px;background:#fff;box-shadow:0 30px 90px rgba(0,0,0,.35);overflow:hidden}.tbxOrgModalHead{display:flex;justify-content:space-between;gap:18px;padding:20px 22px;border-bottom:1px solid var(--line)}.tbxOrgModalHead h3{margin:0;font-size:17px}.tbxOrgModalHead p{margin:5px 0 0;color:var(--muted);font-size:10px}.tbxOrgModalHead button{width:32px;height:32px;border:0;border-radius:9px;background:#f1f3f7;cursor:pointer}.tbxOrgDrop{margin:22px;padding:34px 20px;border:1.5px dashed #b8c3d1;border-radius:16px;background:#f9fbfd;text-align:center;cursor:pointer}.tbxOrgDrop strong{display:block;color:var(--ink);font-size:13px}.tbxOrgDrop span{display:block;margin-top:5px;color:var(--muted);font-size:10px}.tbxOrgUploadStatus{min-height:18px;margin:-8px 22px 22px;font-size:10px}.tbxOrgUploadStatus.is-working{color:#0e9cdc}.tbxOrgUploadStatus.is-success{color:#188c62}.tbxOrgUploadStatus.is-error{color:#c94335}
      @media(max-width:900px){.tbxOrgHead{flex-direction:column}.tbxOrgActions{justify-content:flex-start}.tbxOrgStage{min-width:640px}}
    `;
  }
}
