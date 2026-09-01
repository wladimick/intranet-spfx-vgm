import { MSGraphClientV3, SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import TiboxHubWebPart from './TiboxHubWebPart';

type PersonMailItem = {
  Title?: string;
  Email?: string;
};

type GraphDateTime = {
  dateTime?: string;
  timeZone?: string;
};

type GraphScheduleItem = {
  status?: string;
  subject?: string;
  start?: GraphDateTime;
  end?: GraphDateTime;
};

type GraphScheduleInformation = {
  scheduleId?: string;
  availabilityView?: string;
  scheduleItems?: GraphScheduleItem[];
};

type GraphScheduleResponse = {
  value?: GraphScheduleInformation[];
};

type RoomInterval = {
  start: number;
  end: number;
  busy: boolean;
  label: string;
};

export default class TiboxHubGraphWebPart extends TiboxHubWebPart {
  private readonly roomMailbox: string = 'sala@tibox.cl';
  private readonly outlookTimeZone: string = 'Pacific SA Standard Time';
  private readonly displayTimeZone: string = 'America/Santiago';
  private enhancementRun: number = 0;
  private objectUrls: string[] = [];

  public render(): void {
    const run: number = ++this.enhancementRun;
    super.render();
    void this.enhanceMicrosoft365(run);
  }

  protected onDispose(): void {
    this.objectUrls.forEach((url: string) => URL.revokeObjectURL(url));
    this.objectUrls = [];
    super.onDispose();
  }

  private async enhanceMicrosoft365(run: number): Promise<void> {
    const hub: HTMLElement | undefined = await this.waitForHub(run);
    if (!hub || run !== this.enhancementRun) return;

    this.injectEnhancementStyles(hub);
    const roomCard: HTMLElement | undefined = this.createRoomCard();
    const side: HTMLElement | null = hub.querySelector<HTMLElement>('.tbxSide, .tbx-rail');
    if (roomCard && side) {
      side.insertBefore(roomCard, side.firstElementChild);
    }

    try {
      const client: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');
      await Promise.all([
        this.enhanceProfilePhotos(client, run),
        this.enhanceCurrentUserPhoto(client, run),
        this.loadRoomAvailability(client, hub, run)
      ]);
    } catch (error) {
      console.warn('TIBOX HUB Microsoft 365 integration unavailable', error);
      this.renderRoomPermissionPending(hub);
    }
  }

  private async waitForHub(run: number): Promise<HTMLElement | undefined> {
    for (let attempt: number = 0; attempt < 80; attempt++) {
      if (run !== this.enhancementRun) return undefined;
      const hub: HTMLElement | null = this.domElement.querySelector<HTMLElement>('.tbxHub, .tbx-hub');
      if (hub) return hub;
      await this.delay(100);
    }
    return undefined;
  }

  private injectEnhancementStyles(hub: HTMLElement): void {
    if (hub.querySelector('[data-tbx-m365-styles]')) return;
    const style: HTMLStyleElement = document.createElement('style');
    style.setAttribute('data-tbx-m365-styles', 'true');
    style.textContent = `
      .tbxGraphAvatarImg{width:100%;height:100%;display:block;object-fit:cover;border-radius:50%}
      .tbxGraphMe{width:38px;height:38px;padding:0;border:1px solid var(--border,var(--tbx-border,rgba(255,255,255,.08)));border-radius:50%;overflow:hidden;background:var(--surface,var(--tbx-surface,#0A1130));display:grid;place-items:center;color:var(--text,var(--tbx-text,#F4F7FF));font-size:10px;font-weight:800}
      .tbxGraphRoom{border:1px solid var(--border,var(--tbx-border,rgba(255,255,255,.08)));border-radius:18px;background:var(--surface,var(--tbx-surface,#0A1130));overflow:hidden;color:var(--text,var(--tbx-text,#F4F7FF))}
      .tbxGraphRoomHead{display:flex;align-items:center;justify-content:space-between;gap:10px;padding:14px 16px;border-bottom:1px solid var(--border,var(--tbx-border,rgba(255,255,255,.08)))}
      .tbxGraphRoomTitle{display:flex;align-items:center;gap:9px;font-size:13px;font-weight:800}
      .tbxGraphRoomIcon{width:30px;height:30px;display:grid;place-items:center;border-radius:9px;color:#00D1FF;background:rgba(0,209,255,.09)}
      .tbxGraphRoomIcon svg{width:17px;height:17px;stroke:currentColor;fill:none;stroke-width:1.8;stroke-linecap:round;stroke-linejoin:round}
      .tbxGraphRoomBody{padding:14px 16px 16px}
      .tbxGraphRoomStatus{display:flex;align-items:center;gap:8px;font-size:14px;font-weight:800;margin-bottom:5px}
      .tbxGraphStatusDot{width:8px;height:8px;border-radius:50%;background:#55D9A6;box-shadow:0 0 0 4px rgba(85,217,166,.09)}
      .tbxGraphRoom.is-busy .tbxGraphStatusDot{background:#FF725E;box-shadow:0 0 0 4px rgba(255,114,94,.09)}
      .tbxGraphRoom.is-pending .tbxGraphStatusDot{background:#FFB200;box-shadow:0 0 0 4px rgba(255,178,0,.09)}
      .tbxGraphRoomSub{font-size:10.5px;line-height:1.45;color:var(--muted,var(--tbx-text-muted,#9BA6C4));min-height:30px}
      .tbxGraphRoomNext{margin-top:12px;padding:10px 11px;border:1px solid var(--border,var(--tbx-border,rgba(255,255,255,.08)));border-radius:11px;background:var(--surface2,var(--tbx-surface-2,#121A40))}
      .tbxGraphRoomNext label{display:block;color:var(--subtle,var(--tbx-text-subtle,#5E6A8A));font-size:8.5px;font-weight:800;text-transform:uppercase;letter-spacing:.06em}
      .tbxGraphRoomNext strong{display:block;margin-top:4px;font-size:11px;color:var(--text,var(--tbx-text,#F4F7FF))}
      .tbxGraphRoomBtn{width:100%;margin-top:12px;padding:9px 11px;border:1px solid rgba(0,209,255,.22);border-radius:9px;background:rgba(0,209,255,.08);color:#00D1FF;font-size:10.5px;font-weight:800;cursor:pointer}
      .tbxGraphRoomBtn:hover{background:rgba(0,209,255,.14)}
      .tbxGraphRoomOverlay{position:fixed;inset:0;z-index:1400;display:none;align-items:center;justify-content:center;padding:24px;background:rgba(0,3,16,.74);backdrop-filter:blur(7px)}
      .tbxGraphRoomOverlay.open{display:flex}
      .tbxGraphRoomModal{width:min(650px,96vw);max-height:86vh;overflow:auto;border:1px solid rgba(255,255,255,.10);border-radius:22px;background:#0A1130;color:#F4F7FF;box-shadow:0 32px 100px rgba(0,0,0,.48)}
      .tbxGraphRoomModalHead{position:sticky;top:0;z-index:2;padding:17px 19px;display:flex;align-items:flex-start;justify-content:space-between;gap:14px;background:rgba(10,17,48,.97);border-bottom:1px solid rgba(255,255,255,.08)}
      .tbxGraphRoomModalHead h3{margin:0;font-size:16px}.tbxGraphRoomModalHead p{margin:4px 0 0;color:#9BA6C4;font-size:10.5px}
      .tbxGraphClose{width:34px;height:34px;border:1px solid rgba(255,255,255,.08);border-radius:9px;background:#121A40;color:#F4F7FF;cursor:pointer}
      .tbxGraphTimeline{padding:15px 18px 20px;display:flex;flex-direction:column;gap:7px}
      .tbxGraphSlot{display:grid;grid-template-columns:90px 1fr auto;align-items:center;gap:10px;padding:10px 11px;border:1px solid rgba(255,255,255,.07);border-radius:11px;background:#121A40}
      .tbxGraphSlotTime{font-size:10.5px;font-weight:800;color:#F4F7FF}.tbxGraphSlotLabel{min-width:0;font-size:10.5px;color:#9BA6C4;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
      .tbxGraphSlotState{padding:4px 7px;border-radius:999px;font-size:8.5px;font-weight:800;color:#55D9A6;background:rgba(85,217,166,.08)}
      .tbxGraphSlot.busy .tbxGraphSlotState{color:#FF725E;background:rgba(255,114,94,.08)}
      @media(max-width:760px){.tbxGraphSlot{grid-template-columns:74px 1fr}.tbxGraphSlotState{grid-column:2;justify-self:start}.tbxGraphRoomOverlay{padding:12px}}
    `;
    hub.prepend(style);
  }

  private createRoomCard(): HTMLElement {
    const card: HTMLElement = document.createElement('section');
    card.className = 'tbxGraphRoom is-pending';
    card.setAttribute('data-tbx-room-card', 'true');
    card.innerHTML = `
      <div class="tbxGraphRoomHead">
        <div class="tbxGraphRoomTitle">
          <span class="tbxGraphRoomIcon" aria-hidden="true"><svg viewBox="0 0 24 24"><path d="M4 21V5a2 2 0 0 1 2-2h8a2 2 0 0 1 2 2v16"/><path d="M16 8h3a1 1 0 0 1 1 1v12"/><path d="M8 12h.01M8 16h.01M12 12h.01M12 16h.01"/></svg></span>
          <span>Sala de reuniones</span>
        </div>
      </div>
      <div class="tbxGraphRoomBody">
        <div class="tbxGraphRoomStatus"><span class="tbxGraphStatusDot"></span><span data-tbx-room-status>Consultando disponibilidad…</span></div>
        <div class="tbxGraphRoomSub" data-tbx-room-sub>Sala de reuniones Tibox · ${this.roomMailbox}</div>
        <div class="tbxGraphRoomNext" data-tbx-room-next><label>Microsoft 365</label><strong>Conectando con Outlook…</strong></div>
        <button type="button" class="tbxGraphRoomBtn" data-tbx-room-open disabled>Ver disponibilidad</button>
      </div>`;
    return card;
  }

  private async enhanceProfilePhotos(client: MSGraphClientV3, run: number): Promise<void> {
    const emailMap: Map<string, string> = await this.getPeopleEmailMap();
    if (run !== this.enhancementRun || !emailMap.size) return;

    const rows: HTMLElement[] = Array.from(this.domElement.querySelectorAll<HTMLElement>('.tbxPerson[data-person], .tbxBirthday, .tbxMovement, .tbx-bday-row, .tbx-person-row, .tbx-movement-row'));
    const photoCache: Map<string, string | undefined> = new Map<string, string | undefined>();

    for (const row of rows) {
      if (run !== this.enhancementRun) return;
      const rawName: string = (row.getAttribute('data-person') || row.querySelector<HTMLElement>('.tbxPersonText strong, .tbx-person-name')?.textContent || '').trim();
      const email: string | undefined = emailMap.get(this.normalize(rawName));
      if (!email) continue;

      let photoUrl: string | undefined = photoCache.get(email);
      if (!photoCache.has(email)) {
        photoUrl = await this.getPhotoUrl(client, `/users/${encodeURIComponent(email)}/photos/48x48/$value`);
        photoCache.set(email, photoUrl);
      }
      if (!photoUrl) continue;

      const avatar: HTMLElement | null = row.querySelector<HTMLElement>('.tbxAvatar, .tbx-avatar');
      if (avatar) this.applyPhoto(avatar, photoUrl, rawName);
    }
  }

  private async enhanceCurrentUserPhoto(client: MSGraphClientV3, run: number): Promise<void> {
    const photoUrl: string | undefined = await this.getPhotoUrl(client, '/me/photos/48x48/$value');
    if (!photoUrl || run !== this.enhancementRun) return;

    const currentImage: HTMLImageElement | null = this.domElement.querySelector<HTMLImageElement>('.tbx-hero-user-photo img, .tbx-avatar-btn img');
    if (currentImage) {
      currentImage.src = photoUrl;
      return;
    }

    const actions: HTMLElement | null = this.domElement.querySelector<HTMLElement>('.tbxTopActions, .tbx-header-actions');
    if (!actions || actions.querySelector('[data-tbx-me-photo]')) return;
    const holder: HTMLDivElement = document.createElement('div');
    holder.className = 'tbxGraphMe';
    holder.setAttribute('data-tbx-me-photo', 'true');
    holder.title = this.context.pageContext.user.displayName || 'Mi perfil';
    this.applyPhoto(holder, photoUrl, holder.title);
    actions.appendChild(holder);
  }

  private async getPhotoUrl(client: MSGraphClientV3, path: string): Promise<string | undefined> {
    try {
      const blob: Blob = await client.api(path).get() as Blob;
      if (!(blob instanceof Blob) || blob.size === 0) return undefined;
      const url: string = URL.createObjectURL(blob);
      this.objectUrls.push(url);
      return url;
    } catch {
      return undefined;
    }
  }

  private applyPhoto(avatar: HTMLElement, photoUrl: string, name: string): void {
    avatar.innerHTML = '';
    const image: HTMLImageElement = document.createElement('img');
    image.className = 'tbxGraphAvatarImg';
    image.src = photoUrl;
    image.alt = `Foto de ${name}`;
    avatar.appendChild(image);
  }

  private async getPeopleEmailMap(): Promise<Map<string, string>> {
    const map: Map<string, string> = new Map<string, string>();
    const webUrl: string = this.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
    const queries: string[] = [
      "TIBOX HUB - Colaboradores",
      "TIBOX HUB - Movimientos"
    ];

    for (const title of queries) {
      try {
        const endpoint: string = `${webUrl}/_api/web/lists/getbytitle('${this.odata(title)}')/items?$select=Title,Email&$top=250`;
        const response: SPHttpClientResponse = await this.context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1, { headers: { Accept: 'application/json;odata=nometadata' } });
        if (!response.ok) continue;
        const payload: { value?: PersonMailItem[] } = await response.json() as { value?: PersonMailItem[] };
        (payload.value || []).forEach((item: PersonMailItem) => {
          if (item.Title && item.Email) map.set(this.normalize(item.Title), item.Email.trim());
        });
      } catch {
        // Initials remain as fallback.
      }
    }
    return map;
  }

  private async loadRoomAvailability(client: MSGraphClientV3, hub: HTMLElement, run: number): Promise<void> {
    const dateKey: string = this.santiagoDateKey();
    const body: Record<string, unknown> = {
      schedules: [this.roomMailbox],
      startTime: { dateTime: `${dateKey}T08:00:00`, timeZone: this.outlookTimeZone },
      endTime: { dateTime: `${dateKey}T20:00:00`, timeZone: this.outlookTimeZone },
      availabilityViewInterval: 30
    };

    try {
      const response: GraphScheduleResponse = await client
        .api('/me/calendar/getSchedule')
        .header('Prefer', `outlook.timezone="${this.outlookTimeZone}"`)
        .post(body) as GraphScheduleResponse;
      if (run !== this.enhancementRun) return;

      const schedule: GraphScheduleInformation | undefined = response.value?.[0];
      if (!schedule) throw new Error('No schedule returned');
      const items: GraphScheduleItem[] = (schedule.scheduleItems || []).filter((item: GraphScheduleItem) => this.isBusyStatus(item.status));
      this.renderRoomState(hub, items, dateKey);
    } catch (error) {
      console.warn('TIBOX HUB room availability unavailable', error);
      this.renderRoomPermissionPending(hub);
    }
  }

  private renderRoomState(hub: HTMLElement, items: GraphScheduleItem[], dateKey: string): void {
    const card: HTMLElement | null = hub.querySelector<HTMLElement>('[data-tbx-room-card]');
    if (!card) return;

    const nowMinutes: number = this.santiagoCurrentMinutes();
    const sorted: GraphScheduleItem[] = [...items].sort((a: GraphScheduleItem, b: GraphScheduleItem) => this.timeMinutes(a.start?.dateTime) - this.timeMinutes(b.start?.dateTime));
    const current: GraphScheduleItem | undefined = sorted.find((item: GraphScheduleItem) => {
      const start: number = this.timeMinutes(item.start?.dateTime);
      const end: number = this.timeMinutes(item.end?.dateTime);
      return nowMinutes >= start && nowMinutes < end;
    });
    const next: GraphScheduleItem | undefined = sorted.find((item: GraphScheduleItem) => this.timeMinutes(item.start?.dateTime) > nowMinutes);

    card.classList.toggle('is-busy', Boolean(current));
    card.classList.remove('is-pending');
    const status: HTMLElement | null = card.querySelector<HTMLElement>('[data-tbx-room-status]');
    const sub: HTMLElement | null = card.querySelector<HTMLElement>('[data-tbx-room-sub]');
    const nextArea: HTMLElement | null = card.querySelector<HTMLElement>('[data-tbx-room-next]');
    const button: HTMLButtonElement | null = card.querySelector<HTMLButtonElement>('[data-tbx-room-open]');

    if (status) status.textContent = current ? `Ocupada hasta ${this.formatGraphTime(current.end?.dateTime)}` : 'Disponible ahora';
    if (sub) sub.textContent = current ? (current.subject || 'Hay una reserva activa en este momento.') : (next ? `Disponible hasta ${this.formatGraphTime(next.start?.dateTime)}` : 'Disponible durante el resto de la jornada.');
    if (nextArea) {
      if (next) nextArea.innerHTML = `<label>Próxima reserva</label><strong>${this.formatGraphTime(next.start?.dateTime)} – ${this.formatGraphTime(next.end?.dateTime)} · ${this.escape(next.subject || 'Reserva')}</strong>`;
      else nextArea.innerHTML = '<label>Próxima reserva</label><strong>No hay más reservas hoy</strong>';
    }
    if (button) {
      button.disabled = false;
      button.addEventListener('click', (): void => this.openRoomModal(hub, sorted, dateKey));
    }
  }

  private renderRoomPermissionPending(hub: HTMLElement): void {
    const card: HTMLElement | null = hub.querySelector<HTMLElement>('[data-tbx-room-card]');
    if (!card) return;
    card.classList.remove('is-busy');
    card.classList.add('is-pending');
    const status: HTMLElement | null = card.querySelector<HTMLElement>('[data-tbx-room-status]');
    const sub: HTMLElement | null = card.querySelector<HTMLElement>('[data-tbx-room-sub]');
    const nextArea: HTMLElement | null = card.querySelector<HTMLElement>('[data-tbx-room-next]');
    const button: HTMLButtonElement | null = card.querySelector<HTMLButtonElement>('[data-tbx-room-open]');
    if (status) status.textContent = 'Integración Microsoft 365 pendiente';
    if (sub) sub.textContent = 'Aprueba Calendars.ReadBasic en SharePoint → Acceso a API.';
    if (nextArea) nextArea.innerHTML = '<label>Sala de reuniones Tibox</label><strong>sala@tibox.cl</strong>';
    if (button) button.disabled = true;
  }

  private openRoomModal(hub: HTMLElement, items: GraphScheduleItem[], dateKey: string): void {
    hub.querySelector('[data-tbx-room-overlay]')?.remove();
    const overlay: HTMLDivElement = document.createElement('div');
    overlay.className = 'tbxGraphRoomOverlay open';
    overlay.setAttribute('data-tbx-room-overlay', 'true');
    const intervals: RoomInterval[] = this.buildIntervals(items);
    const formattedDate: string = this.formatDateKey(dateKey);
    overlay.innerHTML = `
      <div class="tbxGraphRoomModal" role="dialog" aria-modal="true" aria-label="Disponibilidad Sala de reuniones Tibox">
        <div class="tbxGraphRoomModalHead">
          <div><h3>Sala de reuniones Tibox</h3><p>${formattedDate} · disponibilidad de 08:00 a 20:00</p></div>
          <button type="button" class="tbxGraphClose" data-tbx-room-close aria-label="Cerrar">×</button>
        </div>
        <div class="tbxGraphTimeline">
          ${intervals.map((interval: RoomInterval) => `<div class="tbxGraphSlot ${interval.busy ? 'busy' : ''}"><span class="tbxGraphSlotTime">${this.minutesLabel(interval.start)} – ${this.minutesLabel(interval.end)}</span><span class="tbxGraphSlotLabel">${this.escape(interval.label)}</span><span class="tbxGraphSlotState">${interval.busy ? 'Ocupada' : 'Disponible'}</span></div>`).join('')}
        </div>
      </div>`;
    hub.appendChild(overlay);

    const close: (): void => void = (): void => overlay.remove();
    overlay.querySelector('[data-tbx-room-close]')?.addEventListener('click', close);
    overlay.addEventListener('click', (event: MouseEvent): void => { if (event.target === overlay) close(); });
    const keyHandler: (event: KeyboardEvent) => void = (event: KeyboardEvent): void => {
      if (event.key === 'Escape') {
        close();
        document.removeEventListener('keydown', keyHandler);
      }
    };
    document.addEventListener('keydown', keyHandler);
  }

  private buildIntervals(items: GraphScheduleItem[]): RoomInterval[] {
    const dayStart: number = 8 * 60;
    const dayEnd: number = 20 * 60;
    const busy: Array<{ start: number; end: number; label: string }> = items
      .map((item: GraphScheduleItem) => ({
        start: Math.max(dayStart, this.timeMinutes(item.start?.dateTime)),
        end: Math.min(dayEnd, this.timeMinutes(item.end?.dateTime)),
        label: item.subject || 'Reserva'
      }))
      .filter((item: { start: number; end: number }) => item.end > item.start)
      .sort((a: { start: number }, b: { start: number }) => a.start - b.start);

    const result: RoomInterval[] = [];
    let cursor: number = dayStart;
    busy.forEach((item: { start: number; end: number; label: string }) => {
      if (item.start > cursor) result.push({ start: cursor, end: item.start, busy: false, label: 'Sala disponible' });
      result.push({ start: item.start, end: item.end, busy: true, label: item.label });
      cursor = Math.max(cursor, item.end);
    });
    if (cursor < dayEnd) result.push({ start: cursor, end: dayEnd, busy: false, label: 'Sala disponible' });
    return result.length ? result : [{ start: dayStart, end: dayEnd, busy: false, label: 'Sala disponible' }];
  }

  private isBusyStatus(status?: string): boolean {
    const normalized: string = (status || '').toLowerCase();
    return normalized !== '' && normalized !== 'free';
  }

  private santiagoDateKey(): string {
    const parts: Intl.DateTimeFormatPart[] = new Intl.DateTimeFormat('en-CA', {
      timeZone: this.displayTimeZone,
      year: 'numeric', month: '2-digit', day: '2-digit'
    }).formatToParts(new Date());
    const value: Record<string, string> = {};
    parts.forEach((part: Intl.DateTimeFormatPart) => { value[part.type] = part.value; });
    return `${value.year}-${value.month}-${value.day}`;
  }

  private santiagoCurrentMinutes(): number {
    const parts: Intl.DateTimeFormatPart[] = new Intl.DateTimeFormat('en-GB', {
      timeZone: this.displayTimeZone,
      hour: '2-digit', minute: '2-digit', hourCycle: 'h23'
    }).formatToParts(new Date());
    const value: Record<string, string> = {};
    parts.forEach((part: Intl.DateTimeFormatPart) => { value[part.type] = part.value; });
    return Number(value.hour || 0) * 60 + Number(value.minute || 0);
  }

  private timeMinutes(value?: string): number {
    if (!value) return 0;
    const match: RegExpMatchArray | null = value.match(/T(\d{2}):(\d{2})/);
    return match ? Number(match[1]) * 60 + Number(match[2]) : 0;
  }

  private formatGraphTime(value?: string): string {
    const minutes: number = this.timeMinutes(value);
    return this.minutesLabel(minutes);
  }

  private minutesLabel(minutes: number): string {
    const safe: number = Math.max(0, Math.min(24 * 60, minutes));
    return `${String(Math.floor(safe / 60)).padStart(2, '0')}:${String(safe % 60).padStart(2, '0')}`;
  }

  private formatDateKey(dateKey: string): string {
    const parts: string[] = dateKey.split('-');
    if (parts.length !== 3) return dateKey;
    const date: Date = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
    return date.toLocaleDateString('es-CL', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });
  }

  private normalize(value: string): string {
    return value.trim().toLocaleLowerCase('es-CL').normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  }

  private odata(value: string): string {
    return value.replace(/'/g, "''");
  }

  private escape(value: string): string {
    return value.replace(/[&<>'"]/g, (char: string) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' }[char] || char));
  }

  private delay(ms: number): Promise<void> {
    return new Promise<void>((resolve: () => void) => window.setTimeout(resolve, ms));
  }
}
