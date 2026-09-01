import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import TiboxHubGraphWebPart from './TiboxHubGraphWebPart';
import { TIBOX_HUB_DESIGN_CSS } from './TiboxHubDesign';

type CalendarEventItem = {
  Title?: string;
  FechaInicio?: string;
};

/**
 * Presentation adapter for the approved TIBOX_HUB_1.html design.
 *
 * The existing Web Part remains responsible for SharePoint Lists, permissions,
 * settings and Microsoft Graph. This class maps the existing DOM to the final
 * Tibox design-system class names and adds the mini calendar used by the
 * approved prototype.
 */
export default class TiboxHubFinalWebPart extends TiboxHubGraphWebPart {
  private finalRenderRun: number = 0;

  public render(): void {
    const run: number = ++this.finalRenderRun;
    super.render();
    void this.applyFinalExperience(run);
  }

  private async applyFinalExperience(run: number): Promise<void> {
    const hub: HTMLElement | undefined = await this.waitForFinalHub(run);
    if (!hub || run !== this.finalRenderRun) return;

    this.injectApprovedDesign(hub);
    this.mapLegacyMarkupToApprovedClasses(hub);
    this.decorateHero(hub);
    await this.ensureMiniCalendar(hub, run);
    this.orderContextRail(hub);

    // Graph enhancements are asynchronous. Recheck the rail/photo after Graph
    // has had time to add the room card and the current-user photo.
    [500, 1200, 2400].forEach((delayMs: number): void => {
      window.setTimeout((): void => {
        if (run !== this.finalRenderRun) return;
        this.orderContextRail(hub);
        this.syncCurrentUserPhoto(hub);
      }, delayMs);
    });
  }

  private async waitForFinalHub(run: number): Promise<HTMLElement | undefined> {
    for (let attempt: number = 0; attempt < 100; attempt++) {
      if (run !== this.finalRenderRun) return undefined;
      const hub: HTMLElement | null = this.domElement.querySelector<HTMLElement>('.tbxHub, .tbx-hub');
      if (hub) return hub;
      await this.delay(80);
    }
    return undefined;
  }

  private injectApprovedDesign(hub: HTMLElement): void {
    if (hub.querySelector('[data-tbx-approved-design]')) return;

    const style: HTMLStyleElement = document.createElement('style');
    style.setAttribute('data-tbx-approved-design', 'true');

    // TiboxHubDesign was generated from the approved HTML. The normalization
    // supports the escaped representation committed in the source file.
    const approvedCss: string = TIBOX_HUB_DESIGN_CSS
      .replace(/\\n/g, '\n')
      .replace(/\\"/g, '"');

    style.textContent = `
      @import url('https://fonts.googleapis.com/css2?family=Titillium+Web:wght@400;600;700&family=IBM+Plex+Mono:wght@500;600&display=swap');
      ${approvedCss}

      /* Compatibility rules: current connected SPFx markup -> approved UI. */
      .tbx-hub.tbxHub{min-height:100%;border-radius:0;overflow:visible}
      .tbx-hub .tbxShell{max-width:1320px;margin:0 auto;padding:0 28px 64px;background:var(--tbx-bg)}
      .tbx-hub .tbxTop.tbx-header{height:auto;min-height:62px;margin:0 -28px;padding:12px 28px;gap:24px;background:rgba(255,255,255,.88);color:#001233;border-bottom:1px solid var(--tbx-border);backdrop-filter:blur(14px);-webkit-backdrop-filter:blur(14px)}
      .tbx-hub .tbxTop .tbxBrand{color:#001233;font-weight:700;letter-spacing:-.15px}
      .tbx-hub .tbxTop .tbxBrand small{color:#7A8199;font-weight:600}
      .tbx-hub .tbxTop .tbxMark{width:34px;height:34px;border-radius:10px;background:#000310;box-shadow:none;position:relative}
      .tbx-hub .tbxTop .tbxMark:after{content:'';position:absolute;inset:9px;border-radius:4px;background:linear-gradient(135deg,#00D1FF,#0E9CDC)}
      .tbx-hub .tbxTopActions{display:flex;align-items:center;gap:8px}

      .tbx-hub .tbxHero.tbx-hero{min-height:218px;margin:26px 0;padding:30px 34px;border-radius:26px;background:#000310;color:#F4F7FF;position:relative;overflow:hidden}
      .tbx-hub .tbxHero.tbx-hero:after{content:'';position:absolute;top:-70px;right:-60px;width:260px;height:260px;border-radius:50%;background:radial-gradient(circle,rgba(0,209,255,.20),transparent 70%);pointer-events:none}
      .tbx-hub .tbxHero.tbx-hero:before{content:'PANEL DEL DÍA';display:block;position:relative;z-index:2;margin-bottom:10px;font-family:'IBM Plex Mono',ui-monospace,monospace;font-size:10.5px;font-weight:600;letter-spacing:.18em;color:#00D1FF}
      .tbx-hub .tbxHero .tbx-hero-title{position:relative;z-index:2;color:#F4F7FF}
      .tbx-hub .tbxHero .tbx-hero-sub{position:relative;z-index:2;color:#9BA6C4}
      .tbx-hub .tbxHero .tbx-hero-search{position:relative;z-index:2}
      .tbx-hub .tbxHero .tbx-hero-search input{color:#F4F7FF}
      .tbx-hub .tbxHero .tbx-hero-search input::placeholder{color:#5E6A8A}
      .tbxFinalHeroUser{position:absolute;z-index:3;top:30px;right:34px;display:flex;align-items:center;gap:11px;padding:9px 14px 9px 9px;border:1px solid rgba(255,255,255,.08);border-radius:16px;background:rgba(255,255,255,.04);color:#F4F7FF}
      .tbxFinalHeroAvatar{width:44px;height:44px;display:grid;place-items:center;border-radius:50%;overflow:hidden;background:rgba(0,209,255,.15);color:#00D1FF;font-weight:700}
      .tbxFinalHeroAvatar img{width:100%;height:100%;object-fit:cover}
      .tbxFinalHeroName{font-size:13px;font-weight:700;line-height:1.2}.tbxFinalHeroRole{margin-top:3px;color:#9BA6C4;font-size:10.5px}
      .tbxFinalPulse{position:relative;z-index:2;display:flex;flex-wrap:wrap;gap:9px;margin-top:16px}
      .tbxFinalPulse span{display:flex;align-items:center;gap:6px;padding:6px 11px;border:1px solid rgba(255,255,255,.08);border-radius:20px;background:rgba(255,255,255,.05);color:#9BA6C4;font-size:11px;font-weight:600}.tbxFinalPulse b{color:#F4F7FF}

      .tbx-hub .tbxQuick.tbx-quicklinks{display:flex;gap:12px;overflow-x:auto;margin:0 0 6px;padding:4px 2px 10px}
      .tbx-hub .tbxQuick a.tbx-quicklink{min-height:0;flex:0 0 auto;display:flex;flex-direction:row;align-items:center;justify-content:flex-start;gap:10px;padding:11px 16px;background:#fff;color:#001233}

      .tbx-hub .tbxLayout.tbx-grid{grid-template-columns:minmax(0,1fr) 308px;gap:18px;align-items:start;margin-top:22px}
      .tbx-hub .tbxMain.tbx-col{gap:22px}.tbx-hub .tbxSide.tbx-rail{gap:14px;position:sticky;top:78px;align-self:start}
      .tbx-hub .tbxCard.tbx-card{padding:22px;border-radius:22px;overflow:hidden}
      .tbx-hub .tbxCardHead.tbx-card-head{min-height:0;padding:0;margin-bottom:16px;border:0}
      .tbx-hub .tbxSide .tbxCard.tbx-card{padding:16px;border-radius:18px}.tbx-hub .tbxSide .tbxCardHead.tbx-card-head{margin-bottom:12px}

      .tbx-hub .tbxPeopleGrid.tbx-people-grid{padding:0}.tbx-hub .tbxPerson.tbx-person-row{background:var(--tbx-surface-2)}
      .tbx-hub .tbxMovements{padding:0;display:flex;flex-direction:column;gap:9px}.tbx-hub .tbxMovement.tbx-movement-row{border-bottom:0}
      .tbx-hub .tbxCourses.tbx-courses-grid{padding:0}.tbx-hub .tbxCourse.tbx-course-card{background:var(--tbx-surface-2)}
      .tbx-hub .tbxCourse h3.tbx-course-name{margin:9px 0 4px}.tbx-hub .tbxCourseAction.tbx-course-cta{text-decoration:none}
      .tbx-hub .tbxCourseFooter{margin:15px 0 0;padding:15px 0 0}

      .tbx-hub .tbxLearning{padding:0;display:grid;grid-template-columns:1fr 1fr;gap:12px}.tbx-hub .tbxPrompt.tbx-prompt-block{margin:0;min-height:160px;display:block}.tbx-hub .tbxPrompt .tbx-prompt-text{margin:14px 0}.tbx-hub .tbxLearningList{display:flex;flex-direction:column;gap:8px}.tbx-hub .tbxLearningItem{padding:11px 12px;border:1px solid var(--tbx-border);border-radius:13px;background:var(--tbx-surface-2)}

      .tbx-hub .tbxNews{padding:0;display:grid;grid-template-columns:1fr 1fr;gap:9px}.tbx-hub .tbxNewsItem{min-height:90px;padding:13px;border:1px solid var(--tbx-border);border-radius:13px;background:var(--tbx-surface-2)}.tbx-hub .tbxNewsItem label{font-size:10px;font-weight:700;color:var(--tbx-support)}.tbx-hub .tbxNewsItem strong{font-size:13px;color:var(--tbx-text)}.tbx-hub .tbxNewsItem small{font-size:11px;color:var(--tbx-text-subtle)}

      .tbx-hub .tbxBirthdayList,.tbx-hub .tbxEventList{padding:0;gap:7px}.tbx-hub .tbxBirthday.tbx-bday-row,.tbx-hub .tbxEvent.tbx-event-row{padding:8px;border-radius:12px;background:var(--tbx-surface-2)}
      .tbx-hub .tbxBirthday button.tbx-bday-cta{border:0;background:rgba(255,178,0,.22);color:#B57900;padding:7px 10px;font-size:11px}
      .tbx-hub .tbxDate.tbx-event-date-chip{width:44px;height:44px;background:#000310;color:#fff}.tbx-hub .tbxDate.tbx-event-date-chip small{color:#00D1FF}
      .tbx-hub .tbxBenefitHero.tbx-benefits-teaser{min-height:150px;padding:18px;border-radius:18px;background:radial-gradient(circle at 100% 0,rgba(0,209,255,.18),transparent 150px),#0A1130}.tbx-hub .tbxBenefitHero .tbx-benefits-copy{color:rgba(255,255,255,.65)}

      .tbx-hub .tbxGraphRoom{border-color:var(--tbx-border);background:#fff;color:#001233;box-shadow:0 1px 2px rgba(0,18,51,.03)}.tbx-hub .tbxGraphRoomTitle,.tbx-hub .tbxGraphRoomStatus{color:#001233}.tbx-hub .tbxGraphRoomSub{color:#66686C}.tbx-hub .tbxGraphRoomNext{background:#F0F2F7;border-color:var(--tbx-border)}.tbx-hub .tbxGraphRoomNext strong{color:#001233}

      /* Existing settings/modal markup is retained but uses the approved light modal language. */
      .tbx-hub .tbxOverlay{background:rgba(0,3,16,.55);backdrop-filter:blur(3px)}.tbx-hub .tbxModal{border:0;border-radius:26px;background:#F7F7F7;color:#001233;box-shadow:0 30px 80px rgba(0,3,16,.4)}.tbx-hub .tbxModalHead{background:#F7F7F7;border-bottom:1px solid rgba(0,18,51,.07)}.tbx-hub .tbxModalHead h3{color:#001233}.tbx-hub .tbxClose{background:rgba(0,18,51,.06);border:0;color:#5C6478}.tbx-hub .tbxTextarea,.tbx-hub .tbxModalSearch,.tbx-hub .tbxSettingRow input[type=text]{background:#fff;color:#001233;border-color:var(--tbx-border)}.tbx-hub .tbxSettingRow{background:#fff;border-color:var(--tbx-border)}.tbx-hub .tbxAdminNote{background:rgba(14,156,220,.07);border-color:rgba(14,156,220,.20);color:#3C4253}.tbx-hub .tbxAdminNote strong{color:#001233}.tbx-hub .tbxBtn{background:#fff;color:#001233;border-color:var(--tbx-border)}.tbx-hub .tbxBtn.primary{background:var(--tbx-cta-gradient);color:#fff;border-color:#FF4222}

      @media(max-width:1180px){.tbx-hub .tbxLayout.tbx-grid{grid-template-columns:1fr}.tbx-hub .tbxSide.tbx-rail{position:static;top:auto;gap:22px}.tbx-hub .tbxCourses.tbx-courses-grid{grid-template-columns:repeat(2,1fr)}}
      @media(max-width:760px){.tbx-hub .tbxShell{padding:0 16px 40px}.tbx-hub .tbxTop.tbx-header{margin:0 -16px;padding:10px 16px}.tbx-hub .tbxHero.tbx-hero{padding:24px 22px;min-height:0}.tbxFinalHeroUser{position:static;margin-top:16px;width:max-content}.tbx-hub .tbxPeopleGrid.tbx-people-grid,.tbx-hub .tbxCourses.tbx-courses-grid,.tbx-hub .tbxLearning,.tbx-hub .tbxNews{grid-template-columns:1fr}}
    `;
    hub.appendChild(style);
  }

  private mapLegacyMarkupToApprovedClasses(hub: HTMLElement): void {
    hub.classList.add('tbx-hub');
    hub.setAttribute('data-tbx-theme', 'light');

    this.addClasses(hub, '.tbxShell', ['tbx-shell']);
    this.addClasses(hub, '.tbxTop', ['tbx-header']);
    this.addClasses(hub, '.tbxBrand', ['tbx-brand']);
    this.addClasses(hub, '.tbxMark', ['tbx-brand-mark']);
    this.addClasses(hub, '.tbxNav', ['tbx-nav']);
    this.addClasses(hub, '.tbxNav button', ['tbx-nav-btn']);
    hub.querySelectorAll<HTMLElement>('.tbxNav button.active').forEach((button: HTMLElement) => button.classList.add('is-active'));
    this.addClasses(hub, '.tbxIconBtn', ['tbx-icon-btn']);

    this.addClasses(hub, '.tbxHero', ['tbx-hero']);
    this.addClasses(hub, '.tbxHero h1', ['tbx-hero-title']);
    this.addClasses(hub, '.tbxHero p', ['tbx-hero-sub']);
    this.addClasses(hub, '.tbxSearch', ['tbx-hero-search']);

    this.addClasses(hub, '.tbxQuick', ['tbx-quicklinks']);
    this.addClasses(hub, '.tbxQuick a', ['tbx-quicklink']);
    this.addClasses(hub, '.tbxQuickIcon', ['tbx-quicklink-icon']);
    hub.querySelectorAll<HTMLAnchorElement>('.tbxQuick a').forEach((link: HTMLAnchorElement) => {
      const spans: NodeListOf<HTMLSpanElement> = link.querySelectorAll<HTMLSpanElement>('span');
      if (spans.length > 1) spans[spans.length - 1].classList.add('tbx-quicklink-label');
    });

    this.addClasses(hub, '.tbxLayout', ['tbx-grid']);
    this.addClasses(hub, '.tbxMain', ['tbx-col']);
    this.addClasses(hub, '.tbxSide', ['tbx-col', 'tbx-rail']);
    this.addClasses(hub, '.tbxCard', ['tbx-card']);
    this.addClasses(hub, '.tbxCardHead', ['tbx-card-head']);
    this.addClasses(hub, '.tbxCardHead h2,.tbxCardHead h3', ['tbx-card-title']);
    this.addClasses(hub, '.tbxLink', ['tbx-link-btn']);

    this.addClasses(hub, '.tbxPeopleGrid', ['tbx-people-grid']);
    this.addClasses(hub, '.tbxPerson', ['tbx-person-row']);
    this.addClasses(hub, '.tbxAvatar', ['tbx-avatar', 'tbx-avatar-40']);
    this.addClasses(hub, '.tbxPersonText strong', ['tbx-person-name']);
    this.addClasses(hub, '.tbxPersonText span', ['tbx-person-meta']);
    this.addClasses(hub, '.tbxPersonText small', ['tbx-person-joined']);

    this.addClasses(hub, '.tbxMovement', ['tbx-movement-row']);
    hub.querySelectorAll<HTMLElement>('.tbxMovement .change').forEach((badge: HTMLElement) => {
      badge.classList.add('tbx-badge-move');
      badge.classList.add((badge.textContent || '').toLowerCase().includes('ascenso') ? 'ascenso' : 'cambio');
    });

    this.addClasses(hub, '.tbxCourses', ['tbx-courses-grid']);
    this.addClasses(hub, '.tbxCourse', ['tbx-course-card']);
    this.addClasses(hub, '.tbxCourseTop', ['tbx-course-top']);
    this.addClasses(hub, '.tbxCourse h3', ['tbx-course-name']);
    this.addClasses(hub, '.tbxCourseMeta', ['tbx-course-meta']);
    this.addClasses(hub, '.tbxCourseAction', ['tbx-course-cta']);

    this.addClasses(hub, '.tbxPrompt', ['tbx-prompt-block']);
    this.addClasses(hub, '.tbxPrompt label', ['tbx-prompt-eyebrow']);
    this.addClasses(hub, '.tbxPrompt p', ['tbx-prompt-text']);
    this.addClasses(hub, '.tbxPrompt button', ['tbx-prompt-cta']);

    this.addClasses(hub, '.tbxBirthday', ['tbx-bday-row']);
    this.addClasses(hub, '.tbxBirthday button', ['tbx-bday-cta']);
    this.addClasses(hub, '.tbxEvent', ['tbx-event-row']);
    this.addClasses(hub, '.tbxDate', ['tbx-event-date-chip']);
    this.addClasses(hub, '.tbxBenefitHero', ['tbx-benefits-teaser']);
    this.addClasses(hub, '.tbxBenefitHero label', ['tbx-benefits-eyebrow']);
    this.addClasses(hub, '.tbxBenefitHero h3', ['tbx-benefits-title']);
    this.addClasses(hub, '.tbxBenefitHero p', ['tbx-benefits-copy']);
    this.addClasses(hub, '.tbxBenefitHero button', ['tbx-benefits-cta']);
  }

  private addClasses(root: HTMLElement, selector: string, classes: string[]): void {
    root.querySelectorAll<HTMLElement>(selector).forEach((element: HTMLElement) => element.classList.add(...classes));
  }

  private decorateHero(hub: HTMLElement): void {
    const hero: HTMLElement | null = hub.querySelector<HTMLElement>('.tbxHero');
    if (!hero) return;

    const displayName: string = this.context.pageContext.user.displayName || 'Equipo Tibox';
    const firstName: string = displayName.split(/\s+/)[0] || displayName;
    const title: HTMLHeadingElement | null = hero.querySelector<HTMLHeadingElement>('h1');
    if (title) title.textContent = `Hola, ${firstName}`;

    if (!hero.querySelector('[data-tbx-final-hero-user]')) {
      const user: HTMLDivElement = document.createElement('div');
      user.className = 'tbxFinalHeroUser';
      user.setAttribute('data-tbx-final-hero-user', 'true');
      user.innerHTML = `<span class="tbxFinalHeroAvatar" data-tbx-current-avatar-final>${this.initials(displayName)}</span><span><span class="tbxFinalHeroName" style="display:block">${this.escape(displayName)}</span><span class="tbxFinalHeroRole" style="display:block">Microsoft 365 · Tibox</span></span>`;
      hero.appendChild(user);
    }

    if (!hero.querySelector('[data-tbx-final-pulse]')) {
      const pulse: HTMLDivElement = document.createElement('div');
      pulse.className = 'tbxFinalPulse';
      pulse.setAttribute('data-tbx-final-pulse', 'true');
      const hires: number = hub.querySelectorAll('.tbxPerson').length;
      const birthdays: number = hub.querySelectorAll('.tbxBirthday').length;
      const events: number = hub.querySelectorAll('.tbxEvent').length;
      pulse.innerHTML = `<span><b>${hires}</b> nuevos ingresos</span><span><b>${birthdays}</b> próximos cumpleaños</span><span><b>${events}</b> eventos próximos</span>`;
      hero.appendChild(pulse);
    }
  }

  private syncCurrentUserPhoto(hub: HTMLElement): void {
    const source: HTMLImageElement | null = hub.querySelector<HTMLImageElement>('.tbxGraphMe img');
    const target: HTMLElement | null = hub.querySelector<HTMLElement>('[data-tbx-current-avatar-final]');
    if (!source || !target || target.querySelector('img')) return;
    const image: HTMLImageElement = document.createElement('img');
    image.src = source.src;
    image.alt = `Foto de ${this.context.pageContext.user.displayName || 'usuario'}`;
    target.innerHTML = '';
    target.appendChild(image);
  }

  private async ensureMiniCalendar(hub: HTMLElement, run: number): Promise<void> {
    const side: HTMLElement | null = hub.querySelector<HTMLElement>('.tbxSide');
    if (!side || side.querySelector('[data-tbx-final-calendar]')) return;

    const events: CalendarEventItem[] = await this.loadCalendarEvents();
    if (run !== this.finalRenderRun) return;

    const now: Date = new Date();
    const year: number = now.getFullYear();
    const month: number = now.getMonth();
    const eventDays: Set<number> = new Set<number>();
    events.forEach((item: CalendarEventItem) => {
      if (!item.FechaInicio) return;
      const date: Date = new Date(item.FechaInicio);
      if (date.getFullYear() === year && date.getMonth() === month) eventDays.add(date.getDate());
    });

    const first: Date = new Date(year, month, 1);
    let offset: number = first.getDay() - 1;
    if (offset < 0) offset = 6;
    const daysInMonth: number = new Date(year, month + 1, 0).getDate();
    const cells: Array<number | undefined> = [];
    for (let index: number = 0; index < offset; index++) cells.push(undefined);
    for (let day: number = 1; day <= daysInMonth; day++) cells.push(day);

    const card: HTMLElement = document.createElement('section');
    card.className = 'tbx-card tbx-mini-cal';
    card.setAttribute('data-tbx-final-calendar', 'true');
    const monthLabel: string = now.toLocaleDateString('es-CL', { month: 'long', year: 'numeric' });
    card.innerHTML = `<div class="tbx-mini-cal-head"><span class="tbx-mini-cal-month">${this.escape(monthLabel)}</span>${this.calendarSvg()}</div><div class="tbx-mini-cal-grid">${['L','M','X','J','V','S','D'].map((day: string) => `<span class="tbx-mini-cal-dow">${day}</span>`).join('')}${cells.map((day: number | undefined) => {
      if (!day) return '<span class="tbx-mini-cal-day is-blank">·</span>';
      const todayClass: string = day === now.getDate() ? ' is-today' : '';
      const dot: string = eventDays.has(day) ? '<span class="tbx-mini-cal-dot"></span>' : '';
      return `<span class="tbx-mini-cal-day${todayClass}">${day}${dot}</span>`;
    }).join('')}</div><div class="tbx-mini-cal-legend"><span class="tbx-mini-cal-dot-static"></span>Días con eventos agendados</div>`;
    side.insertBefore(card, side.firstElementChild);
  }

  private async loadCalendarEvents(): Promise<CalendarEventItem[]> {
    try {
      const webUrl: string = this.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
      const title: string = 'TIBOX HUB - Eventos';
      const endpoint: string = `${webUrl}/_api/web/lists/getbytitle('${this.odata(title)}')/items?$select=Title,FechaInicio,Activo&$filter=Activo eq 1&$orderby=FechaInicio asc&$top=100`;
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1, { headers: { Accept: 'application/json;odata=nometadata' } });
      if (!response.ok) return [];
      const payload: { value?: CalendarEventItem[] } = await response.json() as { value?: CalendarEventItem[] };
      return payload.value || [];
    } catch {
      return [];
    }
  }

  private orderContextRail(hub: HTMLElement): void {
    const side: HTMLElement | null = hub.querySelector<HTMLElement>('.tbxSide');
    if (!side) return;
    const calendar: Element | null = side.querySelector('[data-tbx-final-calendar]');
    const room: Element | null = side.querySelector('[data-tbx-room-card]');
    if (calendar && side.firstElementChild !== calendar) side.insertBefore(calendar, side.firstElementChild);
    if (calendar && room && calendar.nextElementSibling !== room) side.insertBefore(room, calendar.nextElementSibling);
  }

  private calendarSvg(): string {
    return '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="3.5" y="5" width="17" height="16" rx="3"/><path d="M8 3v4M16 3v4M3.5 10h17"/></svg>';
  }

  private initials(name: string): string {
    return name.split(/\s+/).filter(Boolean).slice(0, 2).map((part: string) => part.charAt(0).toUpperCase()).join('');
  }

  private escape(value: string): string {
    return value.replace(/[&<>'\"]/g, (char: string) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '\"': '&quot;' }[char] || char));
  }

  private odata(value: string): string {
    return value.replace(/'/g, "''");
  }

  private delay(ms: number): Promise<void> {
    return new Promise<void>((resolve: () => void) => window.setTimeout(resolve, ms));
  }
}
