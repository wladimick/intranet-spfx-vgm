import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import TiboxHubGraphWebPart from './TiboxHubGraphWebPart';

type CalendarEventItem = {
  Title?: string;
  FechaInicio?: string;
};

/**
 * Final presentation layer for TIBOX HUB.
 *
 * Keeps the connected SharePoint + Microsoft Graph implementation intact and
 * maps the visual language approved in TIBOX_HUB_1.html onto the SPFx markup.
 */
export default class TiboxHubFinalWebPart extends TiboxHubGraphWebPart {
  private finalRenderRun: number = 0;

  public render(): void {
    const run: number = ++this.finalRenderRun;
    super.render();
    void this.applyFinalExperience(run);
  }

  private async applyFinalExperience(run: number): Promise<void> {
    const hub: HTMLElement | undefined = await this.waitForHub(run);
    if (!hub || run !== this.finalRenderRun) return;

    hub.classList.add('tbxFinal');
    this.injectFinalStyles(hub);
    this.decorateHero(hub);
    await this.ensureMiniCalendar(hub, run);
    this.orderContextRail(hub);

    // Graph enhancements (room/photo) are asynchronous. Re-apply the desired
    // rail order after they have had time to attach.
    window.setTimeout((): void => {
      if (run === this.finalRenderRun) this.orderContextRail(hub);
    }, 900);
    window.setTimeout((): void => {
      if (run === this.finalRenderRun) this.orderContextRail(hub);
    }, 2200);
  }

  private async waitForHub(run: number): Promise<HTMLElement | undefined> {
    for (let attempt: number = 0; attempt < 100; attempt++) {
      if (run !== this.finalRenderRun) return undefined;
      const hub: HTMLElement | null = this.domElement.querySelector<HTMLElement>('.tbxHub, .tbx-hub');
      if (hub) return hub;
      await this.delay(80);
    }
    return undefined;
  }

  private injectFinalStyles(hub: HTMLElement): void {
    if (hub.querySelector('[data-tbx-final-styles]')) return;

    const style: HTMLStyleElement = document.createElement('style');
    style.setAttribute('data-tbx-final-styles', 'true');
    style.textContent = `
      @import url('https://fonts.googleapis.com/css2?family=Titillium+Web:wght@400;600;700&family=IBM+Plex+Mono:wght@500;600&display=swap');

      .tbxHub.tbxFinal,.tbxHub.tbxFinal *{box-sizing:border-box}
      .tbxHub.tbxFinal{
        --bg:#F7F7F7;--surface:#FFFFFF;--surface2:#F0F2F7;
        --border:rgba(0,18,51,.09);--borderAccent:rgba(14,156,220,.45);
        --text:#001233;--muted:#3C4253;--subtle:#66686C;
        --support:#0E9CDC;--supportSolid:#0E9CDC;--cta:#FF4222;
        --yellow:#B57900;--green:#1F8A55;--red:#D6371F;--violet:#6D5BD0;
        width:100%;min-height:100%;background:#F7F7F7;color:#3C4253;
        border-radius:0;overflow:visible;font-family:'Titillium Web',-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;
        font-size:14px;line-height:1.4;-webkit-font-smoothing:antialiased
      }
      .tbxFinal button,.tbxFinal input,.tbxFinal textarea{font-family:inherit}
      .tbxFinal .tbxShell{max-width:1320px;margin:0 auto;padding:0 28px 64px}

      /* Header final: liviano y claro, como el prototipo aprobado. */
      .tbxFinal .tbxTop{
        height:64px;margin:0 -28px;padding:0 28px;position:sticky;top:0;z-index:40;
        background:rgba(255,255,255,.88);backdrop-filter:blur(14px);-webkit-backdrop-filter:blur(14px);
        border-bottom:1px solid rgba(0,18,51,.08);color:#001233
      }
      .tbxFinal .tbxBrand{color:#001233;font-weight:700;letter-spacing:-.1px}
      .tbxFinal .tbxBrand small{color:#7A8199;font-weight:600}
      .tbxFinal .tbxMark{width:34px;height:34px;border-radius:10px;background:#000310;box-shadow:none;position:relative}
      .tbxFinal .tbxMark:after{content:'';position:absolute;inset:9px;border-radius:4px;background:linear-gradient(135deg,#00D1FF,#0E9CDC)}
      .tbxFinal .tbxNav{gap:2px;padding:4px;background:#fff;border:1px solid rgba(0,18,51,.07);border-radius:12px;box-shadow:0 1px 2px rgba(0,18,51,.04)}
      .tbxFinal .tbxNav button{padding:8px 15px;border-radius:9px;color:#5C6478;font-weight:700;font-size:12.5px}
      .tbxFinal .tbxNav button:hover{background:#F3F1FF;color:#001233}
      .tbxFinal .tbxNav button.active{background:#001233;color:#fff}
      .tbxFinal .tbxTopActions{align-items:center;gap:8px}
      .tbxFinal .tbxIconBtn{width:38px;height:38px;border-radius:10px;border:1px solid rgba(0,18,51,.08);background:#fff;color:#5C6478}
      .tbxFinal .tbxIconBtn:hover{color:#001233;border-color:rgba(0,18,51,.18);background:#fff}
      .tbxFinal .tbxGraphMe{border-color:rgba(0,18,51,.08);background:#fff;color:#001233}

      /* Hero oscuro dentro de un portal claro. */
      .tbxFinal .tbxHero{
        margin:26px 0 26px;padding:30px 34px;border-radius:26px;position:relative;overflow:hidden;
        background:#000310;color:#F4F7FF;min-height:218px
      }
      .tbxFinal .tbxHero:after{content:'';position:absolute;top:-70px;right:-55px;width:260px;height:260px;border-radius:50%;background:radial-gradient(circle,rgba(0,209,255,.20),transparent 70%);pointer-events:none}
      .tbxFinal .tbxHero:before{content:'PANEL DEL DÍA';display:block;position:relative;z-index:2;margin-bottom:10px;font-family:'IBM Plex Mono',ui-monospace,monospace;font-size:10.5px;font-weight:600;letter-spacing:.18em;color:#00D1FF}
      .tbxFinal .tbxHero h1{position:relative;z-index:2;margin:0;color:#F4F7FF;font-size:clamp(1.7rem,2.6vw,2.15rem);line-height:1.12;letter-spacing:-.5px;font-weight:700}
      .tbxFinal .tbxHero p{position:relative;z-index:2;margin:6px 0 0;color:#9BA6C4;font-size:14.5px;font-weight:500}
      .tbxFinal .tbxSearch{position:relative;z-index:2;margin-top:22px;width:min(100%,640px);height:50px;padding:0 16px;border-radius:15px;border:1px solid rgba(255,255,255,.08);background:rgba(255,255,255,.06)}
      .tbxFinal .tbxSearch input{color:#F4F7FF;font-size:14px;background:transparent}
      .tbxFinal .tbxSearch input::placeholder{color:#5E6A8A}
      .tbxFinalHeroUser{position:absolute;z-index:3;top:30px;right:34px;display:flex;align-items:center;gap:11px;padding:9px 14px 9px 9px;border-radius:16px;border:1px solid rgba(255,255,255,.08);background:rgba(255,255,255,.04);color:#F4F7FF}
      .tbxFinalHeroAvatar{width:44px;height:44px;border-radius:50%;display:grid;place-items:center;background:rgba(0,209,255,.16);color:#00D1FF;font-weight:700;overflow:hidden}
      .tbxFinalHeroAvatar img{width:100%;height:100%;object-fit:cover}
      .tbxFinalHeroName{font-size:13px;font-weight:700;line-height:1.2}
      .tbxFinalHeroRole{margin-top:3px;color:#9BA6C4;font-size:10.5px}
      .tbxFinalPulse{position:relative;z-index:2;display:flex;flex-wrap:wrap;gap:9px;margin-top:16px}
      .tbxFinalPulse span{display:flex;align-items:center;gap:6px;padding:6px 11px;border:1px solid rgba(255,255,255,.08);border-radius:20px;background:rgba(255,255,255,.05);color:#9BA6C4;font-size:11px;font-weight:600}
      .tbxFinalPulse b{color:#F4F7FF}

      /* Accesos rápidos horizontales. */
      .tbxFinal .tbxQuick{display:flex;grid-template-columns:none;gap:12px;overflow-x:auto;margin:0 0 6px;padding:4px 2px 10px}
      .tbxFinal .tbxQuick a{flex:0 0 auto;min-height:0;padding:11px 16px;display:flex;flex-direction:row;align-items:center;justify-content:flex-start;gap:10px;border:1px solid rgba(0,18,51,.09);border-radius:14px;background:#fff;color:#001233;font-size:13px;font-weight:700;box-shadow:0 1px 2px rgba(0,18,51,.03)}
      .tbxFinal .tbxQuick a:hover{transform:translateY(-2px);border-color:rgba(14,156,220,.45);background:#fff;box-shadow:0 10px 24px rgba(0,18,51,.08)}
      .tbxFinal .tbxQuickIcon{width:32px;height:32px;border-radius:9px;background:rgba(14,156,220,.10);color:#0E9CDC}

      /* Grid final y context rail de 308px. */
      .tbxFinal .tbxLayout{display:grid;grid-template-columns:minmax(0,1fr) 308px;gap:18px;align-items:start;margin-top:22px}
      .tbxFinal .tbxMain{gap:22px;min-width:0}
      .tbxFinal .tbxSide{gap:14px;position:sticky;top:78px;align-self:start}
      .tbxFinal .tbxCard{padding:22px;border:1px solid rgba(0,18,51,.09);border-radius:22px;background:#fff;box-shadow:0 1px 2px rgba(0,18,51,.03);overflow:hidden;color:#001233}
      .tbxFinal .tbxCardHead{min-height:0;padding:0;margin-bottom:16px;border:0;background:transparent}
      .tbxFinal .tbxCardHead h2,.tbxFinal .tbxCardHead h3{color:#001233;font-size:15.5px;font-weight:700}
      .tbxFinal .tbxLink{color:#0E9CDC;font-size:12px;font-weight:700}
      .tbxFinal .tbxSide>.tbxCard{padding:16px;border-radius:18px}
      .tbxFinal .tbxSide>.tbxCard .tbxCardHead{margin-bottom:12px}
      .tbxFinal .tbxSide>.tbxCard .tbxCardHead h2,.tbxFinal .tbxSide>.tbxCard .tbxCardHead h3{font-size:13.5px}

      /* Personas / movimientos. */
      .tbxFinal .tbxPeopleGrid{padding:0;display:grid;grid-template-columns:1fr 1fr;gap:11px}
      .tbxFinal .tbxPerson{padding:11px;gap:12px;border:1px solid rgba(0,18,51,.09);border-radius:14px;background:#F0F2F7;color:#001233}
      .tbxFinal .tbxPerson:hover{border-color:rgba(14,156,220,.45)}
      .tbxFinal .tbxAvatar{width:40px;height:40px;color:#0E9CDC;background:rgba(14,156,220,.18);font-size:12px}
      .tbxFinal .tbxPersonText strong{color:#001233;font-size:13.5px;font-weight:700}
      .tbxFinal .tbxPersonText span{color:#3C4253;font-size:11.5px}
      .tbxFinal .tbxPersonText small{color:#0E9CDC;font-size:11px;font-weight:700}
      .tbxFinal .tbxMovements{padding:0;display:flex;flex-direction:column;gap:9px}
      .tbxFinal .tbxMovement{padding:12px 13px;border:1px solid rgba(0,18,51,.09);border-radius:14px;background:#F0F2F7}
      .tbxFinal .tbxMovement .change{color:#1F8A55;background:rgba(31,138,85,.16);border:0;padding:5px 10px;font-size:10px}

      /* Cursos. */
      .tbxFinal .tbxCourses{padding:0;display:grid;grid-template-columns:repeat(3,1fr);gap:12px}
      .tbxFinal .tbxCourse{padding:15px;border:1px solid rgba(0,18,51,.09);border-radius:16px;background:#F0F2F7;color:#001233}
      .tbxFinal .tbxCourse h3{margin:11px 0 7px;color:#001233;font-size:13.5px;font-weight:700}
      .tbxFinal .tbxCourseMeta{color:#3C4253;font-size:11px;line-height:1.55}
      .tbxFinal .tbxCourseAction{margin-top:10px;padding:7px 13px;border-radius:9px;background:#0E9CDC;color:#fff;font-size:12px}
      .tbxFinal .tbxCourseFooter{margin:15px 0 0;padding-top:15px;border-top:1px solid rgba(0,18,51,.09)}
      .tbxFinal .tbxStatus{color:#3C4253;background:#F0F2F7;border:1px solid rgba(0,18,51,.06)}
      .tbxFinal .tbxPill.blue{color:#0E9CDC;background:rgba(14,156,220,.16)}
      .tbxFinal .tbxPill.warn{color:#D6371F;background:rgba(255,66,34,.14)}
      .tbxFinal .tbxPill.new{color:#6D5BD0;background:rgba(109,91,208,.16)}
      .tbxFinal .tbxPill.required{color:#D6371F;background:transparent}

      /* Aprende / prompt. */
      .tbxFinal .tbxLearning{padding:0;grid-template-columns:1fr 1fr;gap:12px}
      .tbxFinal .tbxPrompt{min-height:160px;padding:22px 24px;border-radius:20px;background:#000310;border:0;color:#F4F7FF}
      .tbxFinal .tbxPrompt label{color:#00D1FF;font-family:'IBM Plex Mono',ui-monospace,monospace;font-size:10px;letter-spacing:.14em}
      .tbxFinal .tbxPrompt p{color:#EAF0FF;font-size:13px;line-height:1.6}
      .tbxFinal .tbxLearningList{gap:8px}
      .tbxFinal .tbxLearningItem{padding:12px;border:1px solid rgba(0,18,51,.09);border-radius:13px;background:#F0F2F7;color:#001233}
      .tbxFinal .tbxLearningItem small{color:#66686C}

      /* Noticias. */
      .tbxFinal .tbxNews{padding:0;display:grid;grid-template-columns:1fr 1fr;gap:9px}
      .tbxFinal .tbxNewsItem{min-height:90px;padding:13px;border:1px solid rgba(0,18,51,.09);border-radius:13px;background:#F0F2F7}
      .tbxFinal .tbxNewsItem label{color:#0E9CDC;font-size:10px;font-weight:700}
      .tbxFinal .tbxNewsItem strong{color:#001233;font-size:13px;line-height:1.35}
      .tbxFinal .tbxNewsItem small{color:#66686C;font-size:11px}

      /* Rail: cumpleaños y eventos compactos. */
      .tbxFinal .tbxBirthdayList,.tbxFinal .tbxEventList{padding:0;gap:7px}
      .tbxFinal .tbxBirthday,.tbxFinal .tbxEvent{padding:8px;border:1px solid rgba(0,18,51,.09);border-radius:12px;background:#F0F2F7}
      .tbxFinal .tbxBirthday button{margin-left:auto;padding:7px 10px;border:0;border-radius:9px;background:rgba(255,178,0,.22);color:#B57900;font-size:11px;font-weight:700}
      .tbxFinal .tbxDate{width:44px;height:44px;border-radius:12px;background:#000310;color:#fff;font-size:13px}
      .tbxFinal .tbxDate small{color:#00D1FF;font-size:8px}
      .tbxFinal .tbxBenefitHero{min-height:150px;padding:18px;border-radius:18px;background:radial-gradient(circle at 100% 0,rgba(0,209,255,.18),transparent 150px),#0A1130;color:#fff}
      .tbxFinal .tbxBenefitHero h3{color:#fff;font-size:15.5px}
      .tbxFinal .tbxBenefitHero p{color:rgba(255,255,255,.65);font-size:12px}
      .tbxFinal .tbxBenefitHero button{padding:10px 16px;border-radius:10px;background:linear-gradient(0deg,#FF4222,#EA7E18);font-size:12px}

      /* Mini calendario añadido a partir de la lista de Eventos. */
      .tbxFinalMiniCal{padding:16px;border:1px solid rgba(0,18,51,.09);border-radius:18px;background:#fff;box-shadow:0 1px 2px rgba(0,18,51,.03)}
      .tbxFinalMiniCalHead{display:flex;align-items:center;justify-content:space-between;margin-bottom:12px;color:#001233}
      .tbxFinalMiniCalHead strong{font-size:13px;text-transform:capitalize}
      .tbxFinalMiniCalIcon{width:28px;height:28px;border-radius:8px;display:grid;place-items:center;color:#0E9CDC;background:rgba(14,156,220,.10)}
      .tbxFinalMiniCalGrid{display:grid;grid-template-columns:repeat(7,1fr);gap:3px;text-align:center}
      .tbxFinalMiniCalDow{font-size:9.5px;font-weight:700;color:#66686C;padding-bottom:4px}
      .tbxFinalMiniCalDay{position:relative;padding:5px 0;border-radius:8px;color:#3C4253;font-size:11px}
      .tbxFinalMiniCalDay.is-empty{visibility:hidden}
      .tbxFinalMiniCalDay.is-today{background:#0E9CDC;color:#fff;font-weight:700}
      .tbxFinalMiniCalDot{position:absolute;left:50%;bottom:1px;transform:translateX(-50%);width:4px;height:4px;border-radius:50%;background:#FF4222}
      .tbxFinalMiniCalDay.is-today .tbxFinalMiniCalDot{background:#fff}
      .tbxFinalMiniCalLegend{display:flex;align-items:center;gap:6px;margin-top:10px;color:#66686C;font-size:10.5px}
      .tbxFinalMiniCalLegendDot{width:5px;height:5px;border-radius:50%;background:#FF4222}

      /* Graph room inherits the context-rail visual language. */
      .tbxFinal .tbxGraphRoom{border-color:rgba(0,18,51,.09);border-radius:18px;background:#fff;color:#001233;box-shadow:0 1px 2px rgba(0,18,51,.03)}
      .tbxFinal .tbxGraphRoomHead{padding:12px 16px;border-bottom:1px solid rgba(0,18,51,.07)}
      .tbxFinal .tbxGraphRoomBody{padding:13px 16px 16px}
      .tbxFinal .tbxGraphRoomTitle,.tbxFinal .tbxGraphRoomStatus{color:#001233}
      .tbxFinal .tbxGraphRoomSub{color:#66686C}
      .tbxFinal .tbxGraphRoomNext{border-color:rgba(0,18,51,.08);background:#F0F2F7}
      .tbxFinal .tbxGraphRoomNext label{color:#66686C}.tbxFinal .tbxGraphRoomNext strong{color:#001233}
      .tbxFinal .tbxGraphRoomBtn{border-color:rgba(14,156,220,.25);background:rgba(14,156,220,.08);color:#0E9CDC}

      /* Modales de la implementación actual llevados al lenguaje final. */
      .tbxFinal .tbxOverlay{background:rgba(0,3,16,.55);backdrop-filter:blur(3px)}
      .tbxFinal .tbxModal{border:0;border-radius:26px;background:#F7F7F7;color:#001233;box-shadow:0 30px 80px rgba(0,3,16,.40)}
      .tbxFinal .tbxModalHead{background:#F7F7F7;border-bottom:1px solid rgba(0,18,51,.07)}
      .tbxFinal .tbxModalHead h3{color:#001233}
      .tbxFinal .tbxClose{background:rgba(0,18,51,.06);border:0;color:#5C6478}
      .tbxFinal .tbxModalSearch,.tbxFinal .tbxTextarea,.tbxFinal .tbxSettingRow input[type='text']{background:#fff;color:#001233;border-color:rgba(0,18,51,.09)}
      .tbxFinal .tbxSettingRow{background:#fff;border-color:rgba(0,18,51,.08)}
      .tbxFinal .tbxAdminNote{background:rgba(14,156,220,.07);border-color:rgba(14,156,220,.20);color:#3C4253}
      .tbxFinal .tbxAdminNote strong{color:#001233}
      .tbxFinal .tbxBtn{background:#fff;color:#001233;border-color:rgba(0,18,51,.09)}
      .tbxFinal .tbxBtn.primary{background:linear-gradient(0deg,#FF4222,#EA7E18);border-color:#FF4222;color:#fff}

      @media(max-width:1180px){
        .tbxFinal .tbxLayout{grid-template-columns:1fr}.tbxFinal .tbxSide{position:static;top:auto;gap:22px}
        .tbxFinal .tbxSide>.tbxCard{padding:22px;border-radius:22px}.tbxFinal .tbxCourses{grid-template-columns:repeat(2,1fr)}
      }
      @media(max-width:760px){
        .tbxFinal .tbxShell{padding:0 16px 40px}.tbxFinal .tbxTop{margin:0 -16px;padding:0 16px}.tbxFinal .tbxNav{display:none}
        .tbxFinal .tbxHero{padding:24px 22px;min-height:0}.tbxFinalHeroUser{position:static;margin-top:16px;width:max-content}
        .tbxFinal .tbxPeopleGrid,.tbxFinal .tbxCourses,.tbxFinal .tbxLearning,.tbxFinal .tbxNews{grid-template-columns:1fr}
      }
    `;
    hub.appendChild(style);
  }

  private decorateHero(hub: HTMLElement): void {
    const hero: HTMLElement | null = hub.querySelector<HTMLElement>('.tbxHero');
    if (!hero || hero.querySelector('[data-tbx-final-hero-user]')) return;

    const displayName: string = this.context.pageContext.user.displayName || 'Equipo Tibox';
    const firstName: string = displayName.split(/\s+/)[0] || displayName;
    const title: HTMLHeadingElement | null = hero.querySelector<HTMLHeadingElement>('h1');
    if (title) title.textContent = `Hola, ${firstName}`;

    const user: HTMLDivElement = document.createElement('div');
    user.className = 'tbxFinalHeroUser';
    user.setAttribute('data-tbx-final-hero-user', 'true');
    user.innerHTML = `
      <span class="tbxFinalHeroAvatar" data-tbx-current-avatar-final>${this.initials(displayName)}</span>
      <span><span class="tbxFinalHeroName" style="display:block">${this.escape(displayName)}</span><span class="tbxFinalHeroRole" style="display:block">Microsoft 365 · Tibox</span></span>`;
    hero.appendChild(user);

    const pulse: HTMLDivElement = document.createElement('div');
    pulse.className = 'tbxFinalPulse';
    pulse.setAttribute('data-tbx-final-pulse', 'true');
    const hires: number = hub.querySelectorAll('.tbxPerson').length;
    const birthdays: number = hub.querySelectorAll('.tbxBirthday').length;
    const events: number = hub.querySelectorAll('.tbxEvent').length;
    pulse.innerHTML = `
      <span><b>${hires}</b> nuevos ingresos</span>
      <span><b>${birthdays}</b> próximos cumpleaños</span>
      <span><b>${events}</b> eventos próximos</span>`;
    hero.appendChild(pulse);
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
    const totalDays: number = new Date(year, month + 1, 0).getDate();
    const cells: Array<number | undefined> = [];
    for (let index: number = 0; index < offset; index++) cells.push(undefined);
    for (let day: number = 1; day <= totalDays; day++) cells.push(day);

    const card: HTMLElement = document.createElement('section');
    card.className = 'tbxFinalMiniCal';
    card.setAttribute('data-tbx-final-calendar', 'true');
    const monthLabel: string = now.toLocaleDateString('es-CL', { month: 'long', year: 'numeric' });
    card.innerHTML = `
      <div class="tbxFinalMiniCalHead"><strong>${this.escape(monthLabel)}</strong><span class="tbxFinalMiniCalIcon" aria-hidden="true">${this.calendarSvg()}</span></div>
      <div class="tbxFinalMiniCalGrid">
        ${['L','M','X','J','V','S','D'].map((day: string) => `<span class="tbxFinalMiniCalDow">${day}</span>`).join('')}
        ${cells.map((day: number | undefined) => {
          if (!day) return '<span class="tbxFinalMiniCalDay is-empty">·</span>';
          const todayClass: string = day === now.getDate() ? ' is-today' : '';
          const dot: string = eventDays.has(day) ? '<span class="tbxFinalMiniCalDot"></span>' : '';
          return `<span class="tbxFinalMiniCalDay${todayClass}">${day}${dot}</span>`;
        }).join('')}
      </div>
      <div class="tbxFinalMiniCalLegend"><span class="tbxFinalMiniCalLegendDot"></span>Días con eventos agendados</div>`;
    side.insertBefore(card, side.firstElementChild);
  }

  private async loadCalendarEvents(): Promise<CalendarEventItem[]> {
    try {
      const webUrl: string = this.context.pageContext.web.absoluteUrl.replace(/\/$/, '');
      const title: string = 'TIBOX HUB - Eventos';
      const endpoint: string = `${webUrl}/_api/web/lists/getbytitle('${this.odata(title)}')/items?$select=Title,FechaInicio,Activo&$filter=Activo eq 1&$orderby=FechaInicio asc&$top=100`;
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1,
        { headers: { Accept: 'application/json;odata=nometadata' } }
      );
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
    return '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="3.5" y="5" width="17" height="16" rx="3"/><path d="M8 3v4M16 3v4M3.5 10h17"/></svg>';
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
