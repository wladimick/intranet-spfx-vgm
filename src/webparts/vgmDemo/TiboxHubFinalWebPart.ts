import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import TiboxHubGraphWebPart from './TiboxHubGraphWebPart';
import { TIBOX_HUB_DESIGN_CSS } from './TiboxHubDesign';

type FinalCalendarEvent = { Title?: string; FechaInicio?: string };

/** Visual adapter for the approved TIBOX_HUB_1.html prototype. */
export default class TiboxHubFinalWebPart extends TiboxHubGraphWebPart {
  private finalThemeRun: number = 0;

  public render(): void {
    const run: number = ++this.finalThemeRun;
    super.render();
    void this.finalApply(run);
  }

  private async finalApply(run: number): Promise<void> {
    const hub: HTMLElement | undefined = await this.finalWaitForHub(run);
    if (!hub || run !== this.finalThemeRun) return;
    this.finalInjectCss(hub);
    this.finalMapMarkup(hub);
    this.finalDecorateHero(hub);
    await this.finalEnsureCalendar(hub, run);
    this.finalOrderRail(hub);
    [500, 1200, 2400].forEach((ms: number): void => {
      window.setTimeout((): void => {
        if (run !== this.finalThemeRun) return;
        this.finalOrderRail(hub);
        this.finalSyncPhoto(hub);
      }, ms);
    });
  }

  private async finalWaitForHub(run: number): Promise<HTMLElement | undefined> {
    for (let i: number = 0; i < 100; i++) {
      if (run !== this.finalThemeRun) return undefined;
      const hub: HTMLElement | null = this.domElement.querySelector<HTMLElement>('.tbxHub, .tbx-hub');
      if (hub) return hub;
      await this.finalDelay(80);
    }
    return undefined;
  }

  private finalInjectCss(hub: HTMLElement): void {
    if (hub.querySelector('[data-tbx-final-css]')) return;
    const style: HTMLStyleElement = document.createElement('style');
    style.setAttribute('data-tbx-final-css', 'true');
    const design: string = TIBOX_HUB_DESIGN_CSS.replace(/\\n/g, '\n').replace(/\\"/g, '"');
    style.textContent = `
      @import url('https://fonts.googleapis.com/css2?family=Titillium+Web:wght@400;600;700&family=IBM+Plex+Mono:wght@500;600&display=swap');
      ${design}
      .tbx-hub.tbxHub{min-height:100%;border-radius:0;overflow:visible}
      .tbx-hub .tbxShell{max-width:1320px;margin:0 auto;padding:0 28px 64px;background:var(--tbx-bg)}
      .tbx-hub .tbxTop.tbx-header{height:auto;min-height:62px;margin:0 -28px;padding:12px 28px;gap:24px;background:rgba(255,255,255,.88);color:#001233;border-bottom:1px solid var(--tbx-border);backdrop-filter:blur(14px)}
      .tbx-hub .tbxTop .tbxBrand{color:#001233;font-weight:700}.tbx-hub .tbxTop .tbxBrand small{color:#7A8199}.tbx-hub .tbxTop .tbxMark{width:34px;height:34px;background:#000310;box-shadow:none;position:relative}.tbx-hub .tbxTop .tbxMark:after{content:'';position:absolute;inset:9px;border-radius:4px;background:linear-gradient(135deg,#00D1FF,#0E9CDC)}
      .tbx-hub .tbxHero.tbx-hero{min-height:218px;margin:26px 0;padding:30px 34px;border-radius:26px;background:#000310;color:#F4F7FF;position:relative;overflow:hidden}.tbx-hub .tbxHero.tbx-hero:after{content:'';position:absolute;top:-70px;right:-60px;width:260px;height:260px;border-radius:50%;background:radial-gradient(circle,rgba(0,209,255,.20),transparent 70%)}.tbx-hub .tbxHero.tbx-hero:before{content:'PANEL DEL DÍA';display:block;position:relative;z-index:2;margin-bottom:10px;font-family:'IBM Plex Mono',monospace;font-size:10.5px;font-weight:600;letter-spacing:.18em;color:#00D1FF}.tbx-hub .tbxHero h1,.tbx-hub .tbxHero p,.tbx-hub .tbxHero .tbxSearch{position:relative;z-index:2}.tbx-hub .tbxHero h1{color:#F4F7FF}.tbx-hub .tbxHero p{color:#9BA6C4}.tbx-hub .tbxHero .tbxSearch{margin-top:22px;width:min(100%,640px);height:50px;border-radius:15px;background:rgba(255,255,255,.06);border-color:rgba(255,255,255,.08)}.tbx-hub .tbxHero .tbxSearch input{color:#F4F7FF}
      .tbxFinalHeroUser{position:absolute;z-index:3;top:30px;right:34px;display:flex;align-items:center;gap:11px;padding:9px 14px 9px 9px;border:1px solid rgba(255,255,255,.08);border-radius:16px;background:rgba(255,255,255,.04);color:#F4F7FF}.tbxFinalHeroAvatar{width:44px;height:44px;display:grid;place-items:center;border-radius:50%;overflow:hidden;background:rgba(0,209,255,.15);color:#00D1FF;font-weight:700}.tbxFinalHeroAvatar img{width:100%;height:100%;object-fit:cover}.tbxFinalHeroName{font-size:13px;font-weight:700}.tbxFinalHeroRole{color:#9BA6C4;font-size:10.5px}.tbxFinalPulse{position:relative;z-index:2;display:flex;gap:9px;flex-wrap:wrap;margin-top:16px}.tbxFinalPulse span{padding:6px 11px;border:1px solid rgba(255,255,255,.08);border-radius:20px;background:rgba(255,255,255,.05);color:#9BA6C4;font-size:11px}.tbxFinalPulse b{color:#F4F7FF}
      .tbx-hub .tbxQuick.tbx-quicklinks{display:flex;gap:12px;overflow-x:auto;margin:0 0 6px;padding:4px 2px 10px}.tbx-hub .tbxQuick a.tbx-quicklink{min-height:0;flex:0 0 auto;flex-direction:row;align-items:center;justify-content:flex-start;gap:10px;padding:11px 16px;background:#fff;color:#001233}
      .tbx-hub .tbxLayout.tbx-grid{grid-template-columns:minmax(0,1fr) 308px;gap:18px;align-items:start;margin-top:22px}.tbx-hub .tbxMain.tbx-col{gap:22px}.tbx-hub .tbxSide.tbx-rail{gap:14px;position:sticky;top:78px;align-self:start}.tbx-hub .tbxCard.tbx-card{padding:22px;border-radius:22px;overflow:hidden}.tbx-hub .tbxCardHead.tbx-card-head{min-height:0;padding:0;margin-bottom:16px;border:0}.tbx-hub .tbxSide>.tbxCard{padding:16px;border-radius:18px}
      .tbx-hub .tbxPeopleGrid{padding:0}.tbx-hub .tbxPerson{background:var(--tbx-surface-2)}.tbx-hub .tbxMovements{padding:0;display:flex;flex-direction:column;gap:9px}.tbx-hub .tbxMovement{border-bottom:0}.tbx-hub .tbxCourses{padding:0}.tbx-hub .tbxCourse{background:var(--tbx-surface-2)}.tbx-hub .tbxCourseFooter{margin:15px 0 0;padding-top:15px}.tbx-hub .tbxLearning{padding:0;gap:12px}.tbx-hub .tbxLearningList{gap:8px}.tbx-hub .tbxNews{padding:0;gap:9px}.tbx-hub .tbxNewsItem{background:var(--tbx-surface-2);border-color:var(--tbx-border)}
      .tbx-hub .tbxBirthdayList,.tbx-hub .tbxEventList{padding:0;gap:7px}.tbx-hub .tbxBirthday,.tbx-hub .tbxEvent{padding:8px;background:var(--tbx-surface-2)}.tbx-hub .tbxBenefitHero{border-radius:18px}
      .tbx-hub .tbxGraphRoom{border-color:var(--tbx-border);background:#fff;color:#001233;box-shadow:0 1px 2px rgba(0,18,51,.03)}.tbx-hub .tbxGraphRoomTitle,.tbx-hub .tbxGraphRoomStatus{color:#001233}.tbx-hub .tbxGraphRoomSub{color:#66686C}.tbx-hub .tbxGraphRoomNext{background:#F0F2F7;border-color:var(--tbx-border)}.tbx-hub .tbxGraphRoomNext strong{color:#001233}
      .tbx-hub .tbxOverlay{background:rgba(0,3,16,.55);backdrop-filter:blur(3px)}.tbx-hub .tbxModal{border:0;border-radius:26px;background:#F7F7F7;color:#001233;box-shadow:0 30px 80px rgba(0,3,16,.4)}.tbx-hub .tbxModalHead{background:#F7F7F7;border-color:rgba(0,18,51,.07)}.tbx-hub .tbxModalHead h3{color:#001233}.tbx-hub .tbxClose{background:rgba(0,18,51,.06);border:0;color:#5C6478}.tbx-hub .tbxTextarea,.tbx-hub .tbxModalSearch,.tbx-hub .tbxSettingRow input[type=text]{background:#fff;color:#001233;border-color:var(--tbx-border)}.tbx-hub .tbxSettingRow{background:#fff;border-color:var(--tbx-border)}.tbx-hub .tbxAdminNote{background:rgba(14,156,220,.07);color:#3C4253}.tbx-hub .tbxAdminNote strong{color:#001233}.tbx-hub .tbxBtn{background:#fff;color:#001233;border-color:var(--tbx-border)}.tbx-hub .tbxBtn.primary{background:linear-gradient(0deg,#FF4222,#EA7E18);color:#fff}
      @media(max-width:1180px){.tbx-hub .tbxLayout.tbx-grid{grid-template-columns:1fr}.tbx-hub .tbxSide.tbx-rail{position:static;top:auto;gap:22px}}@media(max-width:760px){.tbx-hub .tbxShell{padding:0 16px 40px}.tbx-hub .tbxTop.tbx-header{margin:0 -16px;padding:10px 16px}.tbx-hub .tbxHero.tbx-hero{padding:24px 22px;min-height:0}.tbxFinalHeroUser{position:static;margin-top:16px;width:max-content}.tbx-hub .tbxPeopleGrid,.tbx-hub .tbxCourses,.tbx-hub .tbxLearning,.tbx-hub .tbxNews{grid-template-columns:1fr}}
    `;
    hub.appendChild(style);
  }

  private finalMapMarkup(hub: HTMLElement): void {
    hub.classList.add('tbx-hub');
    hub.setAttribute('data-tbx-theme', 'light');
    this.finalAdd(hub,'.tbxShell',['tbx-shell']); this.finalAdd(hub,'.tbxTop',['tbx-header']); this.finalAdd(hub,'.tbxBrand',['tbx-brand']); this.finalAdd(hub,'.tbxMark',['tbx-brand-mark']); this.finalAdd(hub,'.tbxNav',['tbx-nav']); this.finalAdd(hub,'.tbxNav button',['tbx-nav-btn']); this.finalAdd(hub,'.tbxIconBtn',['tbx-icon-btn']);
    hub.querySelectorAll<HTMLElement>('.tbxNav button.active').forEach((e: HTMLElement) => e.classList.add('is-active'));
    this.finalAdd(hub,'.tbxHero',['tbx-hero']); this.finalAdd(hub,'.tbxHero h1',['tbx-hero-title']); this.finalAdd(hub,'.tbxHero p',['tbx-hero-sub']); this.finalAdd(hub,'.tbxSearch',['tbx-hero-search']);
    this.finalAdd(hub,'.tbxQuick',['tbx-quicklinks']); this.finalAdd(hub,'.tbxQuick a',['tbx-quicklink']); this.finalAdd(hub,'.tbxQuickIcon',['tbx-quicklink-icon']);
    this.finalAdd(hub,'.tbxLayout',['tbx-grid']); this.finalAdd(hub,'.tbxMain',['tbx-col']); this.finalAdd(hub,'.tbxSide',['tbx-col','tbx-rail']); this.finalAdd(hub,'.tbxCard',['tbx-card']); this.finalAdd(hub,'.tbxCardHead',['tbx-card-head']); this.finalAdd(hub,'.tbxCardHead h2,.tbxCardHead h3',['tbx-card-title']); this.finalAdd(hub,'.tbxLink',['tbx-link-btn']);
    this.finalAdd(hub,'.tbxPeopleGrid',['tbx-people-grid']); this.finalAdd(hub,'.tbxPerson',['tbx-person-row']); this.finalAdd(hub,'.tbxAvatar',['tbx-avatar','tbx-avatar-40']); this.finalAdd(hub,'.tbxPersonText strong',['tbx-person-name']); this.finalAdd(hub,'.tbxPersonText span',['tbx-person-meta']); this.finalAdd(hub,'.tbxPersonText small',['tbx-person-joined']);
    this.finalAdd(hub,'.tbxMovement',['tbx-movement-row']); this.finalAdd(hub,'.tbxCourses',['tbx-courses-grid']); this.finalAdd(hub,'.tbxCourse',['tbx-course-card']); this.finalAdd(hub,'.tbxCourseTop',['tbx-course-top']); this.finalAdd(hub,'.tbxCourse h3',['tbx-course-name']); this.finalAdd(hub,'.tbxCourseMeta',['tbx-course-meta']); this.finalAdd(hub,'.tbxCourseAction',['tbx-course-cta']);
    this.finalAdd(hub,'.tbxPrompt',['tbx-prompt-block']); this.finalAdd(hub,'.tbxPrompt label',['tbx-prompt-eyebrow']); this.finalAdd(hub,'.tbxPrompt p',['tbx-prompt-text']); this.finalAdd(hub,'.tbxPrompt button',['tbx-prompt-cta']);
    this.finalAdd(hub,'.tbxBirthday',['tbx-bday-row']); this.finalAdd(hub,'.tbxBirthday button',['tbx-bday-cta']); this.finalAdd(hub,'.tbxEvent',['tbx-event-row']); this.finalAdd(hub,'.tbxDate',['tbx-event-date-chip']); this.finalAdd(hub,'.tbxBenefitHero',['tbx-benefits-teaser']); this.finalAdd(hub,'.tbxBenefitHero label',['tbx-benefits-eyebrow']); this.finalAdd(hub,'.tbxBenefitHero h3',['tbx-benefits-title']); this.finalAdd(hub,'.tbxBenefitHero p',['tbx-benefits-copy']); this.finalAdd(hub,'.tbxBenefitHero button',['tbx-benefits-cta']);
  }

  private finalAdd(root: HTMLElement, selector: string, classes: string[]): void {
    root.querySelectorAll<HTMLElement>(selector).forEach((el: HTMLElement) => el.classList.add(...classes));
  }

  private finalDecorateHero(hub: HTMLElement): void {
    const hero: HTMLElement | null = hub.querySelector<HTMLElement>('.tbxHero'); if (!hero) return;
    const name: string = this.context.pageContext.user.displayName || 'Equipo Tibox';
    const title: HTMLHeadingElement | null = hero.querySelector<HTMLHeadingElement>('h1'); if (title) title.textContent = `Hola, ${name.split(/\s+/)[0] || name}`;
    if (!hero.querySelector('[data-tbx-final-user]')) { const user: HTMLDivElement = document.createElement('div'); user.className='tbxFinalHeroUser'; user.setAttribute('data-tbx-final-user','true'); user.innerHTML=`<span class="tbxFinalHeroAvatar" data-tbx-final-avatar>${this.finalInitials(name)}</span><span><span class="tbxFinalHeroName" style="display:block">${this.finalEscape(name)}</span><span class="tbxFinalHeroRole" style="display:block">Microsoft 365 · Tibox</span></span>`; hero.appendChild(user); }
    if (!hero.querySelector('[data-tbx-final-pulse]')) { const pulse: HTMLDivElement=document.createElement('div'); pulse.className='tbxFinalPulse'; pulse.setAttribute('data-tbx-final-pulse','true'); pulse.innerHTML=`<span><b>${hub.querySelectorAll('.tbxPerson').length}</b> nuevos ingresos</span><span><b>${hub.querySelectorAll('.tbxBirthday').length}</b> próximos cumpleaños</span><span><b>${hub.querySelectorAll('.tbxEvent').length}</b> eventos próximos</span>`; hero.appendChild(pulse); }
  }

  private finalSyncPhoto(hub: HTMLElement): void {
    const source: HTMLImageElement | null = hub.querySelector<HTMLImageElement>('.tbxGraphMe img'); const target: HTMLElement | null = hub.querySelector<HTMLElement>('[data-tbx-final-avatar]'); if (!source || !target || target.querySelector('img')) return; const img: HTMLImageElement=document.createElement('img'); img.src=source.src; img.alt='Foto de usuario'; target.innerHTML=''; target.appendChild(img);
  }

  private async finalEnsureCalendar(hub: HTMLElement, run: number): Promise<void> {
    const side: HTMLElement | null=hub.querySelector<HTMLElement>('.tbxSide'); if (!side || side.querySelector('[data-tbx-final-calendar]')) return;
    const events: FinalCalendarEvent[]=await this.finalLoadEvents(); if (run!==this.finalThemeRun) return;
    const now: Date=new Date(); const year:number=now.getFullYear(); const month:number=now.getMonth(); const eventDays:Set<number>=new Set<number>();
    events.forEach((item:FinalCalendarEvent)=>{if(!item.FechaInicio)return;const d:Date=new Date(item.FechaInicio);if(d.getFullYear()===year&&d.getMonth()===month)eventDays.add(d.getDate());});
    let offset:number=new Date(year,month,1).getDay()-1;if(offset<0)offset=6;const total:number=new Date(year,month+1,0).getDate();const cells:Array<number|undefined>=[];for(let i:number=0;i<offset;i++)cells.push(undefined);for(let d:number=1;d<=total;d++)cells.push(d);
    const card:HTMLElement=document.createElement('section');card.className='tbx-card tbx-mini-cal';card.setAttribute('data-tbx-final-calendar','true');card.innerHTML=`<div class="tbx-mini-cal-head"><span class="tbx-mini-cal-month">${this.finalEscape(now.toLocaleDateString('es-CL',{month:'long',year:'numeric'}))}</span>${this.finalCalendarSvg()}</div><div class="tbx-mini-cal-grid">${['L','M','X','J','V','S','D'].map((x:string)=>`<span class="tbx-mini-cal-dow">${x}</span>`).join('')}${cells.map((d:number|undefined)=>!d?'<span class="tbx-mini-cal-day is-blank">·</span>':`<span class="tbx-mini-cal-day${d===now.getDate()?' is-today':''}">${d}${eventDays.has(d)?'<span class="tbx-mini-cal-dot"></span>':''}</span>`).join('')}</div><div class="tbx-mini-cal-legend"><span class="tbx-mini-cal-dot-static"></span>Días con eventos agendados</div>`;side.insertBefore(card,side.firstElementChild);
  }

  private async finalLoadEvents(): Promise<FinalCalendarEvent[]> {
    try { const web:string=this.context.pageContext.web.absoluteUrl.replace(/\/$/,''); const title:string='TIBOX HUB - Eventos'; const response:SPHttpClientResponse=await this.context.spHttpClient.get(`${web}/_api/web/lists/getbytitle('${this.finalOdata(title)}')/items?$select=Title,FechaInicio,Activo&$filter=Activo eq 1&$orderby=FechaInicio asc&$top=100`,SPHttpClient.configurations.v1,{headers:{Accept:'application/json;odata=nometadata'}}); if(!response.ok)return[]; const data:{value?:FinalCalendarEvent[]}=await response.json() as {value?:FinalCalendarEvent[]}; return data.value||[]; } catch { return []; }
  }

  private finalOrderRail(hub: HTMLElement): void { const side:HTMLElement|null=hub.querySelector<HTMLElement>('.tbxSide');if(!side)return;const cal:Element|null=side.querySelector('[data-tbx-final-calendar]');const room:Element|null=side.querySelector('[data-tbx-room-card]');if(cal&&side.firstElementChild!==cal)side.insertBefore(cal,side.firstElementChild);if(cal&&room&&cal.nextElementSibling!==room)side.insertBefore(room,cal.nextElementSibling); }
  private finalCalendarSvg():string{return '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="3.5" y="5" width="17" height="16" rx="3"/><path d="M8 3v4M16 3v4M3.5 10h17"/></svg>';}
  private finalInitials(name:string):string{return name.split(/\s+/).filter(Boolean).slice(0,2).map((p:string)=>p.charAt(0).toUpperCase()).join('');}
  private finalEscape(value:string):string{return value.replace(/[&<>'\"]/g,(c:string)=>({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','\"':'&quot;'}[c]||c));}
  private finalOdata(value:string):string{return value.replace(/'/g,"''");}
  private finalDelay(ms:number):Promise<void>{return new Promise<void>((resolve:()=>void)=>window.setTimeout(resolve,ms));}
}
