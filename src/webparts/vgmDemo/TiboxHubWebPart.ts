import { Version } from '@microsoft/sp-core-library';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  HubBenefitItem,
  HubCourseItem,
  HubDashboardData,
  HubEventItem,
  HubLearningItem,
  HubMovementItem,
  HubNewsItem,
  HubPersonItem,
  TiboxHubSharePointService
} from './TiboxHubSharePointService';

export interface ITiboxHubWebPartProps {}

type QuickLink = {
  key: string;
  label: string;
  url: string;
  enabled: boolean;
};

type CurrentUserInfo = {
  Id: number;
  LoginName: string;
  IsSiteAdmin: boolean;
};

export default class TiboxHubWebPart extends BaseClientSideWebPart<ITiboxHubWebPartProps> {
  private readonly configListTitle: string = 'TiboxHubConfig';
  private readonly adminGroupTitle: string = 'TIBOX HUB - Administradores';
  private readonly quickLinksConfigKey: string = 'QuickLinks';

  private isHubAdmin: boolean = false;
  private isSiteAdmin: boolean = false;
  private quickLinks: QuickLink[] = this.defaultQuickLinks();
  private data: HubDashboardData = this.emptyData();

  public render(): void {
    void this.renderAsync();
  }

  private async renderAsync(): Promise<void> {
    this.renderLoading();
    this.quickLinks = this.defaultQuickLinks();

    const currentUser: CurrentUserInfo | undefined = await this.getCurrentUser();
    if (currentUser) {
      this.isSiteAdmin = currentUser.IsSiteAdmin;
      this.isHubAdmin = currentUser.IsSiteAdmin || await this.isUserInAdminGroup(currentUser.Id);

      if (currentUser.IsSiteAdmin) {
        await this.ensureAdminGroup(currentUser.LoginName);
        this.isHubAdmin = true;
      }
    }

    if (this.isHubAdmin) {
      await this.ensureConfigurationList();
    }
    await this.loadGlobalSettings();

    const dataService: TiboxHubSharePointService = new TiboxHubSharePointService(this.context.spHttpClient, this.webUrl);
    if (this.isSiteAdmin) {
      try {
        await dataService.ensureSchemaAndSeed();
      } catch (error) {
        console.error('TIBOX HUB provisioning error', error);
      }
    }
    this.data = await dataService.loadDashboardData();

    const userName: string = this.context.pageContext.user.displayName || 'Equipo Tibox';
    const adminNav: string = this.isHubAdmin ? '<button data-open="settings">Administrar</button>' : '';
    const adminButton: string = this.isHubAdmin
      ? `<button class="tbxIconBtn" data-open="settings" title="Configuración" aria-label="Configuración">${this.icon('settings')}</button>`
      : '';

    const settingsModal: string = this.isHubAdmin
      ? this.modal(
          'settingsModal',
          'Configuración de TIBOX HUB',
          `<div class="tbxAdminNote"><strong>Configuración global</strong><span>Solo administradores de TIBOX HUB ven este panel. Los cambios se guardan en SharePoint y se aplican a todos los usuarios.</span></div>
           <div class="tbxAdminNote"><strong>Contenido conectado a SharePoint</strong><span>Colaboradores, Movimientos, Cursos, Beneficios, Eventos, Aprende y Noticias ya se leen desde listas del sitio.</span><a class="tbxAdminLink" href="${this.webUrl}/_layouts/15/viewlsts.aspx" target="_blank" rel="noopener noreferrer">Abrir contenido del sitio ↗</a></div>
           <div class="tbxSettingsGrid" data-settings-grid></div>
           <div class="tbxModalActions"><button class="tbxBtn" data-reset-settings>Restaurar valores</button><button class="tbxBtn primary" data-save-settings>Guardar cambios</button></div>`,
          false
        )
      : '';

    const hires: HubPersonItem[] = this.newHires();
    const birthdays: HubPersonItem[] = this.upcomingBirthdays().slice(0, 4);
    const courses: HubCourseItem[] = this.data.courses.slice(0, 3);
    const movements: HubMovementItem[] = this.data.movements.slice(0, 4);
    const events: HubEventItem[] = this.futureEvents().slice(0, 4);
    const news: HubNewsItem[] = this.data.news.slice(0, 4);
    const featuredLearning: HubLearningItem | undefined = this.data.learning.find((item: HubLearningItem) => item.Destacado) || this.data.learning[0];

    this.domElement.innerHTML = `
      <style>
        .tbxHub,.tbxHub *{box-sizing:border-box}
        .tbxHub{--bg:#000310;--surface:#0A1130;--surface2:#121A40;--border:rgba(255,255,255,.08);--borderAccent:rgba(0,209,255,.28);--text:#F4F7FF;--muted:#9BA6C4;--subtle:#5E6A8A;--support:#00D1FF;--supportSolid:#0E9CDC;--cta:#FF4222;--yellow:#FFB200;--green:#55D9A6;--red:#FF725E;--violet:#B9ABFF;width:100%;min-height:100vh;color:var(--muted);background:radial-gradient(circle at 78% -8%,rgba(0,209,255,.10),transparent 34rem),var(--bg);font-family:'Segoe UI',Arial,sans-serif;font-size:14px;border-radius:18px;overflow:hidden}
        .tbxHub button,.tbxHub input,.tbxHub textarea{font:inherit}.tbxShell{max-width:1440px;margin:0 auto;padding:0 28px 54px}
        .tbxTop{position:sticky;top:0;z-index:20;height:72px;margin:0 -28px;padding:0 28px;display:flex;align-items:center;justify-content:space-between;background:rgba(0,3,16,.84);backdrop-filter:blur(16px);border-bottom:1px solid var(--border)}
        .tbxBrand{display:flex;align-items:center;gap:11px;color:var(--text);font-weight:800;letter-spacing:.02em}.tbxMark{width:30px;height:30px;border-radius:9px;background:linear-gradient(135deg,#00D1FF,#0E9CDC);box-shadow:0 0 0 5px rgba(0,209,255,.07)}.tbxBrand small{display:block;margin-top:2px;color:var(--subtle);font-weight:500;font-size:10px}
        .tbxNav{display:flex;align-items:center;gap:6px;padding:4px;border:1px solid var(--border);background:rgba(18,26,64,.55);border-radius:12px}.tbxNav button{border:0;color:var(--muted);background:transparent;padding:8px 13px;border-radius:9px;cursor:pointer;font-weight:600;font-size:12px}.tbxNav button.active{background:var(--surface2);color:var(--text)}
        .tbxTopActions{display:flex;gap:8px}.tbxIconBtn{width:38px;height:38px;display:grid;place-items:center;border:1px solid var(--border);border-radius:10px;color:var(--text);background:var(--surface);cursor:pointer}.tbxIconBtn:hover{border-color:var(--borderAccent);background:var(--surface2)}.tbxIcon{width:18px;height:18px;stroke:currentColor;fill:none;stroke-width:1.8;stroke-linecap:round;stroke-linejoin:round}
        .tbxHero{padding:44px 0 22px}.tbxHero h1{margin:0;color:var(--text);font-size:38px;line-height:1.08;letter-spacing:-.025em}.tbxHero p{margin:9px 0 0;font-size:15px}.tbxSearch{margin-top:25px;width:min(100%,650px);height:46px;display:flex;align-items:center;gap:10px;padding:0 14px;border:1px solid var(--border);border-radius:12px;background:var(--surface)}.tbxSearch input{flex:1;border:0;outline:0;background:transparent;color:var(--text)}.tbxSearch input::placeholder{color:var(--subtle)}
        .tbxQuick{display:grid;grid-template-columns:repeat(7,minmax(0,1fr));gap:10px;margin:12px 0 26px}.tbxQuick a{min-height:74px;padding:13px;display:flex;flex-direction:column;justify-content:space-between;border:1px solid var(--border);border-radius:13px;background:var(--surface);color:var(--text);text-decoration:none;font-size:12px;font-weight:650;transition:.15s}.tbxQuick a:hover{transform:translateY(-1px);border-color:var(--borderAccent);background:var(--surface2)}.tbxQuickIcon{width:29px;height:29px;border-radius:9px;display:grid;place-items:center;color:var(--support);background:rgba(0,209,255,.08)}
        .tbxLayout{display:grid;grid-template-columns:minmax(0,1fr) 300px;gap:18px}.tbxMain{min-width:0;display:flex;flex-direction:column;gap:18px}.tbxSide{display:flex;flex-direction:column;gap:18px}.tbxCard{border:1px solid var(--border);border-radius:18px;background:var(--surface);box-shadow:0 12px 30px rgba(0,0,0,.10);overflow:hidden}.tbxCardHead{min-height:54px;padding:14px 17px;display:flex;align-items:center;justify-content:space-between;gap:12px;border-bottom:1px solid var(--border)}.tbxCardHead h2,.tbxCardHead h3{margin:0;color:var(--text);font-size:15px}.tbxLink{border:0;background:transparent;color:var(--support);font-weight:700;font-size:11px;cursor:pointer}
        .tbxPeopleGrid{padding:12px;display:grid;grid-template-columns:1fr 1fr;gap:9px}.tbxPerson{padding:11px;min-width:0;display:flex;align-items:center;gap:11px;border:1px solid var(--border);border-radius:12px;background:var(--surface2);cursor:pointer}.tbxPerson:hover{border-color:var(--borderAccent)}.tbxAvatar{width:38px;height:38px;flex:none;border-radius:50%;display:grid;place-items:center;color:#001233;background:linear-gradient(135deg,#00D1FF,#88E8FF);font-weight:800;font-size:11px}.tbxPersonText{min-width:0;flex:1}.tbxPersonText strong{display:block;color:var(--text);font-size:12px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.tbxPersonText span{display:block;margin-top:2px;color:var(--muted);font-size:10px}.tbxPersonText small{color:var(--support);font-size:9px}
        .tbxMovements{padding:10px 12px 12px}.tbxMovement{display:flex;align-items:center;gap:11px;padding:11px 7px;border-bottom:1px solid var(--border)}.tbxMovement:last-child{border-bottom:0}.tbxMovement .change{margin-left:auto;padding:4px 8px;border-radius:999px;font-size:9px;font-weight:700;color:var(--green);border:1px solid rgba(85,217,166,.20);background:rgba(85,217,166,.08)}
        .tbxCourses{padding:13px;display:grid;grid-template-columns:repeat(3,1fr);gap:10px}.tbxCourse{padding:14px;border:1px solid var(--border);border-radius:13px;background:var(--surface2)}.tbxCourseTop{display:flex;justify-content:space-between;gap:7px}.tbxPill{display:inline-flex;padding:4px 7px;border-radius:999px;font-size:9px;font-weight:800}.tbxPill.warn{color:var(--yellow);background:rgba(255,178,0,.09)}.tbxPill.blue{color:var(--support);background:rgba(0,209,255,.08)}.tbxPill.new{color:var(--violet);background:rgba(185,171,255,.09)}.tbxPill.required{color:var(--red);background:rgba(255,114,94,.08)}.tbxCourse h3{color:var(--text);margin:12px 0 6px;font-size:12px}.tbxCourseMeta{display:flex;gap:12px;color:var(--muted);font-size:10px;line-height:1.5}.tbxCourseAction{display:inline-flex;margin-top:12px;border:0;border-radius:8px;padding:8px 10px;background:var(--supportSolid);color:white;text-decoration:none;font-size:10px;font-weight:800;cursor:pointer}.tbxCourseFooter{margin:0 13px 13px;padding:11px 0 0;border-top:1px solid var(--border);display:flex;flex-wrap:wrap;gap:7px}.tbxStatus{padding:5px 8px;border-radius:999px;background:var(--surface2);font-size:9px}
        .tbxLearning{padding:13px;display:grid;grid-template-columns:1fr 1fr;gap:12px}.tbxPrompt{min-height:160px;padding:18px;border-radius:14px;color:var(--text);background:linear-gradient(145deg,#121A40,#0A1130);border:1px solid var(--borderAccent)}.tbxPrompt label{color:var(--support);font-size:9px;font-weight:800;letter-spacing:.08em}.tbxPrompt p{margin:14px 0;color:var(--text);font-style:italic;line-height:1.55;font-size:12px}.tbxPrompt button{border:1px solid var(--border);border-radius:8px;padding:8px 10px;background:var(--text);color:#001233;font-weight:800;font-size:10px;cursor:pointer}.tbxLearningList{display:flex;flex-direction:column;gap:7px}.tbxLearningItem{padding:11px 12px;display:flex;align-items:center;justify-content:space-between;border:1px solid var(--border);border-radius:11px;background:var(--surface2);color:var(--text)}.tbxLearningItem small{display:block;margin-top:2px;color:var(--subtle)}
        .tbxNews{padding:12px;display:grid;grid-template-columns:1fr 1fr;gap:9px}.tbxNewsItem{min-height:82px;padding:13px;border:1px solid var(--border);border-radius:12px;background:var(--surface2)}.tbxNewsItem label{font-size:9px;color:var(--support)}.tbxNewsItem strong{display:block;margin-top:8px;color:var(--text);font-size:11px}.tbxNewsItem small{display:block;margin-top:5px;color:var(--subtle);font-size:9px}
        .tbxBirthdayList,.tbxEventList{padding:10px;display:flex;flex-direction:column;gap:8px}.tbxBirthday,.tbxEvent{display:flex;align-items:center;gap:9px;padding:10px;border:1px solid var(--border);border-radius:11px;background:var(--surface2)}.tbxBirthday .tbxAvatar{width:34px;height:34px}.tbxBirthday button{margin-left:auto;border:1px solid rgba(255,178,0,.20);border-radius:8px;background:rgba(255,178,0,.09);color:var(--yellow);padding:6px 8px;font-weight:800;font-size:9px;cursor:pointer}.tbxDate{width:37px;height:42px;flex:none;border-radius:9px;display:grid;place-items:center;background:#050B26;color:var(--support);font-weight:800;line-height:1.05;font-size:13px}.tbxDate small{display:block;font-size:7px}
        .tbxBenefitHero{min-height:155px;padding:18px;background:radial-gradient(circle at 100% 0,rgba(0,209,255,.18),transparent 170px),#050B26}.tbxBenefitHero label{color:var(--support);font-size:9px;font-weight:800}.tbxBenefitHero h3{color:var(--text);margin:12px 0 7px;font-size:15px}.tbxBenefitHero p{margin:0;font-size:10px;line-height:1.55}.tbxBenefitHero button{margin-top:14px;border:0;border-radius:9px;padding:9px 12px;background:var(--cta);color:#fff;font-weight:800;font-size:10px;cursor:pointer}
        .tbxEmpty{padding:18px;color:var(--subtle);font-size:11px}.tbxOverlay{position:fixed;inset:0;z-index:1000;display:none;align-items:center;justify-content:center;padding:28px;background:rgba(0,3,16,.72);backdrop-filter:blur(7px)}.tbxOverlay.open{display:flex}.tbxModal{width:min(920px,96vw);max-height:86vh;overflow:auto;border:1px solid var(--border);border-radius:20px;background:var(--surface);box-shadow:0 30px 100px rgba(0,0,0,.46)}.tbxModal.sm{width:min(560px,96vw)}.tbxModalHead{position:sticky;top:0;z-index:2;min-height:62px;padding:15px 18px;display:flex;align-items:center;justify-content:space-between;background:rgba(10,17,48,.96);border-bottom:1px solid var(--border)}.tbxModalHead h3{margin:0;color:var(--text)}.tbxClose{width:34px;height:34px;border:1px solid var(--border);border-radius:9px;background:var(--surface2);color:var(--text);cursor:pointer}.tbxModalBody{padding:18px}.tbxModalSearch{width:100%;height:42px;padding:0 12px;border:1px solid var(--border);border-radius:10px;background:var(--surface2);color:var(--text);outline:0}.tbxAllHires{margin-top:14px;display:grid;grid-template-columns:1fr 1fr;gap:9px}.tbxTextarea{width:100%;min-height:110px;padding:12px;resize:vertical;border:1px solid var(--border);border-radius:11px;background:var(--surface2);color:var(--text);outline:0}.tbxModalActions{display:flex;justify-content:flex-end;gap:8px;margin-top:14px}.tbxBtn{border:1px solid var(--border);border-radius:9px;padding:9px 12px;background:var(--surface2);color:var(--text);font-weight:700;cursor:pointer}.tbxBtn.primary{background:var(--cta);border-color:var(--cta);color:#fff}
        .tbxAdminNote{display:flex;flex-direction:column;gap:5px;margin-bottom:14px;padding:12px 13px;border:1px solid rgba(0,209,255,.18);border-radius:11px;background:rgba(0,209,255,.06)}.tbxAdminNote strong{color:var(--text);font-size:12px}.tbxAdminNote span{font-size:10px;line-height:1.5}.tbxAdminLink{margin-top:5px;color:var(--support);font-size:10px;font-weight:700;text-decoration:none}.tbxSettingsGrid{display:grid;grid-template-columns:1fr;gap:10px}.tbxSettingRow{padding:12px;display:grid;grid-template-columns:34px 1fr 1.2fr 44px;align-items:center;gap:9px;border:1px solid var(--border);border-radius:11px;background:var(--surface2)}.tbxSettingRow input[type='text']{width:100%;height:34px;padding:0 9px;border:1px solid var(--border);border-radius:8px;background:var(--surface);color:var(--text)}.tbxSettingRow input[type='checkbox']{width:18px;height:18px;accent-color:var(--supportSolid)}
        .tbxToast{position:fixed;z-index:1200;right:24px;bottom:24px;max-width:340px;padding:12px 15px;border:1px solid var(--borderAccent);border-radius:11px;background:#0A1130;color:var(--text);box-shadow:0 18px 50px rgba(0,0,0,.38);opacity:0;transform:translateY(8px);pointer-events:none;transition:.2s}.tbxToast.show{opacity:1;transform:none}
        @media(max-width:1100px){.tbxQuick{grid-template-columns:repeat(4,1fr)}.tbxLayout{grid-template-columns:1fr}.tbxSide{display:grid;grid-template-columns:1fr 1fr}}
        @media(max-width:760px){.tbxShell{padding:0 14px 34px}.tbxTop{margin:0 -14px;padding:0 14px}.tbxNav{display:none}.tbxHero h1{font-size:30px}.tbxQuick{grid-template-columns:repeat(2,1fr)}.tbxPeopleGrid,.tbxCourses,.tbxLearning,.tbxNews,.tbxAllHires{grid-template-columns:1fr}.tbxSide{display:flex}.tbxSettingRow{grid-template-columns:30px 1fr}.tbxSettingRow input[type='text']{grid-column:1/-1}.tbxSettingRow input[type='checkbox']{grid-column:2}}
      </style>

      <div class="tbxHub">
        <div class="tbxShell">
          <header class="tbxTop">
            <div class="tbxBrand"><span class="tbxMark"></span><div>TIBOX HUB<small>Todo Tibox, en un solo lugar.</small></div></div>
            <nav class="tbxNav"><button class="active">Inicio</button><button data-scroll="aprende">Tibox Aprende</button>${adminNav}</nav>
            <div class="tbxTopActions"><button class="tbxIconBtn" title="Notificaciones" aria-label="Notificaciones">${this.icon('bell')}</button>${adminButton}</div>
          </header>

          <section class="tbxHero"><h1>Hola, ${this.escapeHtml(userName)}</h1><p>Esto está pasando en Tibox.</p><label class="tbxSearch">${this.icon('search')}<input data-global-search type="search" placeholder="Buscar personas, cursos, beneficios, documentos..." /></label></section>
          <section class="tbxQuick" data-quick-links></section>

          <div class="tbxLayout">
            <main class="tbxMain">
              <section class="tbxCard" id="personas"><div class="tbxCardHead"><h2>Nuevos en Tibox</h2><button class="tbxLink" data-open="hires">Ver todos los nuevos ingresos ›</button></div><div class="tbxPeopleGrid">${hires.length ? hires.slice(0, 4).map((person: HubPersonItem) => this.personCard(person)).join('') : this.empty('No hay nuevos ingresos publicados.')}</div></section>
              <section class="tbxCard"><div class="tbxCardHead"><h2>Movimientos en Tibox</h2></div><div class="tbxMovements">${movements.length ? movements.map((movement: HubMovementItem) => this.movement(movement)).join('') : this.empty('No hay movimientos publicados.')}</div></section>
              <section class="tbxCard" id="cursos"><div class="tbxCardHead"><h2>Cursos del mes</h2><button class="tbxLink" data-list-link="courses">Ver todos los cursos ›</button></div><div class="tbxCourses">${courses.length ? courses.map((course: HubCourseItem) => this.course(course)).join('') : this.empty('No hay cursos activos.')}</div><div class="tbxCourseFooter">${courses.length ? courses.map((course: HubCourseItem) => `<span class="tbxStatus">${this.escapeHtml(course.Estado || 'Activo')} · ${this.escapeHtml(course.Title)}</span>`).join('') : ''}</div></section>
              <section class="tbxCard" id="aprende"><div class="tbxCardHead"><h2>Tibox Aprende</h2><button class="tbxLink" data-list-link="learning">Explorar ›</button></div><div class="tbxLearning">${featuredLearning ? `<div class="tbxPrompt"><label>✦ ${this.escapeHtml((featuredLearning.TipoContenido || 'DESTACADO').toUpperCase())}</label><p>“${this.escapeHtml(featuredLearning.Descripcion || featuredLearning.Title)}”</p><button data-copy-prompt data-prompt="${this.escapeHtml(featuredLearning.Descripcion || featuredLearning.Title)}">Copiar</button></div>` : this.empty('No hay contenido destacado.')}<div class="tbxLearningList">${this.learningCategories()}</div></div></section>
              <section class="tbxCard"><div class="tbxCardHead"><h2>Tibox informa</h2><button class="tbxLink" data-list-link="news">Ver noticias ›</button></div><div class="tbxNews">${news.length ? news.map((item: HubNewsItem) => this.news(item)).join('') : this.empty('No hay noticias publicadas.')}</div></section>
            </main>
            <aside class="tbxSide">
              <section class="tbxCard"><div class="tbxCardHead"><h3>Próximos cumpleaños</h3></div><div class="tbxBirthdayList">${birthdays.length ? birthdays.map((person: HubPersonItem) => this.birthday(person)).join('') : this.empty('No hay cumpleaños próximos.')}</div></section>
              <section class="tbxCard"><div class="tbxCardHead"><h3>Próximos eventos</h3><button class="tbxLink" data-list-link="events">Ver todos ›</button></div><div class="tbxEventList">${events.length ? events.map((event: HubEventItem) => this.event(event)).join('') : this.empty('No hay eventos próximos.')}</div></section>
              <section class="tbxCard" id="beneficios"><div class="tbxBenefitHero"><label>♡ BENEFICIOS</label><h3>Beneficios pensados para ti</h3><p>${this.data.benefits.length} beneficios activos disponibles.</p><button data-open="benefits">Ver beneficios</button></div></section>
            </aside>
          </div>
        </div>

        ${this.modal('hiresModal','Nuevos en Tibox',`<input class="tbxModalSearch" data-hire-search placeholder="Buscar por nombre, cargo o área..." /><div class="tbxAllHires" data-all-hires>${hires.length ? hires.map((person: HubPersonItem) => this.personCard(person)).join('') : this.empty('No hay nuevos ingresos publicados.')}</div>`,false)}
        ${this.modal('birthdayModal','Felicitar cumpleaños',`<p style="margin:0 0 12px">Escribe tu saludo. La integración con Teams será el siguiente paso.</p><textarea class="tbxTextarea" data-birthday-message></textarea><div class="tbxModalActions"><button class="tbxBtn" data-close>Cancelar</button><button class="tbxBtn primary" data-send-teams>Abrir en Teams</button></div>`,true)}
        ${this.modal('benefitsModal','Beneficios Tibox',`<div class="tbxAllHires">${this.data.benefits.length ? this.data.benefits.map((benefit: HubBenefitItem) => this.benefit(benefit)).join('') : this.empty('No hay beneficios activos.')}</div>`,false)}
        ${settingsModal}
        <div class="tbxToast" data-toast-box></div>
      </div>
    `;

    this.renderQuickLinks();
    if (this.isHubAdmin) this.renderSettings();
    this.bindEvents();
  }

  private renderLoading(): void {
    this.domElement.innerHTML = '<div style="min-height:360px;background:#000310;color:#9BA6C4;border-radius:18px;display:grid;place-items:center;font-family:Segoe UI,Arial,sans-serif">Cargando TIBOX HUB…</div>';
  }

  private bindEvents(): void {
    this.domElement.querySelectorAll<HTMLElement>('[data-open]').forEach((el: HTMLElement) => el.addEventListener('click', () => this.openModal(el.getAttribute('data-open') || '')));
    this.domElement.querySelectorAll<HTMLElement>('[data-close]').forEach((el: HTMLElement) => el.addEventListener('click', () => this.closeAllModals()));
    this.domElement.querySelectorAll<HTMLElement>('.tbxOverlay').forEach((overlay: HTMLElement) => overlay.addEventListener('click', (event: MouseEvent) => { if (event.target === overlay) this.closeAllModals(); }));
    this.domElement.querySelectorAll<HTMLElement>('[data-scroll]').forEach((el: HTMLElement) => el.addEventListener('click', () => { const target: HTMLElement | null = this.domElement.querySelector(`#${el.getAttribute('data-scroll')}`); target?.scrollIntoView({ behavior: 'smooth', block: 'start' }); }));
    this.domElement.querySelectorAll<HTMLElement>('[data-person]').forEach((el: HTMLElement) => el.addEventListener('click', () => this.toast(`${el.getAttribute('data-person')}: el perfil se conectará a Microsoft 365 en la siguiente etapa.`)));
    this.domElement.querySelectorAll<HTMLElement>('[data-birthday]').forEach((el: HTMLElement) => el.addEventListener('click', () => { const name: string = el.getAttribute('data-birthday') || ''; const firstName: string = name.split(' ')[0]; const area: HTMLTextAreaElement | null = this.domElement.querySelector('[data-birthday-message]'); if (area) area.value = `¡Feliz cumpleaños, ${firstName}! 🎉 Que tengas un excelente día y un gran año.`; this.openModal('birthday'); }));

    const hireSearch: HTMLInputElement | null = this.domElement.querySelector('[data-hire-search]');
    hireSearch?.addEventListener('input', () => this.filterHires(hireSearch.value));

    const globalSearch: HTMLInputElement | null = this.domElement.querySelector('[data-global-search]');
    globalSearch?.addEventListener('keydown', (event: KeyboardEvent) => { if (event.key === 'Enter' && globalSearch.value.trim()) this.toast(`Búsqueda: “${globalSearch.value.trim()}”`); });

    const copy: HTMLElement | null = this.domElement.querySelector('[data-copy-prompt]');
    copy?.addEventListener('click', async () => { const prompt: string = copy.getAttribute('data-prompt') || ''; try { await navigator.clipboard.writeText(prompt); this.toast('Contenido copiado'); } catch { this.toast('No se pudo usar el portapapeles en este contexto.'); } });

    const sendTeams: HTMLElement | null = this.domElement.querySelector('[data-send-teams]');
    sendTeams?.addEventListener('click', () => this.toast('La integración con Teams será el siguiente paso.'));

    this.domElement.querySelectorAll<HTMLElement>('[data-list-link]').forEach((el: HTMLElement) => el.addEventListener('click', () => {
      const key: string = el.getAttribute('data-list-link') || '';
      const listTitle: string | undefined = this.listTitleFromKey(key);
      if (listTitle) window.open(`${this.webUrl}/Lists/${encodeURIComponent(listTitle)}/AllItems.aspx`, '_blank', 'noopener,noreferrer');
    }));

    if (this.isHubAdmin) {
      const save: HTMLElement | null = this.domElement.querySelector('[data-save-settings]');
      save?.addEventListener('click', () => { void this.saveSettingsFromForm(); });
      const reset: HTMLElement | null = this.domElement.querySelector('[data-reset-settings]');
      reset?.addEventListener('click', () => { void this.resetGlobalSettings(); });
    }

    document.addEventListener('keydown', this.onDocumentKeyDown);
  }

  private readonly onDocumentKeyDown = (event: KeyboardEvent): void => { if (event.key === 'Escape') this.closeAllModals(); };

  private openModal(name: string): void {
    if (name === 'settings' && !this.isHubAdmin) return;
    const idMap: Record<string,string> = { hires:'hiresModal', birthday:'birthdayModal', benefits:'benefitsModal', settings:'settingsModal' };
    const modal: HTMLElement | null = this.domElement.querySelector(`#${idMap[name] || name}`);
    modal?.classList.add('open');
  }

  private closeAllModals(): void { this.domElement.querySelectorAll<HTMLElement>('.tbxOverlay.open').forEach((el: HTMLElement) => el.classList.remove('open')); }

  private filterHires(value: string): void {
    const query: string = value.trim().toLowerCase();
    const area: HTMLElement | null = this.domElement.querySelector('[data-all-hires]');
    if (!area) return;
    const filtered: HubPersonItem[] = this.newHires().filter((person: HubPersonItem) => `${person.Title} ${person.Cargo || ''} ${person.Area || ''}`.toLowerCase().includes(query));
    area.innerHTML = filtered.length ? filtered.map((person: HubPersonItem) => this.personCard(person)).join('') : this.empty('No se encontraron personas.');
  }

  private renderQuickLinks(): void {
    const area: HTMLElement | null = this.domElement.querySelector('[data-quick-links]');
    if (!area) return;
    area.innerHTML = this.quickLinks.filter((link: QuickLink) => link.enabled).map((link: QuickLink) => `<a href="${this.escapeHtml(link.url || '#')}" ${link.url.startsWith('http') ? 'target="_blank" rel="noopener noreferrer"' : ''}><span class="tbxQuickIcon">${this.icon(this.quickIcon(link.key))}</span><span>${this.escapeHtml(link.label)}</span></a>`).join('');
  }

  private renderSettings(): void {
    if (!this.isHubAdmin) return;
    const area: HTMLElement | null = this.domElement.querySelector('[data-settings-grid]');
    if (!area) return;
    area.innerHTML = this.quickLinks.map((link: QuickLink, index: number) => `<div class="tbxSettingRow" data-setting-index="${index}"><span class="tbxQuickIcon">${this.icon(this.quickIcon(link.key))}</span><input data-label type="text" value="${this.escapeHtml(link.label)}" aria-label="Nombre" /><input data-url type="text" value="${this.escapeHtml(link.url)}" aria-label="URL" /><input data-enabled type="checkbox" ${link.enabled ? 'checked' : ''} aria-label="Activo" /></div>`).join('');
  }

  private async saveSettingsFromForm(): Promise<void> {
    if (!this.isHubAdmin) { this.toast('No tienes permisos para cambiar la configuración.'); return; }
    this.domElement.querySelectorAll<HTMLElement>('[data-setting-index]').forEach((row: HTMLElement) => {
      const index: number = Number(row.getAttribute('data-setting-index'));
      const label: HTMLInputElement | null = row.querySelector('[data-label]');
      const url: HTMLInputElement | null = row.querySelector('[data-url]');
      const enabled: HTMLInputElement | null = row.querySelector('[data-enabled]');
      if (this.quickLinks[index]) {
        this.quickLinks[index].label = label?.value.trim() || this.quickLinks[index].label;
        this.quickLinks[index].url = url?.value.trim() || '#';
        this.quickLinks[index].enabled = Boolean(enabled?.checked);
      }
    });
    try { await this.saveGlobalSettings(); this.renderQuickLinks(); this.closeAllModals(); this.toast('Configuración global guardada en SharePoint'); }
    catch { this.toast('No se pudo guardar la configuración en SharePoint.'); }
  }

  private async resetGlobalSettings(): Promise<void> {
    if (!this.isHubAdmin) return;
    this.quickLinks = this.defaultQuickLinks();
    try { await this.saveGlobalSettings(); this.renderQuickLinks(); this.renderSettings(); this.toast('Configuración global restaurada'); }
    catch { this.toast('No se pudo restaurar la configuración.'); }
  }

  private async getCurrentUser(): Promise<CurrentUserInfo | undefined> {
    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(`${this.webUrl}/_api/web/currentuser?$select=Id,LoginName,IsSiteAdmin`, SPHttpClient.configurations.v1, { headers: { Accept: 'application/json;odata=nometadata' } });
      if (!response.ok) return undefined;
      return await response.json() as CurrentUserInfo;
    } catch { return undefined; }
  }

  private async isUserInAdminGroup(userId: number): Promise<boolean> {
    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(`${this.webUrl}/_api/web/sitegroups/getbyname('${this.odataString(this.adminGroupTitle)}')/users?$select=Id`, SPHttpClient.configurations.v1, { headers: { Accept: 'application/json;odata=nometadata' } });
      if (!response.ok) return false;
      const data: { value?: Array<{ Id: number }> } = await response.json() as { value?: Array<{ Id: number }> };
      return Boolean(data.value?.some((user: { Id: number }) => user.Id === userId));
    } catch { return false; }
  }

  private async ensureAdminGroup(loginName: string): Promise<void> {
    let groupExists: boolean = false;
    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(`${this.webUrl}/_api/web/sitegroups/getbyname('${this.odataString(this.adminGroupTitle)}')?$select=Id`, SPHttpClient.configurations.v1, { headers: { Accept: 'application/json;odata=nometadata' } });
      groupExists = response.ok;
    } catch { groupExists = false; }

    if (!groupExists) {
      const createResponse: SPHttpClientResponse = await this.context.spHttpClient.post(`${this.webUrl}/_api/web/sitegroups`, SPHttpClient.configurations.v1, { headers: { Accept: 'application/json;odata=verbose', 'Content-Type': 'application/json;odata=verbose' }, body: JSON.stringify({ __metadata: { type: 'SP.Group' }, Title: this.adminGroupTitle, Description: 'Usuarios autorizados para administrar la configuración de TIBOX HUB.' }) });
      if (!createResponse.ok) return;
    }

    try {
      await this.context.spHttpClient.post(`${this.webUrl}/_api/web/sitegroups/getbyname('${this.odataString(this.adminGroupTitle)}')/users`, SPHttpClient.configurations.v1, { headers: { Accept: 'application/json;odata=verbose', 'Content-Type': 'application/json;odata=verbose' }, body: JSON.stringify({ __metadata: { type: 'SP.User' }, LoginName: loginName }) });
    } catch { /* site admins retain access */ }
  }

  private async ensureConfigurationList(): Promise<void> {
    const existing: SPHttpClientResponse = await this.context.spHttpClient.get(`${this.webUrl}/_api/web/lists/getbytitle('${this.odataString(this.configListTitle)}')?$select=Id`, SPHttpClient.configurations.v1, { headers: { Accept: 'application/json;odata=nometadata' } });
    if (existing.ok) return;

    const createList: SPHttpClientResponse = await this.context.spHttpClient.post(`${this.webUrl}/_api/web/lists`, SPHttpClient.configurations.v1, { headers: { Accept: 'application/json;odata=verbose', 'Content-Type': 'application/json;odata=verbose' }, body: JSON.stringify({ __metadata: { type: 'SP.List' }, BaseTemplate: 100, Title: this.configListTitle, Description: 'Configuración interna de TIBOX HUB.', Hidden: true, ContentTypesEnabled: false }) });
    if (!createList.ok) return;

    await this.context.spHttpClient.post(`${this.webUrl}/_api/web/lists/getbytitle('${this.odataString(this.configListTitle)}')/fields/createfieldasxml`, SPHttpClient.configurations.v1, { headers: { Accept: 'application/json;odata=verbose', 'Content-Type': 'application/json;odata=verbose' }, body: JSON.stringify({ parameters: { __metadata: { type: 'SP.XmlSchemaFieldCreationInformation' }, SchemaXml: '<Field Type="Note" Name="ConfigValue" DisplayName="ConfigValue" NumLines="20" RichText="FALSE" />', Options: 0 } }) });
  }

  private async loadGlobalSettings(): Promise<void> {
    try {
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(`${this.webUrl}/_api/web/lists/getbytitle('${this.odataString(this.configListTitle)}')/items?$select=Id,Title,ConfigValue&$filter=Title eq '${this.odataString(this.quickLinksConfigKey)}'&$top=1`, SPHttpClient.configurations.v1, { headers: { Accept: 'application/json;odata=nometadata' } });
      if (!response.ok) return;
      const data: { value?: Array<{ ConfigValue?: string }> } = await response.json() as { value?: Array<{ ConfigValue?: string }> };
      const configValue: string | undefined = data.value?.[0]?.ConfigValue;
      if (!configValue) return;
      const parsed: QuickLink[] = JSON.parse(configValue) as QuickLink[];
      if (Array.isArray(parsed) && parsed.length) this.quickLinks = parsed;
    } catch { this.quickLinks = this.defaultQuickLinks(); }
  }

  private async saveGlobalSettings(): Promise<void> {
    await this.ensureConfigurationList();
    const listResponse: SPHttpClientResponse = await this.context.spHttpClient.get(`${this.webUrl}/_api/web/lists/getbytitle('${this.odataString(this.configListTitle)}')?$select=ListItemEntityTypeFullName`, SPHttpClient.configurations.v1, { headers: { Accept: 'application/json;odata=nometadata' } });
    if (!listResponse.ok) throw new Error('Config list unavailable');
    const listData: { ListItemEntityTypeFullName: string } = await listResponse.json() as { ListItemEntityTypeFullName: string };
    const itemResponse: SPHttpClientResponse = await this.context.spHttpClient.get(`${this.webUrl}/_api/web/lists/getbytitle('${this.odataString(this.configListTitle)}')/items?$select=Id&$filter=Title eq '${this.odataString(this.quickLinksConfigKey)}'&$top=1`, SPHttpClient.configurations.v1, { headers: { Accept: 'application/json;odata=nometadata' } });
    const itemData: { value?: Array<{ Id: number }> } = itemResponse.ok ? await itemResponse.json() as { value?: Array<{ Id: number }> } : {};
    const payload: Record<string, unknown> = { __metadata: { type: listData.ListItemEntityTypeFullName }, Title: this.quickLinksConfigKey, ConfigValue: JSON.stringify(this.quickLinks) };
    const existingId: number | undefined = itemData.value?.[0]?.Id;

    if (existingId) {
      const updateResponse: SPHttpClientResponse = await this.context.spHttpClient.post(`${this.webUrl}/_api/web/lists/getbytitle('${this.odataString(this.configListTitle)}')/items(${existingId})`, SPHttpClient.configurations.v1, { headers: { Accept: 'application/json;odata=verbose', 'Content-Type': 'application/json;odata=verbose', 'IF-MATCH': '*', 'X-HTTP-Method': 'MERGE' }, body: JSON.stringify(payload) });
      if (!updateResponse.ok) throw new Error('Could not update configuration');
      return;
    }

    const createResponse: SPHttpClientResponse = await this.context.spHttpClient.post(`${this.webUrl}/_api/web/lists/getbytitle('${this.odataString(this.configListTitle)}')/items`, SPHttpClient.configurations.v1, { headers: { Accept: 'application/json;odata=verbose', 'Content-Type': 'application/json;odata=verbose' }, body: JSON.stringify(payload) });
    if (!createResponse.ok) throw new Error('Could not create configuration');
  }

  private newHires(): HubPersonItem[] {
    const recent: HubPersonItem[] = this.data.people.filter((person: HubPersonItem) => person.FechaIngreso && this.daysFromToday(person.FechaIngreso) <= 60 && this.daysFromToday(person.FechaIngreso) >= 0);
    return recent.length ? recent : this.data.people.slice(0, 8);
  }

  private upcomingBirthdays(): HubPersonItem[] {
    return this.data.people.filter((person: HubPersonItem) => Boolean(person.Cumpleanos)).sort((a: HubPersonItem, b: HubPersonItem) => this.nextBirthday(a.Cumpleanos || '').getTime() - this.nextBirthday(b.Cumpleanos || '').getTime());
  }

  private futureEvents(): HubEventItem[] {
    const now: number = Date.now();
    return this.data.events.filter((event: HubEventItem) => !event.FechaInicio || new Date(event.FechaInicio).getTime() >= now - 86400000);
  }

  private learningCategories(): string {
    if (!this.data.learning.length) return this.empty('No hay contenido publicado.');
    const counts: Map<string, number> = new Map<string, number>();
    this.data.learning.forEach((item: HubLearningItem) => {
      const category: string = item.Categoria || 'Otros';
      counts.set(category, (counts.get(category) || 0) + 1);
    });
    return Array.from(counts.entries()).slice(0, 5).map(([category, count]: [string, number]) => `<div class="tbxLearningItem"><span><strong>${this.escapeHtml(category)}</strong><small>${count} ${count === 1 ? 'recurso' : 'recursos'}</small></span><span style="color:var(--support)">›</span></div>`).join('');
  }

  private personCard(person: HubPersonItem): string {
    return `<div class="tbxPerson" data-person="${this.escapeHtml(person.Title)}"><span class="tbxAvatar">${this.initials(person.Title)}</span><span class="tbxPersonText"><strong>${this.escapeHtml(person.Title)}</strong><span>${this.escapeHtml(person.Cargo || 'Sin cargo')} · ${this.escapeHtml(person.Area || 'Sin área')}</span><small>${person.FechaIngreso ? this.relativeDate(person.FechaIngreso) : ''}</small></span></div>`;
  }

  private movement(item: HubMovementItem): string {
    return `<div class="tbxMovement"><span class="tbxAvatar">${this.initials(item.Title)}</span><span class="tbxPersonText"><strong>${this.escapeHtml(item.Title)}</strong><span>${this.escapeHtml(item.CargoAnterior || '')} → <b style="color:var(--text)">${this.escapeHtml(item.CargoNuevo || '')}</b></span></span><span class="change">↗ ${this.escapeHtml(item.TipoMovimiento || 'Movimiento')}</span></div>`;
  }

  private course(item: HubCourseItem): string {
    const state: string = item.Estado || 'Activo';
    const stateClass: string = state.toLowerCase().includes('vence') ? 'warn' : state.toLowerCase().includes('nuevo') ? 'new' : 'blue';
    const deadline: string = item.FechaLimite ? `${this.formatDate(item.FechaLimite)} · ${this.daysRemaining(item.FechaLimite)}` : 'Sin fecha límite';
    const url: string = item.UrlCurso || '#';
    return `<article class="tbxCourse"><div class="tbxCourseTop"><span class="tbxPill ${stateClass}">${this.escapeHtml(state)}</span>${item.Obligatorio ? '<span class="tbxPill required">Obligatorio</span>' : ''}</div><h3>${this.escapeHtml(item.Title)}</h3><div class="tbxCourseMeta"><span>◷ ${this.duration(item.DuracionMinutos || 0)}</span></div><div class="tbxCourseMeta"><span>Fecha límite: ${deadline}</span></div><a class="tbxCourseAction" href="${this.escapeHtml(url)}" ${url.startsWith('http') ? 'target="_blank" rel="noopener noreferrer"' : ''}>Ir al curso</a></article>`;
  }

  private birthday(person: HubPersonItem): string {
    return `<div class="tbxBirthday"><span class="tbxAvatar">${this.initials(person.Title)}</span><span class="tbxPersonText"><strong>${this.escapeHtml(person.Title)}</strong><span>${this.escapeHtml(person.Area || '')} · ${this.birthdayLabel(person.Cumpleanos || '')}</span></span><button data-birthday="${this.escapeHtml(person.Title)}">Felicitar</button></div>`;
  }

  private event(item: HubEventItem): string {
    const date: Date = item.FechaInicio ? new Date(item.FechaInicio) : new Date();
    const day: string = String(date.getDate()).padStart(2, '0');
    const month: string = date.toLocaleDateString('es-CL', { month: 'short' }).replace('.', '').toUpperCase();
    const time: string = item.FechaInicio ? date.toLocaleTimeString('es-CL', { hour: '2-digit', minute: '2-digit' }) : '';
    return `<div class="tbxEvent"><span class="tbxDate">${day}<small>${month}</small></span><span class="tbxPersonText"><strong>${this.escapeHtml(item.Title)}</strong><span>${time} · ${this.escapeHtml(item.Modalidad || '')}${item.DuracionMinutos ? ` · ${this.duration(item.DuracionMinutos)}` : ''}</span></span></div>`;
  }

  private news(item: HubNewsItem): string {
    return `<article class="tbxNewsItem"><label>${this.escapeHtml(item.Categoria || 'Noticia')}</label><strong>${this.escapeHtml(item.Title)}</strong><small>${item.FechaPublicacion ? this.relativeDate(item.FechaPublicacion) : ''}</small></article>`;
  }

  private benefit(item: HubBenefitItem): string {
    return `<article class="tbxNewsItem" style="min-height:110px"><label>${this.escapeHtml((item.Categoria || 'BENEFICIO').toUpperCase())}</label><strong>${this.escapeHtml(item.Title)}</strong><small style="font-size:10px;line-height:1.5">${this.escapeHtml(item.Descripcion || '')}</small></article>`;
  }

  private listTitleFromKey(key: string): string | undefined {
    const map: Record<string, string> = {
      courses: TiboxHubSharePointService.lists.courses,
      events: TiboxHubSharePointService.lists.events,
      learning: TiboxHubSharePointService.lists.learning,
      news: TiboxHubSharePointService.lists.news
    };
    return map[key];
  }

  private formatDate(value: string): string { return new Date(value).toLocaleDateString('es-CL', { day: '2-digit', month: 'short' }); }
  private relativeDate(value: string): string {
    const days: number = this.daysFromToday(value);
    if (days <= 0) return 'Hoy';
    if (days === 1) return 'Hace 1 día';
    if (days < 7) return `Hace ${days} días`;
    if (days < 14) return 'Hace 1 semana';
    return `Hace ${Math.floor(days / 7)} semanas`;
  }
  private daysFromToday(value: string): number { return Math.max(0, Math.floor((Date.now() - new Date(value).getTime()) / 86400000)); }
  private daysRemaining(value: string): string {
    const days: number = Math.ceil((new Date(value).getTime() - Date.now()) / 86400000);
    if (days < 0) return 'Vencido';
    if (days === 0) return 'Vence hoy';
    return `${days} ${days === 1 ? 'día restante' : 'días restantes'}`;
  }
  private duration(minutes: number): string {
    if (minutes < 60) return `${minutes} min`;
    const hours: number = Math.floor(minutes / 60);
    const rest: number = minutes % 60;
    return rest ? `${hours} h ${rest} min` : `${hours} h`;
  }
  private nextBirthday(value: string): Date {
    const source: Date = new Date(value);
    const today: Date = new Date();
    let result: Date = new Date(today.getFullYear(), source.getUTCMonth(), source.getUTCDate());
    result.setHours(12, 0, 0, 0);
    if (result.getTime() < new Date(today.getFullYear(), today.getMonth(), today.getDate()).getTime()) result = new Date(today.getFullYear() + 1, source.getUTCMonth(), source.getUTCDate(), 12, 0, 0, 0);
    return result;
  }
  private birthdayLabel(value: string): string {
    const date: Date = this.nextBirthday(value);
    const today: Date = new Date();
    const days: number = Math.round((date.getTime() - new Date(today.getFullYear(), today.getMonth(), today.getDate(), 12, 0, 0, 0).getTime()) / 86400000);
    if (days === 0) return 'Hoy';
    if (days === 1) return 'Mañana';
    return date.toLocaleDateString('es-CL', { day: '2-digit', month: 'short' });
  }
  private initials(name: string): string { return name.split(/\s+/).filter(Boolean).slice(0, 2).map((part: string) => part.charAt(0).toUpperCase()).join(''); }
  private empty(message: string): string { return `<div class="tbxEmpty">${this.escapeHtml(message)}</div>`; }
  private modal(id: string, title: string, body: string, small: boolean): string { return `<div class="tbxOverlay" id="${id}"><section class="tbxModal ${small ? 'sm' : ''}"><header class="tbxModalHead"><h3>${title}</h3><button class="tbxClose" data-close aria-label="Cerrar">×</button></header><div class="tbxModalBody">${body}</div></section></div>`; }

  private get webUrl(): string { return this.context.pageContext.web.absoluteUrl.replace(/\/$/, ''); }
  private odataString(value: string): string { return value.replace(/'/g, "''"); }
  private emptyData(): HubDashboardData { return { people: [], movements: [], courses: [], benefits: [], events: [], learning: [], news: [] }; }

  private defaultQuickLinks(): QuickLink[] {
    return [
      { key: 'support', label: 'Soporte TI', url: '#', enabled: true },
      { key: 'expenses', label: 'Rindegastos', url: '#', enabled: true },
      { key: 'docs', label: 'Documentos', url: '#', enabled: true },
      { key: 'calendar', label: 'Calendario', url: '#', enabled: true },
      { key: 'benefits', label: 'Beneficios', url: '#beneficios', enabled: true },
      { key: 'directory', label: 'Directorio', url: '#personas', enabled: true },
      { key: 'webops', label: 'WebOps', url: '#', enabled: true }
    ];
  }

  private toast(message: string): void { const toast: HTMLElement | null = this.domElement.querySelector('[data-toast-box]'); if (!toast) return; toast.textContent = message; toast.classList.add('show'); window.setTimeout(() => toast.classList.remove('show'), 2600); }
  private quickIcon(key: string): string { const map: Record<string,string> = { support:'headset', expenses:'receipt', docs:'file', calendar:'calendar', benefits:'heart', directory:'users', webops:'globe' }; return map[key] || 'link'; }
  private icon(name: string): string {
    const paths: Record<string,string> = {
      search:'<circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/>',
      bell:'<path d="M6 8a6 6 0 0 1 12 0c0 7 3 7 3 9H3c0-2 3-2 3-9"/><path d="M10 20a2 2 0 0 0 4 0"/>',
      settings:'<circle cx="12" cy="12" r="3.5"/><path d="M12 2v3M12 19v3M4.9 4.9 7 7M17 17l2.1 2.1M2 12h3M19 12h3M4.9 19.1 7 17M17 7l2.1-2.1"/>',
      headset:'<path d="M4 12a8 8 0 0 1 16 0"/><rect x="2" y="12" width="4" height="6" rx="2"/><rect x="18" y="12" width="4" height="6" rx="2"/>',
      receipt:'<path d="M6 3h12v18l-2-1.5L14 21l-2-1.5L10 21l-2-1.5L6 21V3Z"/><path d="M9 8h6M9 12h6M9 16h4"/>',
      file:'<path d="M6 2h8l4 4v16H6z"/><path d="M14 2v5h5M9 12h6M9 16h6"/>',
      calendar:'<rect x="3" y="4" width="18" height="17" rx="3"/><path d="M7 2v4M17 2v4M3 9h18"/>',
      heart:'<path d="M12 20s-7-4.4-9-8.4A5.4 5.4 0 0 1 12 5a5.4 5.4 0 0 1 9 6.6C19 15.6 12 20 12 20Z"/>',
      users:'<circle cx="9" cy="8" r="3"/><path d="M3 20a6 6 0 0 1 12 0M16 5a3 3 0 0 1 0 6M18 15a5 5 0 0 1 3 5"/>',
      globe:'<circle cx="12" cy="12" r="9"/><path d="M3 12h18M12 3a14 14 0 0 1 0 18M12 3a14 14 0 0 0 0 18"/>',
      link:'<path d="M10 13a5 5 0 0 0 7.5.5l2-2a5 5 0 0 0-7-7l-1 1"/><path d="M14 11a5 5 0 0 0-7.5-.5l-2 2a5 5 0 0 0 7 7l1-1"/>'
    };
    return `<svg class="tbxIcon" viewBox="0 0 24 24" aria-hidden="true">${paths[name] || paths.link}</svg>`;
  }
  private escapeHtml(value: string): string { return value.replace(/[&<>'"]/g, (char: string) => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', "'":'&#39;', '"':'&quot;' }[char] || char)); }

  protected onDispose(): void { document.removeEventListener('keydown', this.onDocumentKeyDown); }
  protected get dataVersion(): Version { return Version.parse('1.0'); }
}
