import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

export interface IVgmDemoWebPartProps {}

type QuickLink = {
  key: string;
  label: string;
  url: string;
  enabled: boolean;
};

type Person = {
  initials: string;
  name: string;
  role: string;
  area: string;
  date: string;
};

export default class VgmDemoWebPart extends BaseClientSideWebPart<IVgmDemoWebPartProps> {
  private readonly storageKey: string = 'tiboxHubPrototypeSettingsV1';

  private quickLinks: QuickLink[] = [
    { key: 'support', label: 'Soporte TI', url: '#', enabled: true },
    { key: 'expenses', label: 'Rindegastos', url: '#', enabled: true },
    { key: 'docs', label: 'Documentos', url: '#', enabled: true },
    { key: 'calendar', label: 'Calendario', url: '#', enabled: true },
    { key: 'benefits', label: 'Beneficios', url: '#beneficios', enabled: true },
    { key: 'directory', label: 'Directorio', url: '#personas', enabled: true },
    { key: 'webops', label: 'WebOps', url: '#', enabled: true }
  ];

  private readonly hires: Person[] = [
    { initials: 'CR', name: 'Camila Rojas', role: 'Product Designer', area: 'Diseño', date: 'Hace 5 días' },
    { initials: 'MS', name: 'Matías Silva', role: 'Cloud Engineer', area: 'Infraestructura', date: 'Hace 5 días' },
    { initials: 'FC', name: 'Fernanda Castro', role: 'Data Analyst', area: 'Datos', date: 'Hace 1 semana' },
    { initials: 'DF', name: 'Diego Fuentes', role: 'QA Engineer', area: 'Calidad', date: 'Hace 2 semanas' },
    { initials: 'AM', name: 'Andrea Muñoz', role: 'Account Manager', area: 'Comercial', date: 'Hace 3 semanas' },
    { initials: 'JP', name: 'Joaquín Pérez', role: 'Support Engineer', area: 'Soporte', date: 'Hace 3 semanas' }
  ];

  public render(): void {
    this.loadSettings();
    const userName: string = this.context.pageContext.user.displayName || 'Equipo Tibox';

    this.domElement.innerHTML = `
      <style>
        .tbxHub, .tbxHub * { box-sizing: border-box; }
        .tbxHub {
          --bg:#000310; --surface:#0A1130; --surface2:#121A40;
          --border:rgba(255,255,255,.08); --borderAccent:rgba(0,209,255,.28);
          --text:#F4F7FF; --muted:#9BA6C4; --subtle:#5E6A8A;
          --support:#00D1FF; --supportSolid:#0E9CDC; --cta:#FF4222; --yellow:#FFB200;
          --green:#55D9A6; --red:#FF725E; --violet:#B9ABFF;
          width:100%; min-height:100vh; color:var(--muted); background:
            radial-gradient(circle at 78% -8%, rgba(0,209,255,.10), transparent 34rem),
            var(--bg);
          font-family: 'Segoe UI', Arial, sans-serif; font-size:14px;
          border-radius:18px; overflow:hidden;
        }
        .tbxHub button, .tbxHub input { font:inherit; }
        .tbxShell { max-width:1440px; margin:0 auto; padding:0 28px 54px; }
        .tbxTop {
          position:sticky; top:0; z-index:20; height:72px; margin:0 -28px; padding:0 28px;
          display:flex; align-items:center; justify-content:space-between;
          background:rgba(0,3,16,.84); backdrop-filter:blur(16px); border-bottom:1px solid var(--border);
        }
        .tbxBrand { display:flex; align-items:center; gap:11px; color:var(--text); font-weight:800; letter-spacing:.02em; }
        .tbxMark { width:30px; height:30px; border-radius:9px; background:linear-gradient(135deg,#00D1FF,#0E9CDC); box-shadow:0 0 0 5px rgba(0,209,255,.07); }
        .tbxBrand small { display:block; margin-top:2px; color:var(--subtle); font-weight:500; font-size:10px; }
        .tbxNav { display:flex; align-items:center; gap:6px; padding:4px; border:1px solid var(--border); background:rgba(18,26,64,.55); border-radius:12px; }
        .tbxNav button { border:0; color:var(--muted); background:transparent; padding:8px 13px; border-radius:9px; cursor:pointer; font-weight:600; font-size:12px; }
        .tbxNav button.active { background:var(--surface2); color:var(--text); box-shadow:0 1px 10px rgba(0,0,0,.22); }
        .tbxTopActions { display:flex; gap:8px; }
        .tbxIconBtn { width:38px; height:38px; display:grid; place-items:center; border:1px solid var(--border); border-radius:10px; color:var(--text); background:var(--surface); cursor:pointer; }
        .tbxIconBtn:hover { border-color:var(--borderAccent); background:var(--surface2); }
        .tbxIcon { width:18px; height:18px; stroke:currentColor; fill:none; stroke-width:1.8; stroke-linecap:round; stroke-linejoin:round; }
        .tbxHero { padding:44px 0 22px; }
        .tbxHero h1 { margin:0; color:var(--text); font-size:38px; line-height:1.08; letter-spacing:-.025em; }
        .tbxHero p { margin:9px 0 0; font-size:15px; }
        .tbxSearch { margin-top:25px; width:min(100%,650px); height:46px; display:flex; align-items:center; gap:10px; padding:0 14px; border:1px solid var(--border); border-radius:12px; background:var(--surface); }
        .tbxSearch input { flex:1; border:0; outline:0; background:transparent; color:var(--text); }
        .tbxSearch input::placeholder { color:var(--subtle); }
        .tbxQuick { display:grid; grid-template-columns:repeat(7,minmax(0,1fr)); gap:10px; margin:12px 0 26px; }
        .tbxQuick a { min-height:74px; padding:13px; display:flex; flex-direction:column; justify-content:space-between; border:1px solid var(--border); border-radius:13px; background:var(--surface); color:var(--text); text-decoration:none; font-size:12px; font-weight:650; transition:.15s; }
        .tbxQuick a:hover { transform:translateY(-1px); border-color:var(--borderAccent); background:var(--surface2); }
        .tbxQuickIcon { width:29px; height:29px; border-radius:9px; display:grid; place-items:center; color:var(--support); background:rgba(0,209,255,.08); }
        .tbxLayout { display:grid; grid-template-columns:minmax(0,1fr) 300px; gap:18px; }
        .tbxMain { min-width:0; display:flex; flex-direction:column; gap:18px; }
        .tbxSide { display:flex; flex-direction:column; gap:18px; }
        .tbxCard { border:1px solid var(--border); border-radius:18px; background:var(--surface); box-shadow:0 12px 30px rgba(0,0,0,.10); overflow:hidden; }
        .tbxCardHead { min-height:54px; padding:14px 17px; display:flex; align-items:center; justify-content:space-between; gap:12px; border-bottom:1px solid var(--border); }
        .tbxCardHead h2, .tbxCardHead h3 { margin:0; color:var(--text); font-size:15px; }
        .tbxLink { border:0; background:transparent; color:var(--support); font-weight:700; font-size:11px; cursor:pointer; }
        .tbxPeopleGrid { padding:12px; display:grid; grid-template-columns:1fr 1fr; gap:9px; }
        .tbxPerson { padding:11px; min-width:0; display:flex; align-items:center; gap:11px; border:1px solid var(--border); border-radius:12px; background:var(--surface2); cursor:pointer; }
        .tbxPerson:hover { border-color:var(--borderAccent); }
        .tbxAvatar { width:38px; height:38px; flex:none; border-radius:50%; display:grid; place-items:center; color:#001233; background:linear-gradient(135deg,#00D1FF,#88E8FF); font-weight:800; font-size:11px; }
        .tbxPersonText { min-width:0; flex:1; }
        .tbxPersonText strong { display:block; color:var(--text); font-size:12px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
        .tbxPersonText span { display:block; margin-top:2px; color:var(--muted); font-size:10px; }
        .tbxPersonText small { color:var(--support); font-size:9px; }
        .tbxMovements { padding:10px 12px 12px; }
        .tbxMovement { display:flex; align-items:center; gap:11px; padding:11px 7px; border-bottom:1px solid var(--border); }
        .tbxMovement:last-child { border-bottom:0; }
        .tbxMovement .change { margin-left:auto; padding:4px 8px; border-radius:999px; font-size:9px; font-weight:700; color:var(--green); border:1px solid rgba(85,217,166,.20); background:rgba(85,217,166,.08); }
        .tbxCourses { padding:13px; display:grid; grid-template-columns:repeat(3,1fr); gap:10px; }
        .tbxCourse { padding:14px; border:1px solid var(--border); border-radius:13px; background:var(--surface2); }
        .tbxCourseTop { display:flex; justify-content:space-between; gap:7px; }
        .tbxPill { display:inline-flex; padding:4px 7px; border-radius:999px; font-size:9px; font-weight:800; }
        .tbxPill.warn { color:var(--yellow); background:rgba(255,178,0,.09); }
        .tbxPill.blue { color:var(--support); background:rgba(0,209,255,.08); }
        .tbxPill.new { color:var(--violet); background:rgba(185,171,255,.09); }
        .tbxPill.required { color:var(--red); background:rgba(255,114,94,.08); }
        .tbxCourse h3 { color:var(--text); margin:12px 0 6px; font-size:12px; }
        .tbxCourseMeta { display:flex; gap:12px; color:var(--muted); font-size:10px; line-height:1.5; }
        .tbxCourse button { margin-top:12px; border:0; border-radius:8px; padding:8px 10px; background:var(--supportSolid); color:white; font-size:10px; font-weight:800; cursor:pointer; }
        .tbxCourseFooter { margin:0 13px 13px; padding:11px 0 0; border-top:1px solid var(--border); display:flex; flex-wrap:wrap; gap:7px; }
        .tbxStatus { padding:5px 8px; border-radius:999px; background:var(--surface2); font-size:9px; }
        .tbxLearning { padding:13px; display:grid; grid-template-columns:1fr 1fr; gap:12px; }
        .tbxPrompt { min-height:160px; padding:18px; border-radius:14px; color:var(--text); background:linear-gradient(145deg,#121A40,#0A1130); border:1px solid var(--borderAccent); }
        .tbxPrompt label { color:var(--support); font-size:9px; font-weight:800; letter-spacing:.08em; }
        .tbxPrompt p { margin:14px 0; color:var(--text); font-style:italic; line-height:1.55; font-size:12px; }
        .tbxPrompt button { border:1px solid var(--border); border-radius:8px; padding:8px 10px; background:var(--text); color:#001233; font-weight:800; font-size:10px; cursor:pointer; }
        .tbxLearningList { display:flex; flex-direction:column; gap:7px; }
        .tbxLearningItem { padding:11px 12px; display:flex; align-items:center; justify-content:space-between; border:1px solid var(--border); border-radius:11px; background:var(--surface2); color:var(--text); }
        .tbxLearningItem small { display:block; margin-top:2px; color:var(--subtle); }
        .tbxNews { padding:12px; display:grid; grid-template-columns:1fr 1fr; gap:9px; }
        .tbxNewsItem { min-height:82px; padding:13px; border:1px solid var(--border); border-radius:12px; background:var(--surface2); }
        .tbxNewsItem label { font-size:9px; color:var(--support); }
        .tbxNewsItem strong { display:block; margin-top:8px; color:var(--text); font-size:11px; }
        .tbxNewsItem small { display:block; margin-top:5px; color:var(--subtle); font-size:9px; }
        .tbxBirthdayList, .tbxEventList { padding:10px; display:flex; flex-direction:column; gap:8px; }
        .tbxBirthday, .tbxEvent { display:flex; align-items:center; gap:9px; padding:10px; border:1px solid var(--border); border-radius:11px; background:var(--surface2); }
        .tbxBirthday .tbxAvatar { width:34px; height:34px; }
        .tbxBirthday button { margin-left:auto; border:1px solid rgba(255,178,0,.20); border-radius:8px; background:rgba(255,178,0,.09); color:var(--yellow); padding:6px 8px; font-weight:800; font-size:9px; cursor:pointer; }
        .tbxDate { width:37px; height:42px; flex:none; border-radius:9px; display:grid; place-items:center; background:#050B26; color:var(--support); font-weight:800; line-height:1.05; font-size:13px; }
        .tbxDate small { display:block; font-size:7px; }
        .tbxBenefitHero { min-height:155px; padding:18px; background:radial-gradient(circle at 100% 0,rgba(0,209,255,.18),transparent 170px),#050B26; }
        .tbxBenefitHero label { color:var(--support); font-size:9px; font-weight:800; }
        .tbxBenefitHero h3 { color:var(--text); margin:12px 0 7px; font-size:15px; }
        .tbxBenefitHero p { margin:0; font-size:10px; line-height:1.55; }
        .tbxBenefitHero button { margin-top:14px; border:0; border-radius:9px; padding:9px 12px; background:var(--cta); color:#fff; font-weight:800; font-size:10px; cursor:pointer; }
        .tbxOverlay { position:fixed; inset:0; z-index:1000; display:none; align-items:center; justify-content:center; padding:28px; background:rgba(0,3,16,.72); backdrop-filter:blur(7px); }
        .tbxOverlay.open { display:flex; }
        .tbxModal { width:min(920px,96vw); max-height:86vh; overflow:auto; border:1px solid var(--border); border-radius:20px; background:var(--surface); box-shadow:0 30px 100px rgba(0,0,0,.46); }
        .tbxModal.sm { width:min(560px,96vw); }
        .tbxModalHead { position:sticky; top:0; z-index:2; min-height:62px; padding:15px 18px; display:flex; align-items:center; justify-content:space-between; background:rgba(10,17,48,.96); border-bottom:1px solid var(--border); }
        .tbxModalHead h3 { margin:0; color:var(--text); }
        .tbxClose { width:34px; height:34px; border:1px solid var(--border); border-radius:9px; background:var(--surface2); color:var(--text); cursor:pointer; }
        .tbxModalBody { padding:18px; }
        .tbxModalSearch { width:100%; height:42px; padding:0 12px; border:1px solid var(--border); border-radius:10px; background:var(--surface2); color:var(--text); outline:0; }
        .tbxAllHires { margin-top:14px; display:grid; grid-template-columns:1fr 1fr; gap:9px; }
        .tbxTextarea { width:100%; min-height:110px; padding:12px; resize:vertical; border:1px solid var(--border); border-radius:11px; background:var(--surface2); color:var(--text); outline:0; }
        .tbxModalActions { display:flex; justify-content:flex-end; gap:8px; margin-top:14px; }
        .tbxBtn { border:1px solid var(--border); border-radius:9px; padding:9px 12px; background:var(--surface2); color:var(--text); font-weight:700; cursor:pointer; }
        .tbxBtn.primary { background:var(--cta); border-color:var(--cta); color:#fff; }
        .tbxSettingsGrid { display:grid; grid-template-columns:1fr; gap:10px; }
        .tbxSettingRow { padding:12px; display:grid; grid-template-columns:34px 1fr 1.2fr 44px; align-items:center; gap:9px; border:1px solid var(--border); border-radius:11px; background:var(--surface2); }
        .tbxSettingRow input[type='text'] { width:100%; height:34px; padding:0 9px; border:1px solid var(--border); border-radius:8px; background:var(--surface); color:var(--text); }
        .tbxSettingRow input[type='checkbox'] { width:18px; height:18px; accent-color:var(--supportSolid); }
        .tbxToast { position:fixed; z-index:1200; right:24px; bottom:24px; max-width:340px; padding:12px 15px; border:1px solid var(--borderAccent); border-radius:11px; background:#0A1130; color:var(--text); box-shadow:0 18px 50px rgba(0,0,0,.38); opacity:0; transform:translateY(8px); pointer-events:none; transition:.2s; }
        .tbxToast.show { opacity:1; transform:none; }
        @media(max-width:1100px){ .tbxQuick{grid-template-columns:repeat(4,1fr)} .tbxLayout{grid-template-columns:1fr} .tbxSide{display:grid;grid-template-columns:1fr 1fr} }
        @media(max-width:760px){ .tbxShell{padding:0 14px 34px}.tbxTop{margin:0 -14px;padding:0 14px}.tbxNav{display:none}.tbxHero h1{font-size:30px}.tbxQuick{grid-template-columns:repeat(2,1fr)}.tbxPeopleGrid,.tbxCourses,.tbxLearning,.tbxNews,.tbxAllHires{grid-template-columns:1fr}.tbxSide{display:flex}.tbxSettingRow{grid-template-columns:30px 1fr}.tbxSettingRow input[type='text']{grid-column:1/-1}.tbxSettingRow input[type='checkbox']{grid-column:2} }
      </style>

      <div class="tbxHub">
        <div class="tbxShell">
          <header class="tbxTop">
            <div class="tbxBrand"><span class="tbxMark"></span><div>TIBOX HUB<small>Todo Tibox, en un solo lugar.</small></div></div>
            <nav class="tbxNav"><button class="active">Inicio</button><button data-scroll="aprende">Tibox Aprende</button><button data-open="settings">Administrar</button></nav>
            <div class="tbxTopActions">
              <button class="tbxIconBtn" title="Notificaciones" aria-label="Notificaciones">${this.icon('bell')}</button>
              <button class="tbxIconBtn" data-open="settings" title="Configuración" aria-label="Configuración">${this.icon('settings')}</button>
            </div>
          </header>

          <section class="tbxHero">
            <h1>Hola, ${this.escapeHtml(userName)}</h1>
            <p>Esto está pasando en Tibox.</p>
            <label class="tbxSearch">${this.icon('search')}<input data-global-search type="search" placeholder="Buscar personas, cursos, beneficios, documentos..." /></label>
          </section>

          <section class="tbxQuick" data-quick-links></section>

          <div class="tbxLayout">
            <main class="tbxMain">
              <section class="tbxCard" id="personas">
                <div class="tbxCardHead"><h2>Nuevos en Tibox</h2><button class="tbxLink" data-open="hires">Ver todos los nuevos ingresos ›</button></div>
                <div class="tbxPeopleGrid">${this.hires.slice(0,4).map(p => this.personCard(p)).join('')}</div>
              </section>

              <section class="tbxCard">
                <div class="tbxCardHead"><h2>Movimientos en Tibox</h2></div>
                <div class="tbxMovements">
                  ${this.movement('BC','Braulio Contreras','Developer','Senior Developer','Ascenso')}
                  ${this.movement('JM','Javier Morales','Product Designer','AI & Design Specialist','Cambio de rol')}
                  ${this.movement('DH','Daniela Herrera','Analista de Proyectos','Coordinadora de Proyectos','Ascenso')}
                </div>
              </section>

              <section class="tbxCard" id="cursos">
                <div class="tbxCardHead"><h2>Cursos del mes</h2><button class="tbxLink" data-toast="Vista completa de cursos: se conectará a la lista SharePoint.">Ver todos los cursos ›</button></div>
                <div class="tbxCourses">
                  ${this.course('Vence pronto','warn','Obligatorio','Prevención del acoso laboral','45 min','08 septiembre · 7 días restantes')}
                  ${this.course('Activo','blue','Obligatorio','Seguridad de la información','1 h 20 min','16 septiembre · 15 días restantes')}
                  ${this.course('Nuevo','new','','Introducción a la Ley de Datos','35 min','30 septiembre · 29 días restantes')}
                </div>
                <div class="tbxCourseFooter"><span class="tbxStatus" style="color:var(--green)">✓ Prevención del acoso laboral · Completado</span><span class="tbxStatus">○ Seguridad de la información · Pendiente</span><span class="tbxStatus" style="color:var(--red)">! Introducción a la Ley de Datos · Próximo a vencer</span></div>
              </section>

              <section class="tbxCard" id="aprende">
                <div class="tbxCardHead"><h2>Tibox Aprende</h2><button class="tbxLink" data-toast="Aquí podremos abrir el centro completo de conocimiento.">Explorar ›</button></div>
                <div class="tbxLearning">
                  <div class="tbxPrompt"><label>✦ PROMPT DE LA SEMANA</label><p>“Analiza este correo, identifica las acciones que debo realizar, ordénalas por prioridad y redacta una respuesta breve y profesional.”</p><button data-copy-prompt>Copiar prompt</button></div>
                  <div class="tbxLearningList">
                    ${this.learningItem('Inteligencia Artificial','8 recursos')}
                    ${this.learningItem('Microsoft 365','4 recursos')}
                    ${this.learningItem('Seguridad','4 recursos')}
                    ${this.learningItem('Productividad','4 recursos')}
                  </div>
                </div>
              </section>

              <section class="tbxCard">
                <div class="tbxCardHead"><h2>Tibox informa</h2></div>
                <div class="tbxNews">
                  ${this.news('Nuevo cliente','Tibox certifica a nuevo cliente del retail','Hace 2 días')}
                  ${this.news('Proyecto terminado','Cerramos con éxito el proyecto Aurora','Hace 4 días')}
                  ${this.news('Nuevo colaborador','Le damos la bienvenida al nuevo equipo de Datos','Hace 5 días')}
                  ${this.news('Comunicado','Actualización de la política de teletrabajo','Hace 1 semana')}
                </div>
              </section>
            </main>

            <aside class="tbxSide">
              <section class="tbxCard">
                <div class="tbxCardHead"><h3>Próximos cumpleaños</h3></div>
                <div class="tbxBirthdayList">
                  ${this.birthday('PF','Paula Farías','Marketing · Hoy')}
                  ${this.birthday('SR','Sebastián Rivas','Infraestructura · Mañana')}
                  ${this.birthday('AL','Antonia López','Finanzas · 05 sep')}
                  ${this.birthday('RP','Rodrigo Paredes','Producto · 09 sep')}
                </div>
              </section>

              <section class="tbxCard">
                <div class="tbxCardHead"><h3>Próximos eventos</h3></div>
                <div class="tbxEventList">
                  ${this.event('04','SEP','Workshop Inteligencia Artificial','15:00 · Teams · 60 min')}
                  ${this.event('11','SEP','Encuentro Tibox','17:30 · Presencial · 90 min')}
                  ${this.event('18','SEP','Capacitación Excel Avanzado','10:00 · Teams · 45 min')}
                </div>
              </section>

              <section class="tbxCard" id="beneficios">
                <div class="tbxBenefitHero"><label>♡ BENEFICIOS</label><h3>Beneficios pensados para ti</h3><p>Salud, educación y convenios comerciales disponibles para todo el equipo Tibox.</p><button data-open="benefits">Ver beneficios</button></div>
              </section>
            </aside>
          </div>
        </div>

        ${this.modal('hiresModal','Nuevos en Tibox',`<input class="tbxModalSearch" data-hire-search placeholder="Buscar por nombre, cargo o área..." /><div class="tbxAllHires" data-all-hires>${this.hires.map(p => this.personCard(p)).join('')}</div>`,false)}
        ${this.modal('birthdayModal','Felicitar cumpleaños',`<p data-birthday-copy style="margin:0 0 12px">Escribe tu saludo. En la versión integrada podremos abrir Teams con el mensaje preparado.</p><textarea class="tbxTextarea" data-birthday-message></textarea><div class="tbxModalActions"><button class="tbxBtn" data-close>Cancelar</button><button class="tbxBtn primary" data-send-teams>Abrir en Teams</button></div>`,true)}
        ${this.modal('benefitsModal','Beneficios Tibox',`<div class="tbxAllHires">${this.benefit('Giftcard de cumpleaños','Beneficio entregado durante el mes de cumpleaños.')}${this.benefit('Convenio de salud','Condiciones preferenciales para colaboradores.')}${this.benefit('Capacitaciones','Acceso a instancias seleccionadas de formación.')}${this.benefit('Convenios comerciales','Descuentos y convenios disponibles para el equipo.')}</div>`,false)}
        ${this.modal('settingsModal','Configuración de TIBOX HUB',`<p style="margin:0 0 14px">Configura los accesos rápidos del prototipo. En la versión real esta configuración vivirá en SharePoint.</p><div class="tbxSettingsGrid" data-settings-grid></div><div class="tbxModalActions"><button class="tbxBtn" data-reset-settings>Restaurar</button><button class="tbxBtn primary" data-save-settings>Guardar cambios</button></div>`,false)}
        <div class="tbxToast" data-toast-box></div>
      </div>
    `;

    this.renderQuickLinks();
    this.renderSettings();
    this.bindEvents();
  }

  private bindEvents(): void {
    this.domElement.querySelectorAll<HTMLElement>('[data-open]').forEach((el: HTMLElement) => {
      el.addEventListener('click', () => this.openModal(el.getAttribute('data-open') || ''));
    });

    this.domElement.querySelectorAll<HTMLElement>('[data-close]').forEach((el: HTMLElement) => {
      el.addEventListener('click', () => this.closeAllModals());
    });

    this.domElement.querySelectorAll<HTMLElement>('.tbxOverlay').forEach((overlay: HTMLElement) => {
      overlay.addEventListener('click', (event: MouseEvent) => {
        if (event.target === overlay) this.closeAllModals();
      });
    });

    this.domElement.querySelectorAll<HTMLElement>('[data-toast]').forEach((el: HTMLElement) => {
      el.addEventListener('click', () => this.toast(el.getAttribute('data-toast') || 'Listo'));
    });

    this.domElement.querySelectorAll<HTMLElement>('[data-scroll]').forEach((el: HTMLElement) => {
      el.addEventListener('click', () => {
        const target: HTMLElement | null = this.domElement.querySelector(`#${el.getAttribute('data-scroll')}`);
        target?.scrollIntoView({ behavior: 'smooth', block: 'start' });
      });
    });

    this.domElement.querySelectorAll<HTMLElement>('[data-person]').forEach((el: HTMLElement) => {
      el.addEventListener('click', () => this.toast(`${el.getAttribute('data-person')}: el perfil completo se conectará a Microsoft 365.`));
    });

    this.domElement.querySelectorAll<HTMLElement>('[data-birthday]').forEach((el: HTMLElement) => {
      el.addEventListener('click', () => {
        const name: string = el.getAttribute('data-birthday') || '';
        const firstName: string = name.split(' ')[0];
        const area: HTMLTextAreaElement | null = this.domElement.querySelector('[data-birthday-message]');
        if (area) area.value = `¡Feliz cumpleaños, ${firstName}! 🎉 Que tengas un excelente día y un gran año.`;
        this.openModal('birthday');
      });
    });

    const hireSearch: HTMLInputElement | null = this.domElement.querySelector('[data-hire-search]');
    hireSearch?.addEventListener('input', () => this.filterHires(hireSearch.value));

    const globalSearch: HTMLInputElement | null = this.domElement.querySelector('[data-global-search]');
    globalSearch?.addEventListener('keydown', (event: KeyboardEvent) => {
      if (event.key === 'Enter' && globalSearch.value.trim()) this.toast(`Búsqueda: “${globalSearch.value.trim()}” (demo)`);
    });

    const copy: HTMLElement | null = this.domElement.querySelector('[data-copy-prompt]');
    copy?.addEventListener('click', async () => {
      const prompt: string = 'Analiza este correo, identifica las acciones que debo realizar, ordénalas por prioridad y redacta una respuesta breve y profesional.';
      try { await navigator.clipboard.writeText(prompt); this.toast('Prompt copiado'); }
      catch { this.toast('No se pudo usar el portapapeles en este contexto.'); }
    });

    const sendTeams: HTMLElement | null = this.domElement.querySelector('[data-send-teams]');
    sendTeams?.addEventListener('click', () => this.toast('En la integración real abriremos Teams con el mensaje preparado.'));

    const save: HTMLElement | null = this.domElement.querySelector('[data-save-settings]');
    save?.addEventListener('click', () => this.saveSettingsFromForm());

    const reset: HTMLElement | null = this.domElement.querySelector('[data-reset-settings]');
    reset?.addEventListener('click', () => {
      try { localStorage.removeItem(this.storageKey); } catch { /* no-op */ }
      this.quickLinks = this.defaultQuickLinks();
      this.renderQuickLinks();
      this.renderSettings();
      this.toast('Configuración restaurada');
    });

    document.addEventListener('keydown', this.onDocumentKeyDown);
  }

  private readonly onDocumentKeyDown = (event: KeyboardEvent): void => {
    if (event.key === 'Escape') this.closeAllModals();
  };

  private openModal(name: string): void {
    const idMap: Record<string,string> = { hires:'hiresModal', birthday:'birthdayModal', benefits:'benefitsModal', settings:'settingsModal' };
    const modal: HTMLElement | null = this.domElement.querySelector(`#${idMap[name] || name}`);
    modal?.classList.add('open');
  }

  private closeAllModals(): void {
    this.domElement.querySelectorAll<HTMLElement>('.tbxOverlay.open').forEach((el: HTMLElement) => el.classList.remove('open'));
  }

  private filterHires(value: string): void {
    const query: string = value.trim().toLowerCase();
    const area: HTMLElement | null = this.domElement.querySelector('[data-all-hires]');
    if (!area) return;
    const filtered: Person[] = this.hires.filter((p: Person) => `${p.name} ${p.role} ${p.area}`.toLowerCase().includes(query));
    area.innerHTML = filtered.length ? filtered.map((p: Person) => this.personCard(p)).join('') : '<p>No se encontraron personas.</p>';
    area.querySelectorAll<HTMLElement>('[data-person]').forEach((el: HTMLElement) => el.addEventListener('click', () => this.toast(`${el.getAttribute('data-person')}: perfil demo.`)));
  }

  private renderQuickLinks(): void {
    const area: HTMLElement | null = this.domElement.querySelector('[data-quick-links]');
    if (!area) return;
    area.innerHTML = this.quickLinks.filter((link: QuickLink) => link.enabled).map((link: QuickLink) => `
      <a href="${this.escapeHtml(link.url || '#')}" ${link.url.startsWith('http') ? 'target="_blank" rel="noopener noreferrer"' : ''}>
        <span class="tbxQuickIcon">${this.icon(this.quickIcon(link.key))}</span><span>${this.escapeHtml(link.label)}</span>
      </a>`).join('');
  }

  private renderSettings(): void {
    const area: HTMLElement | null = this.domElement.querySelector('[data-settings-grid]');
    if (!area) return;
    area.innerHTML = this.quickLinks.map((link: QuickLink, index: number) => `
      <div class="tbxSettingRow" data-setting-index="${index}">
        <span class="tbxQuickIcon">${this.icon(this.quickIcon(link.key))}</span>
        <input data-label type="text" value="${this.escapeHtml(link.label)}" aria-label="Nombre" />
        <input data-url type="text" value="${this.escapeHtml(link.url)}" aria-label="URL" />
        <input data-enabled type="checkbox" ${link.enabled ? 'checked' : ''} aria-label="Activo" />
      </div>`).join('');
  }

  private saveSettingsFromForm(): void {
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
    try { localStorage.setItem(this.storageKey, JSON.stringify(this.quickLinks)); } catch { /* storage may be disabled */ }
    this.renderQuickLinks();
    this.closeAllModals();
    this.toast('Accesos rápidos guardados');
  }

  private loadSettings(): void {
    try {
      const saved: string | null = localStorage.getItem(this.storageKey);
      if (saved) {
        const parsed: QuickLink[] = JSON.parse(saved) as QuickLink[];
        if (Array.isArray(parsed) && parsed.length) this.quickLinks = parsed;
      }
    } catch { this.quickLinks = this.defaultQuickLinks(); }
  }

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

  private toast(message: string): void {
    const toast: HTMLElement | null = this.domElement.querySelector('[data-toast-box]');
    if (!toast) return;
    toast.textContent = message;
    toast.classList.add('show');
    window.setTimeout(() => toast.classList.remove('show'), 2600);
  }

  private personCard(person: Person): string {
    return `<div class="tbxPerson" data-person="${this.escapeHtml(person.name)}"><span class="tbxAvatar">${person.initials}</span><span class="tbxPersonText"><strong>${this.escapeHtml(person.name)}</strong><span>${this.escapeHtml(person.role)} · ${this.escapeHtml(person.area)}</span><small>${this.escapeHtml(person.date)}</small></span></div>`;
  }

  private movement(initials: string, name: string, before: string, after: string, type: string): string {
    return `<div class="tbxMovement"><span class="tbxAvatar">${initials}</span><span class="tbxPersonText"><strong>${name}</strong><span>${before} → <b style="color:var(--text)">${after}</b></span></span><span class="change">↗ ${type}</span></div>`;
  }

  private course(state: string, stateClass: string, required: string, title: string, duration: string, deadline: string): string {
    return `<article class="tbxCourse"><div class="tbxCourseTop"><span class="tbxPill ${stateClass}">${state}</span>${required ? `<span class="tbxPill required">${required}</span>` : ''}</div><h3>${title}</h3><div class="tbxCourseMeta"><span>◷ ${duration}</span></div><div class="tbxCourseMeta"><span>Fecha límite: ${deadline}</span></div><button data-toast="Abrir curso: ${title}">Ir al curso</button></article>`;
  }

  private learningItem(title: string, count: string): string {
    return `<div class="tbxLearningItem"><span><strong>${title}</strong><small>${count}</small></span><span style="color:var(--support)">›</span></div>`;
  }

  private news(tag: string, title: string, time: string): string {
    return `<article class="tbxNewsItem"><label>${tag}</label><strong>${title}</strong><small>${time}</small></article>`;
  }

  private birthday(initials: string, name: string, meta: string): string {
    return `<div class="tbxBirthday"><span class="tbxAvatar">${initials}</span><span class="tbxPersonText"><strong>${name}</strong><span>${meta}</span></span><button data-birthday="${name}">Felicitar</button></div>`;
  }

  private event(day: string, month: string, title: string, meta: string): string {
    return `<div class="tbxEvent"><span class="tbxDate">${day}<small>${month}</small></span><span class="tbxPersonText"><strong>${title}</strong><span>${meta}</span></span></div>`;
  }

  private benefit(title: string, text: string): string {
    return `<article class="tbxNewsItem" style="min-height:110px"><label>BENEFICIO</label><strong>${title}</strong><small style="font-size:10px;line-height:1.5">${text}</small></article>`;
  }

  private modal(id: string, title: string, body: string, small: boolean): string {
    return `<div class="tbxOverlay" id="${id}"><section class="tbxModal ${small ? 'sm' : ''}"><header class="tbxModalHead"><h3>${title}</h3><button class="tbxClose" data-close aria-label="Cerrar">×</button></header><div class="tbxModalBody">${body}</div></section></div>`;
  }

  private quickIcon(key: string): string {
    const map: Record<string,string> = { support:'headset', expenses:'receipt', docs:'file', calendar:'calendar', benefits:'heart', directory:'users', webops:'globe' };
    return map[key] || 'link';
  }

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

  private escapeHtml(value: string): string {
    return value.replace(/[&<>'"]/g, (char: string) => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', "'":'&#39;', '"':'&quot;' }[char] || char));
  }

  protected onDispose(): void {
    document.removeEventListener('keydown', this.onDocumentKeyDown);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}
