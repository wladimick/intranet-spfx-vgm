import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export type HubPersonItem = {
  Id: number;
  Title: string;
  Email?: string;
  Cargo?: string;
  Area?: string;
  FechaIngreso?: string;
  Cumpleanos?: string;
  Activo?: boolean;
};

export type HubMovementItem = {
  Id: number;
  Title: string;
  Email?: string;
  CargoAnterior?: string;
  CargoNuevo?: string;
  TipoMovimiento?: string;
  Area?: string;
  FechaMovimiento?: string;
  Activo?: boolean;
};

export type HubCourseItem = {
  Id: number;
  Title: string;
  Descripcion?: string;
  Categoria?: string;
  Obligatorio?: boolean;
  DuracionMinutos?: number;
  FechaLimite?: string;
  UrlCurso?: string;
  Estado?: string;
  ResponsableEmail?: string;
  CorreosNotificacion?: string;
  Activo?: boolean;
};

export type HubBenefitItem = {
  Id: number;
  Title: string;
  Descripcion?: string;
  Categoria?: string;
  Url?: string;
  ImagenUrl?: string;
  VigenciaHasta?: string;
  Activo?: boolean;
};

export type HubEventItem = {
  Id: number;
  Title: string;
  Tipo?: string;
  FechaInicio?: string;
  Modalidad?: string;
  DuracionMinutos?: number;
  Url?: string;
  ImagenUrl?: string;
  Activo?: boolean;
};

export type HubLearningItem = {
  Id: number;
  Title: string;
  Categoria?: string;
  TipoContenido?: string;
  Descripcion?: string;
  DuracionMinutos?: number;
  Url?: string;
  ImagenUrl?: string;
  Destacado?: boolean;
  Activo?: boolean;
};

export type HubNewsItem = {
  Id: number;
  Title: string;
  Resumen?: string;
  Categoria?: string;
  FechaPublicacion?: string;
  Url?: string;
  ImagenUrl?: string;
  Activo?: boolean;
};

export type HubDashboardData = {
  people: HubPersonItem[];
  movements: HubMovementItem[];
  courses: HubCourseItem[];
  benefits: HubBenefitItem[];
  events: HubEventItem[];
  learning: HubLearningItem[];
  news: HubNewsItem[];
};

type FieldDefinition = {
  name: string;
  xml: string;
};

type ListDefinition = {
  title: string;
  description: string;
  fields: FieldDefinition[];
  seed: Array<Record<string, unknown>>;
};

export class TiboxHubSharePointService {
  public static readonly lists = {
    people: 'TIBOX HUB - Colaboradores',
    movements: 'TIBOX HUB - Movimientos',
    courses: 'TIBOX HUB - Cursos',
    benefits: 'TIBOX HUB - Beneficios',
    events: 'TIBOX HUB - Eventos',
    learning: 'TIBOX HUB - Aprende',
    news: 'TIBOX HUB - Noticias'
  } as const;

  private readonly entityTypeCache: Map<string, string> = new Map<string, string>();

  public constructor(
    private readonly spHttpClient: SPHttpClient,
    private readonly webUrl: string
  ) {}

  public async ensureSchemaAndSeed(): Promise<void> {
    for (const definition of this.definitions()) {
      await this.ensureList(definition);
      await this.ensureFields(definition);
      await this.seedIfEmpty(definition);
    }
  }

  public async loadDashboardData(): Promise<HubDashboardData> {
    const [people, movements, courses, benefits, events, learning, news] = await Promise.all([
      this.getItems<HubPersonItem>(TiboxHubSharePointService.lists.people, '$select=Id,Title,Email,Cargo,Area,FechaIngreso,Cumpleanos,Activo&$filter=Activo eq 1&$orderby=FechaIngreso desc&$top=100'),
      this.getItems<HubMovementItem>(TiboxHubSharePointService.lists.movements, '$select=Id,Title,Email,CargoAnterior,CargoNuevo,TipoMovimiento,Area,FechaMovimiento,Activo&$filter=Activo eq 1&$orderby=FechaMovimiento desc&$top=50'),
      this.getItems<HubCourseItem>(TiboxHubSharePointService.lists.courses, '$select=Id,Title,Descripcion,Categoria,Obligatorio,DuracionMinutos,FechaLimite,UrlCurso,Estado,ResponsableEmail,CorreosNotificacion,Activo&$filter=Activo eq 1&$orderby=FechaLimite asc&$top=50'),
      this.getItems<HubBenefitItem>(TiboxHubSharePointService.lists.benefits, '$select=Id,Title,Descripcion,Categoria,Url,ImagenUrl,VigenciaHasta,Activo&$filter=Activo eq 1&$orderby=Title asc&$top=50'),
      this.getItems<HubEventItem>(TiboxHubSharePointService.lists.events, '$select=Id,Title,Tipo,FechaInicio,Modalidad,DuracionMinutos,Url,ImagenUrl,Activo&$filter=Activo eq 1&$orderby=FechaInicio asc&$top=50'),
      this.getItems<HubLearningItem>(TiboxHubSharePointService.lists.learning, '$select=Id,Title,Categoria,TipoContenido,Descripcion,DuracionMinutos,Url,ImagenUrl,Destacado,Activo&$filter=Activo eq 1&$orderby=Destacado desc,Title asc&$top=100'),
      this.getItems<HubNewsItem>(TiboxHubSharePointService.lists.news, '$select=Id,Title,Resumen,Categoria,FechaPublicacion,Url,ImagenUrl,Activo&$filter=Activo eq 1&$orderby=FechaPublicacion desc&$top=50')
    ]);

    return { people, movements, courses, benefits, events, learning, news };
  }

  public contentSiteUrl(): string {
    return `${this.webUrl}/_layouts/15/viewlsts.aspx`;
  }

  private async getItems<T>(listTitle: string, query: string): Promise<T[]> {
    try {
      const response: SPHttpClientResponse = await this.spHttpClient.get(
        `${this.webUrl}/_api/web/lists/getbytitle('${this.odata(listTitle)}')/items?${query}`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: 'application/json;odata=nometadata' } }
      );
      if (!response.ok) return [];
      const data: { value?: T[] } = await response.json() as { value?: T[] };
      return data.value || [];
    } catch {
      return [];
    }
  }

  private async ensureList(definition: ListDefinition): Promise<void> {
    const response: SPHttpClientResponse = await this.spHttpClient.get(
      `${this.webUrl}/_api/web/lists/getbytitle('${this.odata(definition.title)}')?$select=Id`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: 'application/json;odata=nometadata' } }
    );
    if (response.ok) return;

    const create: SPHttpClientResponse = await this.spHttpClient.post(
      `${this.webUrl}/_api/web/lists`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose'
        },
        body: JSON.stringify({
          __metadata: { type: 'SP.List' },
          BaseTemplate: 100,
          Title: definition.title,
          Description: definition.description,
          ContentTypesEnabled: false,
          EnableVersioning: true
        })
      }
    );

    if (!create.ok) {
      throw new Error(`No se pudo crear la lista ${definition.title}.`);
    }
  }

  private async ensureFields(definition: ListDefinition): Promise<void> {
    for (const field of definition.fields) {
      const existing: SPHttpClientResponse = await this.spHttpClient.get(
        `${this.webUrl}/_api/web/lists/getbytitle('${this.odata(definition.title)}')/fields/getbyinternalnameortitle('${this.odata(field.name)}')?$select=Id`,
        SPHttpClient.configurations.v1,
        { headers: { Accept: 'application/json;odata=nometadata' } }
      );
      if (existing.ok) continue;

      const create: SPHttpClientResponse = await this.spHttpClient.post(
        `${this.webUrl}/_api/web/lists/getbytitle('${this.odata(definition.title)}')/fields/createfieldasxml`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/json;odata=verbose',
            'Content-Type': 'application/json;odata=verbose'
          },
          body: JSON.stringify({
            parameters: {
              __metadata: { type: 'SP.XmlSchemaFieldCreationInformation' },
              SchemaXml: field.xml,
              Options: 0
            }
          })
        }
      );

      if (!create.ok) {
        throw new Error(`No se pudo crear el campo ${field.name} en ${definition.title}.`);
      }
    }
  }

  private async seedIfEmpty(definition: ListDefinition): Promise<void> {
    const response: SPHttpClientResponse = await this.spHttpClient.get(
      `${this.webUrl}/_api/web/lists/getbytitle('${this.odata(definition.title)}')/items?$select=Id&$top=1`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: 'application/json;odata=nometadata' } }
    );
    if (!response.ok) return;
    const data: { value?: Array<{ Id: number }> } = await response.json() as { value?: Array<{ Id: number }> };
    if (data.value?.length) return;

    for (const item of definition.seed) {
      await this.createItem(definition.title, item);
    }
  }

  private async createItem(listTitle: string, item: Record<string, unknown>): Promise<void> {
    const entityType: string = await this.getEntityType(listTitle);
    const response: SPHttpClientResponse = await this.spHttpClient.post(
      `${this.webUrl}/_api/web/lists/getbytitle('${this.odata(listTitle)}')/items`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=verbose',
          'Content-Type': 'application/json;odata=verbose'
        },
        body: JSON.stringify({ __metadata: { type: entityType }, ...item })
      }
    );
    if (!response.ok) {
      throw new Error(`No se pudo crear un registro inicial en ${listTitle}.`);
    }
  }

  private async getEntityType(listTitle: string): Promise<string> {
    const cached: string | undefined = this.entityTypeCache.get(listTitle);
    if (cached) return cached;

    const response: SPHttpClientResponse = await this.spHttpClient.get(
      `${this.webUrl}/_api/web/lists/getbytitle('${this.odata(listTitle)}')?$select=ListItemEntityTypeFullName`,
      SPHttpClient.configurations.v1,
      { headers: { Accept: 'application/json;odata=nometadata' } }
    );
    if (!response.ok) throw new Error(`No se pudo obtener el tipo de elemento de ${listTitle}.`);
    const data: { ListItemEntityTypeFullName: string } = await response.json() as { ListItemEntityTypeFullName: string };
    this.entityTypeCache.set(listTitle, data.ListItemEntityTypeFullName);
    return data.ListItemEntityTypeFullName;
  }

  private text(name: string, displayName: string): FieldDefinition {
    return { name, xml: `<Field Type="Text" Name="${name}" DisplayName="${displayName}" MaxLength="255" AddToDefaultView="TRUE" />` };
  }

  private note(name: string, displayName: string): FieldDefinition {
    return { name, xml: `<Field Type="Note" Name="${name}" DisplayName="${displayName}" NumLines="6" RichText="FALSE" AddToDefaultView="TRUE" />` };
  }

  private date(name: string, displayName: string, dateOnly: boolean = true): FieldDefinition {
    return { name, xml: `<Field Type="DateTime" Name="${name}" DisplayName="${displayName}" Format="${dateOnly ? 'DateOnly' : 'DateTime'}" AddToDefaultView="TRUE" />` };
  }

  private bool(name: string, displayName: string, defaultValue: boolean = true): FieldDefinition {
    return { name, xml: `<Field Type="Boolean" Name="${name}" DisplayName="${displayName}" AddToDefaultView="TRUE"><Default>${defaultValue ? '1' : '0'}</Default></Field>` };
  }

  private number(name: string, displayName: string): FieldDefinition {
    return { name, xml: `<Field Type="Number" Name="${name}" DisplayName="${displayName}" Decimals="0" Min="0" AddToDefaultView="TRUE" />` };
  }

  private definitions(): ListDefinition[] {
    return [
      {
        title: TiboxHubSharePointService.lists.people,
        description: 'Colaboradores visibles en TIBOX HUB. La foto se podrá resolver desde Microsoft 365 usando el correo.',
        fields: [
          this.text('Email', 'Correo'),
          this.text('Cargo', 'Cargo'),
          this.text('Area', 'Área'),
          this.date('FechaIngreso', 'Fecha de ingreso'),
          this.date('Cumpleanos', 'Cumpleaños'),
          this.bool('Activo', 'Activo')
        ],
        seed: [
          { Title: 'Camila Rojas', Email: 'camila.rojas@tibox.cl', Cargo: 'Product Designer', Area: 'Diseño', FechaIngreso: '2026-08-27T00:00:00Z', Cumpleanos: '2000-10-14T00:00:00Z', Activo: true },
          { Title: 'Matías Silva', Email: 'matias.silva@tibox.cl', Cargo: 'Cloud Engineer', Area: 'Infraestructura', FechaIngreso: '2026-08-27T00:00:00Z', Cumpleanos: '2000-11-22T00:00:00Z', Activo: true },
          { Title: 'Fernanda Castro', Email: 'fernanda.castro@tibox.cl', Cargo: 'Data Analyst', Area: 'Datos', FechaIngreso: '2026-08-24T00:00:00Z', Cumpleanos: '2000-12-03T00:00:00Z', Activo: true },
          { Title: 'Diego Fuentes', Email: 'diego.fuentes@tibox.cl', Cargo: 'QA Engineer', Area: 'Calidad', FechaIngreso: '2026-08-18T00:00:00Z', Cumpleanos: '2000-12-20T00:00:00Z', Activo: true },
          { Title: 'Paula Farías', Email: 'paula.farias@tibox.cl', Cargo: 'Marketing', Area: 'Marketing', FechaIngreso: '2024-03-12T00:00:00Z', Cumpleanos: '2000-09-01T00:00:00Z', Activo: true },
          { Title: 'Sebastián Rivas', Email: 'sebastian.rivas@tibox.cl', Cargo: 'Ingeniero', Area: 'Infraestructura', FechaIngreso: '2024-04-15T00:00:00Z', Cumpleanos: '2000-09-02T00:00:00Z', Activo: true },
          { Title: 'Antonia López', Email: 'antonia.lopez@tibox.cl', Cargo: 'Analista', Area: 'Finanzas', FechaIngreso: '2023-10-01T00:00:00Z', Cumpleanos: '2000-09-05T00:00:00Z', Activo: true },
          { Title: 'Rodrigo Paredes', Email: 'rodrigo.paredes@tibox.cl', Cargo: 'Product Manager', Area: 'Producto', FechaIngreso: '2023-06-10T00:00:00Z', Cumpleanos: '2000-09-09T00:00:00Z', Activo: true }
        ]
      },
      {
        title: TiboxHubSharePointService.lists.movements,
        description: 'Ascensos, cambios de rol y movimientos internos mostrados en TIBOX HUB.',
        fields: [
          this.text('Email', 'Correo'),
          this.text('CargoAnterior', 'Cargo anterior'),
          this.text('CargoNuevo', 'Cargo nuevo'),
          this.text('TipoMovimiento', 'Tipo de movimiento'),
          this.text('Area', 'Área'),
          this.date('FechaMovimiento', 'Fecha del movimiento'),
          this.bool('Activo', 'Activo')
        ],
        seed: [
          { Title: 'Braulio Contreras', Email: 'braulio.contreras@tibox.cl', CargoAnterior: 'Developer', CargoNuevo: 'Senior Developer', TipoMovimiento: 'Ascenso', Area: 'Desarrollo', FechaMovimiento: '2026-08-28T00:00:00Z', Activo: true },
          { Title: 'Javier Morales', Email: 'javier.morales@tibox.cl', CargoAnterior: 'Product Designer', CargoNuevo: 'AI & Design Specialist', TipoMovimiento: 'Cambio de rol', Area: 'Diseño', FechaMovimiento: '2026-08-20T00:00:00Z', Activo: true },
          { Title: 'Daniela Herrera', Email: 'daniela.herrera@tibox.cl', CargoAnterior: 'Analista de Proyectos', CargoNuevo: 'Coordinadora de Proyectos', TipoMovimiento: 'Ascenso', Area: 'Proyectos', FechaMovimiento: '2026-08-10T00:00:00Z', Activo: true }
        ]
      },
      {
        title: TiboxHubSharePointService.lists.courses,
        description: 'Cursos y capacitaciones mensuales de TIBOX HUB.',
        fields: [
          this.note('Descripcion', 'Descripción'),
          this.text('Categoria', 'Categoría'),
          this.bool('Obligatorio', 'Obligatorio', false),
          this.number('DuracionMinutos', 'Duración (minutos)'),
          this.date('FechaLimite', 'Fecha límite'),
          this.text('UrlCurso', 'URL del curso'),
          this.text('Estado', 'Estado'),
          this.text('ResponsableEmail', 'Correo responsable'),
          this.note('CorreosNotificacion', 'Correos a notificar'),
          this.bool('Activo', 'Activo')
        ],
        seed: [
          { Title: 'Prevención del acoso laboral', Descripcion: 'Contenido obligatorio para todos los colaboradores.', Categoria: 'Personas', Obligatorio: true, DuracionMinutos: 45, FechaLimite: '2026-09-08T00:00:00Z', UrlCurso: 'https://example.com/curso-acoso', Estado: 'Vence pronto', ResponsableEmail: 'capacitaciones@tibox.cl', CorreosNotificacion: 'todos@tibox.cl', Activo: true },
          { Title: 'Seguridad de la información', Descripcion: 'Buenas prácticas para proteger datos y credenciales.', Categoria: 'Seguridad', Obligatorio: true, DuracionMinutos: 80, FechaLimite: '2026-09-16T00:00:00Z', UrlCurso: 'https://example.com/curso-seguridad', Estado: 'Activo', ResponsableEmail: 'capacitaciones@tibox.cl', CorreosNotificacion: 'todos@tibox.cl', Activo: true },
          { Title: 'Introducción a la Ley de Datos', Descripcion: 'Principios esenciales sobre tratamiento de datos personales.', Categoria: 'Cultura', Obligatorio: false, DuracionMinutos: 35, FechaLimite: '2026-09-30T00:00:00Z', UrlCurso: 'https://example.com/curso-datos', Estado: 'Nuevo', ResponsableEmail: 'capacitaciones@tibox.cl', CorreosNotificacion: 'todos@tibox.cl', Activo: true }
        ]
      },
      {
        title: TiboxHubSharePointService.lists.benefits,
        description: 'Beneficios y convenios para colaboradores.',
        fields: [
          this.note('Descripcion', 'Descripción'),
          this.text('Categoria', 'Categoría'),
          this.text('Url', 'URL'),
          this.text('ImagenUrl', 'URL de imagen'),
          this.date('VigenciaHasta', 'Vigencia hasta'),
          this.bool('Activo', 'Activo')
        ],
        seed: [
          { Title: 'Giftcard de cumpleaños', Descripcion: 'Beneficio entregado durante el mes de cumpleaños.', Categoria: 'Celebraciones', Url: '#', ImagenUrl: '', VigenciaHasta: '2026-12-31T00:00:00Z', Activo: true },
          { Title: 'Convenio de salud', Descripcion: 'Condiciones preferenciales para colaboradores.', Categoria: 'Salud', Url: '#', ImagenUrl: '', VigenciaHasta: '2026-12-31T00:00:00Z', Activo: true },
          { Title: 'Capacitaciones', Descripcion: 'Acceso a instancias seleccionadas de formación.', Categoria: 'Educación', Url: '#', ImagenUrl: '', VigenciaHasta: '2026-12-31T00:00:00Z', Activo: true },
          { Title: 'Convenios comerciales', Descripcion: 'Descuentos y convenios disponibles para el equipo.', Categoria: 'Convenios', Url: '#', ImagenUrl: '', VigenciaHasta: '2026-12-31T00:00:00Z', Activo: true }
        ]
      },
      {
        title: TiboxHubSharePointService.lists.events,
        description: 'Eventos y actividades internas mostradas en la portada.',
        fields: [
          this.text('Tipo', 'Tipo'),
          this.date('FechaInicio', 'Fecha y hora', false),
          this.text('Modalidad', 'Modalidad'),
          this.number('DuracionMinutos', 'Duración (minutos)'),
          this.text('Url', 'URL'),
          this.text('ImagenUrl', 'URL de imagen'),
          this.bool('Activo', 'Activo')
        ],
        seed: [
          { Title: 'Workshop Inteligencia Artificial', Tipo: 'Capacitación', FechaInicio: '2026-09-04T15:00:00Z', Modalidad: 'Teams', DuracionMinutos: 60, Url: '#', ImagenUrl: '', Activo: true },
          { Title: 'Encuentro Tibox', Tipo: 'Empresa', FechaInicio: '2026-09-11T17:30:00Z', Modalidad: 'Presencial', DuracionMinutos: 90, Url: '#', ImagenUrl: '', Activo: true },
          { Title: 'Capacitación Excel Avanzado', Tipo: 'Capacitación', FechaInicio: '2026-09-18T10:00:00Z', Modalidad: 'Teams', DuracionMinutos: 45, Url: '#', ImagenUrl: '', Activo: true }
        ]
      },
      {
        title: TiboxHubSharePointService.lists.learning,
        description: 'Contenido de TIBOX Aprende: IA, Microsoft 365, seguridad y productividad.',
        fields: [
          this.text('Categoria', 'Categoría'),
          this.text('TipoContenido', 'Tipo de contenido'),
          this.note('Descripcion', 'Descripción'),
          this.number('DuracionMinutos', 'Duración (minutos)'),
          this.text('Url', 'URL'),
          this.text('ImagenUrl', 'URL de imagen'),
          this.bool('Destacado', 'Destacado', false),
          this.bool('Activo', 'Activo')
        ],
        seed: [
          { Title: 'Prompt de la semana', Categoria: 'Inteligencia Artificial', TipoContenido: 'Prompt', Descripcion: 'Analiza este correo, identifica las acciones que debo realizar, ordénalas por prioridad y redacta una respuesta breve y profesional.', DuracionMinutos: 5, Url: '#', ImagenUrl: '', Destacado: true, Activo: true },
          { Title: 'Cómo empezar con ChatGPT', Categoria: 'Inteligencia Artificial', TipoContenido: 'Guía', Descripcion: 'Conceptos básicos para usar IA de manera productiva.', DuracionMinutos: 10, Url: '#', ImagenUrl: '', Destacado: false, Activo: true },
          { Title: 'Copilot en Microsoft 365', Categoria: 'Microsoft 365', TipoContenido: 'Guía', Descripcion: 'Casos de uso para Outlook, Word y Teams.', DuracionMinutos: 12, Url: '#', ImagenUrl: '', Destacado: false, Activo: true },
          { Title: 'Buenas prácticas de contraseñas', Categoria: 'Seguridad', TipoContenido: 'Tip', Descripcion: 'Recomendaciones simples para proteger cuentas.', DuracionMinutos: 4, Url: '#', ImagenUrl: '', Destacado: false, Activo: true },
          { Title: 'Atajos para trabajar mejor', Categoria: 'Productividad', TipoContenido: 'Tip', Descripcion: 'Pequeñas mejoras para ahorrar tiempo todos los días.', DuracionMinutos: 5, Url: '#', ImagenUrl: '', Destacado: false, Activo: true }
        ]
      },
      {
        title: TiboxHubSharePointService.lists.news,
        description: 'Noticias y comunicados de TIBOX HUB.',
        fields: [
          this.note('Resumen', 'Resumen'),
          this.text('Categoria', 'Categoría'),
          this.date('FechaPublicacion', 'Fecha de publicación'),
          this.text('Url', 'URL'),
          this.text('ImagenUrl', 'URL de imagen'),
          this.bool('Activo', 'Activo')
        ],
        seed: [
          { Title: 'Tibox certifica a nuevo cliente del retail', Resumen: 'Nuevo cliente incorporado a nuestro ecosistema de servicios.', Categoria: 'Nuevo cliente', FechaPublicacion: '2026-08-30T00:00:00Z', Url: '#', ImagenUrl: '', Activo: true },
          { Title: 'Cerramos con éxito el proyecto Aurora', Resumen: 'El equipo completó una nueva implementación.', Categoria: 'Proyecto terminado', FechaPublicacion: '2026-08-28T00:00:00Z', Url: '#', ImagenUrl: '', Activo: true },
          { Title: 'Le damos la bienvenida al nuevo equipo de Datos', Resumen: 'Nuevas incorporaciones fortalecen las capacidades del equipo.', Categoria: 'Nuevo colaborador', FechaPublicacion: '2026-08-27T00:00:00Z', Url: '#', ImagenUrl: '', Activo: true },
          { Title: 'Actualización de la política de teletrabajo', Resumen: 'Revisa los principales cambios y consideraciones.', Categoria: 'Comunicado', FechaPublicacion: '2026-08-24T00:00:00Z', Url: '#', ImagenUrl: '', Activo: true }
        ]
      }
    ];
  }

  private odata(value: string): string {
    return value.replace(/'/g, "''");
  }
}
