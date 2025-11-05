
/**
 * Constantes de la aplicación: nombres de hojas y encabezados esperados.
 */
const SHEET_NAMES = {
  groups: 'Grupos de trabajo',
  events: 'Eventos',
  notices: 'Avisos'
};

const EXTERNAL_CALENDAR = {
  icsUrl:
    'https://calendar.google.com/calendar/ical/5e54a0c7b80d14449a969845a8547fd9d0f0724cfcfd3492a0311a8e319703d9%40group.calendar.google.com/public/basic.ics',
  calendarId: '5e54a0c7b80d14449a969845a8547fd9d0f0724cfcfd3492a0311a8e319703d9@group.calendar.google.com'
};

const ICS_CACHE_TTL_SECONDS = 300;

const SHEET_HEADERS = {
  [SHEET_NAMES.groups]: ['Grupo', 'Responsable', 'Contacto', 'Miembros', 'Próxima reunión', 'Notas'],
  [SHEET_NAMES.events]: ['Fecha inicio', 'Fecha fin', 'Hora inicio', 'Hora fin', 'Título', 'Descripción', 'Ubicación', 'Responsable'],
  [SHEET_NAMES.notices]: ['Tipo', 'Título', 'Descripción', 'Fecha', 'Contacto', 'Prioridad', 'Visibilidad', 'Adjuntos']
};

const NOTICE_VISIBILITY_WINDOW_DAYS = 60;

const NOTICE_PRIORITY_CONFIG = {
  high: { key: 'high', sheetValue: 'Alta', label: 'Alta prioridad', weight: 3 },
  medium: { key: 'medium', sheetValue: 'Media', label: 'Prioridad media', weight: 2 },
  normal: { key: 'normal', sheetValue: 'Normal', label: 'Prioridad normal', weight: 1 },
  low: { key: 'low', sheetValue: 'Baja', label: 'Prioridad baja', weight: 0 }
};

const NOTICE_DEFAULT_VISIBILITY = 'Sí';
const NOTICE_ATTACHMENT_SEPARATOR = '|';
const NOTICE_ALLOWED_PROTOCOLS = ['http:', 'https:'];
const NOTICE_IMAGE_EXTENSIONS = ['.jpg', '.jpeg', '.png', '.gif', '.webp', '.bmp', '.svg'];
const NOTICE_DEFAULT_ATTACHMENT_MIME_PREFIX = 'application/';

const CONFIGURED_ADMIN_EMAILS = (function () {
  try {
    const property = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAILS');
    if (!property) {
      return [];
    }
    return property
      .split(',')
      .map((email) => email.trim().toLowerCase())
      .filter(Boolean);
  } catch (error) {
    console.warn('No se pudieron leer los administradores configurados:', error);
    return [];
  }
})();

/**
 * Añade un menú personalizado para facilitar el acceso a la configuración.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Organizador comunidad')
    .addItem('Configurar estructura', 'crearEstructuraHoja')
    .addToUi();
}

/**
 * Crea (o actualiza) las hojas necesarias y coloca los encabezados requeridos.
 */
function crearEstructuraHoja() {
  const ss = SpreadsheetApp.getActive();

  Object.entries(SHEET_HEADERS).forEach(([sheetName, headers]) => {
    const sheet = getOrCreateSheet_(ss, sheetName);
    setHeaders_(sheet, headers);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
  });

  SpreadsheetApp.getUi().alert('La estructura se configuró correctamente.');
}

/**
 * Devuelve la información inicial para poblar la interfaz web.
 */
function getInitialPayload() {
  const today = new Date();
  const year = today.getFullYear();
  const month = today.getMonth() + 1;

  return {
    calendar: getCalendarData(year, month),
    groups: getGroupsData(),
    notices: getNoticesData(),
    permissions: {
      canManage: userCanManage_()
    }
  };
}

/**
 * Obtiene los eventos del mes indicado y metadatos del calendario.
 *
 * @param {number} year - Año con cuatro dígitos.
 * @param {number} month - Mes (1-12).
 * @return {Object} - Metadatos del calendario y eventos agrupados por día.
 */
function getCalendarData(year, month) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAMES.events);
  const metadata = buildCalendarMetadata_(year, month);

  if (!sheet) {
    return {
      metadata,
      events: fetchExternalCalendarEvents_(year, month)
    };
  }

  const rows = getRows_(sheet);

  if (!rows.length) {
    return {
      metadata,
      events: fetchExternalCalendarEvents_(year, month)
    };
  }

  const headers = rows.shift();

  // Intentar obtener columnas con la nueva estructura (Fecha inicio/fin)
  let col = getColumnIndexes_(headers, {
    dateStart: 'Fecha inicio',
    dateEnd: 'Fecha fin',
    start: 'Hora inicio',
    end: 'Hora fin',
    title: 'Título',
    description: 'Descripción',
    location: 'Ubicación',
    owner: 'Responsable'
  });

  // RETROCOMPATIBILIDAD: Si no encuentra "Fecha inicio", usar "Fecha" (estructura antigua)
  if (col.dateStart === null && headers.indexOf('Fecha') > -1) {
    console.log('⚠️ Usando estructura antigua con columna "Fecha". Se recomienda ejecutar "Configurar estructura" desde el menú.');
    col = getColumnIndexes_(headers, {
      dateStart: 'Fecha',
      dateEnd: 'Fecha', // En la estructura antigua, inicio y fin son la misma fecha
      start: 'Hora inicio',
      end: 'Hora fin',
      title: 'Título',
      description: 'Descripción',
      location: 'Ubicación',
      owner: 'Responsable'
    });
  }

  const sheetEvents = rows
    .map((row, i) => {
      const event = mapEventRow_(row, col);
      if (event) {
        event.rowIndex = i + 2; // +2 porque: índice 0-based + 1 fila de encabezado + 1 para 1-based
      }
      return event;
    })
    .filter((event) => {
      if (!event) return false;

      // Para eventos multi-día, verificar si el mes solicitado está en el rango
      const startYear = event.dateStart.getFullYear();
      const startMonth = event.dateStart.getMonth() + 1;
      const endYear = event.dateEnd.getFullYear();
      const endMonth = event.dateEnd.getMonth() + 1;

      // El evento aparece si:
      // - Empieza en este mes/año, O
      // - Termina en este mes/año, O
      // - Abarca este mes/año (empieza antes y termina después)
      const startsInMonth = startYear === year && startMonth === month;
      const endsInMonth = endYear === year && endMonth === month;
      const spansMonth = (startYear < year || (startYear === year && startMonth < month)) &&
                         (endYear > year || (endYear === year && endMonth > month));

      return startsInMonth || endsInMonth || spansMonth;
    })
    .flatMap((event) => {
      // Calcular todos los días en los que aparece el evento
      const days = [];
      const currentDate = new Date(event.dateStart);
      const endDate = new Date(event.dateEnd);

      // Límite de seguridad: máximo 90 días
      let dayCount = 0;
      const maxDays = 90;

      while (currentDate <= endDate && dayCount < maxDays) {
        const currentYear = currentDate.getFullYear();
        const currentMonth = currentDate.getMonth() + 1;

        // Solo incluir si está en el mes solicitado
        if (currentYear === year && currentMonth === month) {
          const dayNum = currentDate.getDate();
          const isFirstDay = currentDate.getTime() === event.dateStart.getTime();
          const isLastDay = currentDate.getTime() === endDate.getTime();
          const isMultiDay = event.dateStart.getTime() !== event.dateEnd.getTime();

          days.push({
            day: dayNum,
            title: event.title,
            description: event.description,
            location: event.location,
            owner: event.owner,
            startTime: event.startTime,
            endTime: event.endTime,
            rowIndex: event.rowIndex,
            // Metadata para eventos multi-día
            isMultiDay: isMultiDay,
            isFirstDay: isFirstDay,
            isLastDay: isLastDay,
            dateStart: Utilities.formatDate(event.dateStart, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
            dateEnd: Utilities.formatDate(event.dateEnd, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
            totalDays: Math.ceil((endDate - event.dateStart) / (1000 * 60 * 60 * 24)) + 1
          });
        }

        currentDate.setDate(currentDate.getDate() + 1);
        dayCount++;
      }

      return days;
    });

  const externalEvents = fetchExternalCalendarEvents_(year, month);
  const combinedEvents = sheetEvents
    .concat(externalEvents)
    .sort((a, b) => {
      if (a.day !== b.day) {
        return a.day - b.day;
      }
      return minutesFromTimeString_(a.startTime) - minutesFromTimeString_(b.startTime);
    });

  return { metadata, events: combinedEvents };
}

/**
 * Recupera la lista de grupos de trabajo.
 */
function getGroupsData() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAMES.groups);
  if (!sheet) {
    return [];
  }

  const rows = getRows_(sheet);
  if (!rows.length) {
    return [];
  }

  const headers = rows.shift();
  const col = getColumnIndexes_(headers, {
    group: 'Grupo',
    lead: 'Responsable',
    contact: 'Contacto',
    members: 'Miembros',
    meeting: 'Próxima reunión',
    notes: 'Notas'
  });

  return rows
    .map((row, i) => ({
      name: valueToString_(columnValue_(row, col.group)),
      lead: valueToString_(columnValue_(row, col.lead)),
      contact: valueToString_(columnValue_(row, col.contact)),
      members: splitMembers_(columnValue_(row, col.members)),
      nextMeeting: formatDateTime_(columnValue_(row, col.meeting)),
      notes: valueToString_(columnValue_(row, col.notes)),
      rowIndex: i + 2  // +2 porque: índice 0-based + 1 fila de encabezado + 1 para 1-based
    }))
    .filter((group) => group.name);
}

/**
 * Recupera los avisos y los ordena por tipo.
 */
function getNoticesData() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAMES.notices);
  if (!sheet) {
    return { parroquia: [], comunidad: [] };
  }

  const rows = getRows_(sheet);
  if (!rows.length) {
    return { parroquia: [], comunidad: [] };
  }

  let headers = rows.shift();
  headers = ensureNoticeHeaders_(sheet, headers);

  const col = getColumnIndexes_(headers, {
    type: 'Tipo',
    title: 'Título',
    description: 'Descripción',
    date: 'Fecha',
    contact: 'Contacto',
    priority: 'Prioridad',
    visibility: 'Visibilidad',
    attachments: 'Adjuntos'
  });

  const today = startOfDay_(new Date());
  const visibleUntil = addDays_(today, NOTICE_VISIBILITY_WINDOW_DAYS);

  const notices = { parroquia: [], comunidad: [] };

  rows.forEach((row, i) => {
    const typeRaw = getRowValue_(row, col.type);
    const type = valueToString_(typeRaw).toLowerCase();

    const title = valueToString_(getRowValue_(row, col.title));
    const description = valueToString_(getRowValue_(row, col.description));

    if (!title && !description) {
      return;
    }

    const rawDate = getRowValue_(row, col.date);
    const normalizedDate = normalizeNoticeDate_(rawDate);

    if (normalizedDate) {
      if (normalizedDate.getTime() < today.getTime()) {
        return;
      }
      if (normalizedDate.getTime() > visibleUntil.getTime()) {
        return;
      }
    }

    const visibilityValue = getRowValue_(row, col.visibility);
    if (!isNoticeVisible_(visibilityValue)) {
      return;
    }

    const priorityInfo = getNoticePriorityMetadata_(getRowValue_(row, col.priority));
    const contact = valueToString_(getRowValue_(row, col.contact));
    const formattedDate = normalizedDate ? formatDate_(normalizedDate) : '';
    const isoDate = normalizedDate
      ? Utilities.formatDate(normalizedDate, Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : '';
    const attachments = parseStoredAttachments_(getRowValue_(row, col.attachments));

    const notice = {
      title,
      description,
      date: formattedDate,
      dateIso: isoDate,
      contact,
      priority: priorityInfo.sheetValue,
      priorityLabel: priorityInfo.label,
      priorityKey: priorityInfo.key,
      priorityLevel: priorityInfo.weight,
      attachments,
      sortDateValue: normalizedDate ? normalizedDate.getTime() : Number.MAX_SAFE_INTEGER,
      rowIndex: i + 2  // +2 porque: índice 0-based + 1 fila de encabezado + 1 para 1-based
    };

    if (type.includes('parro')) {
      notices.parroquia.push(notice);
      return;
    }

    notices.comunidad.push(notice);
  });

  const comparator = (a, b) => {
    if (b.priorityLevel !== a.priorityLevel) {
      return b.priorityLevel - a.priorityLevel;
    }
    if (a.sortDateValue !== b.sortDateValue) {
      return a.sortDateValue - b.sortDateValue;
    }
    const titleA = (a.title || '').toLocaleUpperCase('es');
    const titleB = (b.title || '').toLocaleUpperCase('es');
    if (titleA < titleB) {
      return -1;
    }
    if (titleA > titleB) {
      return 1;
    }
    return 0;
  };

  return {
    parroquia: finalizeNoticeList_(notices.parroquia, comparator),
    comunidad: finalizeNoticeList_(notices.comunidad, comparator)
  };
}

/**
 * Agrega un nuevo evento al calendario.
 *
 * @param {Object} payload - Datos del evento enviados desde la UI.
 */
function addEvent(payload) {
  ensureCanManage_();

  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAMES.events);
  if (!sheet) {
    throw new Error('La hoja de "Eventos" no existe. Ejecuta "Configurar estructura" desde el menú.');
  }

  const eventData = sanitizeEventPayload_(payload);
  sheet.appendRow([
    eventData.dateStart,
    eventData.dateEnd,
    eventData.startTime,
    eventData.endTime,
    eventData.title,
    eventData.description,
    eventData.location,
    eventData.owner
  ]);
}

/**
 * Agrega un nuevo grupo de trabajo.
 *
 * @param {Object} payload - Datos del grupo.
 */
function addGroup(payload) {
  ensureCanManage_();

  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAMES.groups);
  if (!sheet) {
    throw new Error('La hoja de "Grupos de trabajo" no existe. Ejecuta "Configurar estructura" desde el menú.');
  }

  const group = sanitizeGroupPayload_(payload);
  sheet.appendRow([
    group.name,
    group.lead,
    group.contact,
    group.members,
    group.nextMeeting,
    group.notes
  ]);
}

/**
 * Agrega un nuevo aviso.
 *
 * @param {Object} payload - Datos del aviso.
 */
function addNotice(payload) {
  ensureCanManage_();

  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAMES.notices);
  if (!sheet) {
    throw new Error('La hoja de "Avisos" no existe. Ejecuta "Configurar estructura" desde el menú.');
  }

  const notice = sanitizeNoticePayload_(payload);
  sheet.appendRow([
    notice.type,
    notice.title,
    notice.description,
    notice.date,
    notice.contact,
    notice.priority,
    notice.visibility,
    notice.attachments
  ]);
}

/**
 * Renderiza la interfaz como una Web App.
 */
function doGet() {
  const template = HtmlService.createTemplateFromFile('Index');
  template.initialState = serializeInitialState_(getInitialPayload());
  return template
    .evaluate()
    .setTitle('Organizador de la comunidad')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Utilidad para incluir archivos HTML (si se necesitaran parciales).
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Serializa el estado inicial asegurando que no rompa la etiqueta <script>.
 * @private
 */
function serializeInitialState_(payload) {
  const json = JSON.stringify(payload || {});
  return json
    .replace(/</g, '\\u003c')
    .replace(/>/g, '\\u003e')
    .replace(/\u2028/g, '\\u2028')
    .replace(/\u2029/g, '\\u2029');
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/**
 * Determina si el usuario actual tiene permisos de administración.
 * @private
 */
function userCanManage_() {
  const email = (Session.getActiveUser().getEmail() || '').toLowerCase();
  if (!email) {
    return false;
  }

  const allowed = new Set(CONFIGURED_ADMIN_EMAILS);
  try {
    const owner = SpreadsheetApp.getActive().getOwner();
    if (owner) {
      allowed.add((owner.getEmail() || '').toLowerCase());
    }
  } catch (error) {
    console.warn('No se pudo obtener el propietario de la hoja:', error);
  }

  if (!allowed.size) {
    return false;
  }

  return allowed.has(email);
}

/**
 * Lanza un error si el usuario no tiene permisos de administración.
 * @private
 */
function ensureCanManage_() {
  if (!userCanManage_()) {
    throw new Error('No tienes permisos para realizar esta acción.');
  }
}

/**
 * Limpia y valida los datos de un evento.
 * @private
 */
function sanitizeEventPayload_(payload) {
  if (!payload) {
    throw new Error('No se recibieron datos del evento.');
  }

  const title = sanitizeTextInput_(payload.title);
  if (!title) {
    throw new Error('El título del evento es obligatorio.');
  }

  const dateStart = parseDateInput_(payload.dateStart);

  // Fecha fin es opcional
  let dateEnd = null;
  if (payload.dateEnd && sanitizeTextInput_(payload.dateEnd)) {
    dateEnd = parseDateInput_(payload.dateEnd);

    // Validar que dateEnd >= dateStart
    if (dateEnd < dateStart) {
      throw new Error('La fecha de fin no puede ser anterior a la fecha de inicio.');
    }

    // Validar duración máxima (90 días)
    const diffMs = dateEnd.getTime() - dateStart.getTime();
    const diffDays = Math.ceil(diffMs / (1000 * 60 * 60 * 24));
    if (diffDays > 90) {
      throw new Error('La duración máxima de un evento es de 90 días.');
    }
  } else {
    // Si no hay fecha fin, usar fecha inicio
    dateEnd = new Date(dateStart);
  }

  const startTime = sanitizeTimeInput_(sanitizeTextInput_(payload.startTime));
  const endTime = sanitizeTimeInput_(sanitizeTextInput_(payload.endTime));

  return {
    dateStart,
    dateEnd,
    startTime,
    endTime,
    title,
    description: sanitizeTextInput_(payload.description),
    location: sanitizeTextInput_(payload.location),
    owner: sanitizeTextInput_(payload.owner)
  };
}

/**
 * Limpia y valida los datos de un grupo de trabajo.
 * @private
 */
function sanitizeGroupPayload_(payload) {
  if (!payload) {
    throw new Error('No se recibieron datos del grupo.');
  }

  const name = sanitizeTextInput_(payload.name);
  if (!name) {
    throw new Error('El nombre del grupo es obligatorio.');
  }

  const membersRaw = sanitizeTextInput_(payload.members);
  const members = membersRaw
    ? membersRaw
        .split(',')
        .map((member) => member.trim())
        .filter(Boolean)
        .join(', ')
    : '';

  const nextMeetingValue = sanitizeTextInput_(payload.nextMeeting);
  const nextMeeting = parseDateTimeInput_(nextMeetingValue);

  return {
    name,
    lead: sanitizeTextInput_(payload.lead),
    contact: sanitizeTextInput_(payload.contact),
    members,
    nextMeeting,
    notes: sanitizeTextInput_(payload.notes)
  };
}

/**
 * Limpia y valida los datos de un aviso.
 * @private
 */
function sanitizeNoticePayload_(payload) {
  if (!payload) {
    throw new Error('No se recibieron datos del aviso.');
  }

  const title = sanitizeTextInput_(payload.title);
  const description = sanitizeTextInput_(payload.description);

  if (!title && !description) {
    throw new Error('Debes indicar al menos un título o una descripción para el aviso.');
  }

  const type = sanitizeTextInput_(payload.type).toLowerCase();
  const normalizedType = type.startsWith('parro') ? 'Parroquia' : 'Comunidad';

  const dateValue = sanitizeTextInput_(payload.date);
  const date = dateValue ? parseDateInput_(dateValue) : '';
  const priorityInfo = getNoticePriorityMetadata_(payload.priority);
  const visibility = normalizeVisibilityValue_(payload.visibility);
  const attachments = serializeNoticeAttachments_(payload.attachments);

  return {
    type: normalizedType,
    title,
    description,
    date,
    contact: sanitizeTextInput_(payload.contact),
    priority: priorityInfo.sheetValue,
    visibility,
    attachments
  };
}

/**
 * Normaliza un texto eliminando espacios extra.
 * @private
 */
function sanitizeTextInput_(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value).trim();
}

/**
 * Convierte una cadena AAAA-MM-DD en Date.
 * @private
 */
function parseDateInput_(value) {
  const text = sanitizeTextInput_(value);
  if (!text) {
    throw new Error('La fecha es obligatoria.');
  }

  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(text);
  if (!match) {
    throw new Error('Formato de fecha inválido. Usa AAAA-MM-DD.');
  }

  const year = Number(match[1]);
  const month = Number(match[2]) - 1;
  const day = Number(match[3]);

  return new Date(year, month, day);
}

/**
 * Convierte una cadena AAAA-MM-DD o AAAA-MM-DDTHH:mm en Date.
 * @private
 */
function parseDateTimeInput_(value) {
  const text = sanitizeTextInput_(value);
  if (!text) {
    return '';
  }

  const match = /^(\d{4})-(\d{2})-(\d{2})(?:[T\s](\d{2}):(\d{2}))?$/.exec(text);
  if (!match) {
    return text;
  }

  const year = Number(match[1]);
  const month = Number(match[2]) - 1;
  const day = Number(match[3]);
  const hours = match[4] ? Number(match[4]) : 0;
  const minutes = match[5] ? Number(match[5]) : 0;

  return new Date(year, month, day, hours, minutes);
}

/**
 * Valida y normaliza una hora HH:mm.
 * @private
 */
function sanitizeTimeInput_(value) {
  if (!value) {
    return '';
  }

  const match = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(value);
  if (!match) {
    throw new Error('Formato de hora inválido. Usa HH:MM.');
  }

  const hours = Number(match[1]);
  const minutes = Number(match[2]);

  if (hours > 23 || minutes > 59) {
    throw new Error('La hora indicada no es válida.');
  }

  return padNumber_(hours) + ':' + padNumber_(minutes);
}

/**
 * Añade un cero a la izquierda.
 * @private
 */
function padNumber_(value) {
  return value < 10 ? '0' + value : String(value);
}

/**
 * Obtiene o crea una hoja con el nombre indicado.
 * @private
 */
function getOrCreateSheet_(spreadsheet, sheetName) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }
  return sheet;
}

/**
 * Escribe los encabezados en la fila 1.
 * @private
 */
function setHeaders_(sheet, headers) {
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f1f3f4');
}

/**
 * Obtiene todas las filas de la hoja como matriz, si existe contenido.
 * @private
 */
function getRows_(sheet) {
  const range = sheet.getDataRange();
  const values = range.getValues();
  return values.filter((row) => row.some((cell) => cell !== ''));
}

/**
 * Construye los índices de columnas a partir de los encabezados.
 * @private
 */
function getColumnIndexes_(headers, mapping) {
  const indexMap = {};
  Object.entries(mapping).forEach(([key, header]) => {
    const idx = headers.indexOf(header);
    indexMap[key] = idx > -1 ? idx : null;
  });
  return indexMap;
}

/**
 * Garantiza que la hoja de avisos disponga de las columnas adicionales requeridas.
 * @private
 */
function ensureNoticeHeaders_(sheet, headers) {
  let currentHeaders = Array.isArray(headers) ? headers.slice() : [];
  currentHeaders = ensureHeaderExists_(sheet, currentHeaders, 'Prioridad');
  currentHeaders = ensureHeaderExists_(sheet, currentHeaders, 'Visibilidad');
  currentHeaders = ensureHeaderExists_(sheet, currentHeaders, 'Adjuntos');
  return currentHeaders;
}

/**
 * Asegura la existencia de un encabezado en la hoja, creándolo si fuera necesario.
 * @private
 */
function ensureHeaderExists_(sheet, headers, headerLabel) {
  const label = sanitizeTextInput_(headerLabel);
  if (!label) {
    return headers;
  }

  const normalizedIndex = headers.findIndex(
    (header) => sanitizeTextInput_(header).toLowerCase() === label.toLowerCase()
  );

  if (normalizedIndex > -1) {
    if (headers[normalizedIndex] !== label) {
      headers[normalizedIndex] = label;
      sheet.getRange(1, normalizedIndex + 1).setValue(label).setFontWeight('bold').setBackground('#f1f3f4');
    }
    return headers;
  }

  const blankIndex = headers.findIndex((header) => !sanitizeTextInput_(header));
  if (blankIndex > -1) {
    headers[blankIndex] = label;
    sheet.getRange(1, blankIndex + 1).setValue(label).setFontWeight('bold').setBackground('#f1f3f4');
    return headers;
  }

  const columnPosition = headers.length + 1;
  sheet.getRange(1, columnPosition).setValue(label).setFontWeight('bold').setBackground('#f1f3f4');
  headers.push(label);
  return headers;
}

/**
 * Obtiene un valor seguro para un índice de fila.
 * @private
 */
function getRowValue_(row, index) {
  if (!row || index === null || index === undefined) {
    return '';
  }
  return row[index];
}

/**
 * Normaliza una fecha a las 00:00.
 * @private
 */
function startOfDay_(date) {
  if (!(date instanceof Date)) {
    return null;
  }
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

/**
 * Suma días conservando sólo la porción de fecha.
 * @private
 */
function addDays_(date, days) {
  const base = startOfDay_(date);
  if (!base) {
    return null;
  }
  const result = new Date(base);
  result.setDate(result.getDate() + Number(days || 0));
  return startOfDay_(result);
}

/**
 * Devuelve una fecha Date a partir del valor almacenado en la hoja.
 * @private
 */
function normalizeNoticeDate_(value) {
  if (value instanceof Date) {
    return startOfDay_(value);
  }

  const text = sanitizeTextInput_(value);
  if (!text) {
    return null;
  }

  const isoMatch = /^(\d{4})-(\d{2})-(\d{2})/.exec(text);
  if (isoMatch) {
    return new Date(Number(isoMatch[1]), Number(isoMatch[2]) - 1, Number(isoMatch[3]));
  }

  const datePart = text.split(/[ T]/)[0];
  const europeanMatch = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(datePart);
  if (europeanMatch) {
    return new Date(Number(europeanMatch[3]), Number(europeanMatch[2]) - 1, Number(europeanMatch[1]));
  }

  return null;
}

/**
 * Determina si un aviso debe mostrarse según su valor de visibilidad.
 * @private
 */
function isNoticeVisible_(value) {
  const text = sanitizeTextInput_(value).toLowerCase();
  if (!text) {
    return true;
  }
  return !['no', 'n', 'false', '0'].includes(text);
}

/**
 * Normaliza el valor almacenado de visibilidad.
 * @private
 */
function normalizeVisibilityValue_(value) {
  return isNoticeVisible_(value) ? NOTICE_DEFAULT_VISIBILITY : 'No';
}

/**
 * Serializa los adjuntos recibidos desde la interfaz.
 * @private
 */
function serializeNoticeAttachments_(input) {
  const attachments = parseNoticeAttachmentsInput_(input);
  if (!attachments.length) {
    return '';
  }
  return JSON.stringify(attachments);
}

/**
 * Parsea la entrada del formulario para construir adjuntos.
 * Cada línea puede contener "Etiqueta|valor" o sólo un valor.
 * El valor puede ser una URL o un ID de archivo de Drive.
 * @private
 */
function parseNoticeAttachmentsInput_(value) {
  if (!value) {
    return [];
  }

  if (Array.isArray(value)) {
    return value
      .map((entry) => normalizeAttachmentRecord_(entry))
      .filter(Boolean);
  }

  const text = sanitizeTextInput_(value);
  if (!text) {
    return [];
  }

  return text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => {
      const separatorIndex = line.indexOf(NOTICE_ATTACHMENT_SEPARATOR);
      let labelPart = '';
      let valuePart = line;

      if (separatorIndex > -1) {
        labelPart = sanitizeTextInput_(line.substring(0, separatorIndex));
        valuePart = line.substring(separatorIndex + 1);
      }

      return buildAttachmentFromInput_(labelPart, valuePart);
    })
    .filter(Boolean);
}

/**
 * Convierte la representación almacenada de adjuntos a objetos utilizables.
 * @private
 */
function parseStoredAttachments_(value) {
  if (!value) {
    return [];
  }

  if (Array.isArray(value)) {
    return value.map((entry) => normalizeAttachmentRecord_(entry)).filter(Boolean);
  }

  if (typeof value === 'string') {
    const text = value.trim();
    if (!text) {
      return [];
    }

    if (text.startsWith('[')) {
      try {
        const json = JSON.parse(text);
        if (Array.isArray(json)) {
          return json.map((entry) => normalizeAttachmentRecord_(entry)).filter(Boolean);
        }
      } catch (error) {
        console.warn('No se pudo interpretar el JSON de adjuntos:', error);
      }
    }

    return parseNoticeAttachmentsInput_(text);
  }

  return [];
}

/**
 * Normaliza un objeto adjunto.
 * @private
 */
function normalizeAttachmentRecord_(record) {
  if (!record) {
    return null;
  }

  const preferredLabel = sanitizeTextInput_(record.label || record.name || record.title);
  const driveId = sanitizeTextInput_(record.driveId || record.id || record.fileId);

  let resolvedSource = null;

  const urlCandidate = record.url || record.href || record.link;
  if (urlCandidate) {
    resolvedSource = resolveAttachmentSource_(urlCandidate);
  }

  if (!resolvedSource && driveId) {
    resolvedSource = tryResolveDriveAttachment_(driveId);
  }

  if (!resolvedSource) {
    return null;
  }

  const typeInput = sanitizeTextInput_(record.type || record.kind).toLowerCase();
  const explicitType = ['image', 'pdf', 'file'].includes(typeInput) ? typeInput : '';

  return {
    label: preferredLabel || resolvedSource.label || deriveAttachmentLabel_(resolvedSource.url),
    url: resolvedSource.url,
    type: explicitType || resolvedSource.type || detectAttachmentType_(resolvedSource.url),
    driveId: resolvedSource.driveId || driveId || ''
  };
}

/**
 * Determina el tipo de adjunto a partir de la URL.
 * @private
 */
function detectAttachmentType_(url) {
  const normalized = sanitizeTextInput_(url).toLowerCase();
  if (!normalized) {
    return 'file';
  }
  if (normalized.endsWith('.pdf')) {
    return 'pdf';
  }
  if (NOTICE_IMAGE_EXTENSIONS.some((ext) => normalized.endsWith(ext))) {
    return 'image';
  }
  return 'file';
}

/**
 * Determina el tipo de adjunto a partir del mimetype.
 * @private
 */
function detectAttachmentTypeFromMime_(mimeType) {
  const text = sanitizeTextInput_(mimeType).toLowerCase();
  if (!text) {
    return 'file';
  }
  if (text.startsWith('image/')) {
    return 'image';
  }
  if (text === 'application/pdf' || text.endsWith('.pdf')) {
    return 'pdf';
  }
  if (text.startsWith(NOTICE_DEFAULT_ATTACHMENT_MIME_PREFIX)) {
    return 'file';
  }
  return 'file';
}

/**
 * Intenta normalizar una URL de adjunto y validar su protocolo.
 * @private
 */
function normalizeAttachmentUrl_(value) {
  const text = sanitizeTextInput_(value);
  if (!text) {
    return '';
  }

  const trimmed = text.trim();
  if (!/^https?:\/\//i.test(trimmed)) {
    return '';
  }

  try {
    const url = new URL(trimmed);
    if (!NOTICE_ALLOWED_PROTOCOLS.includes(url.protocol.toLowerCase())) {
      return '';
    }
    return url.toString();
  } catch (error) {
    return '';
  }
}

/**
 * Construye un adjunto a partir de la entrada del formulario.
 * @private
 */
function buildAttachmentFromInput_(labelPart, rawValue) {
  const source = resolveAttachmentSource_(rawValue);
  if (!source) {
    return null;
  }

  const label =
    sanitizeTextInput_(labelPart) || source.label || deriveAttachmentLabel_(source.url);

  return {
    label,
    url: source.url,
    type: source.type,
    driveId: source.driveId || ''
  };
}

/**
 * Determina la procedencia de un adjunto a partir del valor ingresado.
 * Puede ser una URL o un ID de archivo de Drive.
 * @private
 */
function resolveAttachmentSource_(rawValue) {
  const text = sanitizeTextInput_(rawValue);
  if (!text) {
    return null;
  }

  if (/^https?:\/\//i.test(text)) {
    const normalizedUrl = normalizeAttachmentUrl_(text);
    if (!normalizedUrl) {
      return null;
    }
    return {
      url: normalizedUrl,
      type: detectAttachmentType_(normalizedUrl)
    };
  }

  return tryResolveDriveAttachment_(text);
}

/**
 * Intenta resolver un adjunto de Drive a partir de un ID.
 * @private
 */
function tryResolveDriveAttachment_(fileId) {
  const normalizedId = sanitizeTextInput_(fileId);
  if (!normalizedId) {
    return null;
  }
  try {
    const file = DriveApp.getFileById(normalizedId);
    const url = file.getUrl();
    const mimeType = file.getMimeType();
    const label = file.getName();
    return {
      url,
      label,
      type: detectAttachmentTypeFromMime_(mimeType),
      driveId: normalizedId
    };
  } catch (error) {
    console.warn('No se pudo resolver el archivo adjunto de Drive:', normalizedId, error);
    return null;
  }
}

/**
 * Genera una etiqueta amigable a partir de una URL.
 * @private
 */
function deriveAttachmentLabel_(url) {
  try {
    const parsed = new URL(url);
    const pathname = parsed.pathname || '';
    const segments = pathname.split('/').filter(Boolean);
    if (segments.length) {
      const lastSegment = decodeURIComponent(segments[segments.length - 1]);
      if (lastSegment) {
        return lastSegment;
      }
    }
  } catch (error) {
    // Ignorar y continuar con el valor por defecto.
  }
  return detectAttachmentType_(url) === 'image' ? 'Imagen' : 'Archivo';
}

/**
 * Devuelve la configuración de prioridad para un aviso.
 * @private
 */
function getNoticePriorityMetadata_(value) {
  const text = sanitizeTextInput_(value).toLowerCase();

  if (text.startsWith('alta') || text === 'urgente' || text === 'high' || text === 'important') {
    return NOTICE_PRIORITY_CONFIG.high;
  }

  if (text.startsWith('media') || text === 'medio' || text === 'medium') {
    return NOTICE_PRIORITY_CONFIG.medium;
  }

  if (text.startsWith('baja') || text === 'low' || text === 'minima') {
    return NOTICE_PRIORITY_CONFIG.low;
  }

  if (text.startsWith('normal') || text === 'estandar' || text === 'standard') {
    return NOTICE_PRIORITY_CONFIG.normal;
  }

  return NOTICE_PRIORITY_CONFIG.normal;
}

/**
 * Ordena y elimina metadatos internos de la lista de avisos.
 * @private
 */
function finalizeNoticeList_(list, comparator) {
  if (!Array.isArray(list) || !list.length) {
    return [];
  }
  return list
    .slice()
    .sort(comparator)
    .map((notice) => {
      const { sortDateValue, ...rest } = notice;
      if (!Array.isArray(rest.attachments)) {
        rest.attachments = [];
      }
      return rest;
    });
}

/**
 * Convierte a string seguro.
 * @private
 */
function valueToString_(value) {
  if (value === null || value === undefined) {
    return '';
  }
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  }
  return String(value).trim();
}

/**
 * Formatea una fecha sencilla (sin hora).
 * @private
 */
function formatDate_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }
  return valueToString_(value);
}

/**
 * Formatea fecha y hora de una celda, si aplica.
 * @private
 */
function formatDateTime_(value) {
  if (value instanceof Date) {
    const hasTime = value.getHours() + value.getMinutes() + value.getSeconds() > 0;
    const pattern = hasTime ? 'dd/MM/yyyy HH:mm' : 'dd/MM/yyyy';
    return Utilities.formatDate(value, Session.getScriptTimeZone(), pattern);
  }
  return valueToString_(value);
}

/**
 * Crea los metadatos básicos para renderizar un calendario.
 * @private
 */
function buildCalendarMetadata_(year, month) {
  const firstDay = new Date(year, month - 1, 1);
  const lastDay = new Date(year, month, 0);
  return {
    year,
    month,
    firstWeekday: firstDay.getDay(), // 0 = domingo
    totalDays: lastDay.getDate()
  };
}

/**
 * Mapea una fila de evento a un objeto evento.
 * @private
 */
function mapEventRow_(row, col) {
  const dateStart = row[col.dateStart];
  if (!(dateStart instanceof Date)) {
    return null;
  }

  // Fecha fin es opcional, si no existe = fecha inicio
  let dateEnd = row[col.dateEnd];
  if (!(dateEnd instanceof Date) || dateEnd < dateStart) {
    dateEnd = new Date(dateStart);
  }

  return {
    dateStart,
    dateEnd,
    // Mantener 'date' para compatibilidad (deprecated)
    date: dateStart,
    title: valueToString_(columnValue_(row, col.title)),
    description: valueToString_(columnValue_(row, col.description)),
    location: valueToString_(columnValue_(row, col.location)),
    owner: valueToString_(columnValue_(row, col.owner)),
    startTime: formatHour_(columnValue_(row, col.start)),
    endTime: formatHour_(columnValue_(row, col.end))
  };
}

/**
 * Obtiene el valor de la columna indicada, considerando índices nulos.
 * @private
 */
function columnValue_(row, index) {
  return index === null ? '' : row[index];
}

/**
 * Formatea horas en HH:mm cuando hay información disponible.
 * @private
 */
function formatHour_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'HH:mm');
  }
  const asString = valueToString_(value);
  return asString || '';
}

/**
 * Separa la lista de miembros delimitada por comas en un arreglo limpio.
 * @private
 */
function splitMembers_(value) {
  const text = valueToString_(value);
  if (!text) {
    return [];
  }

  return text
    .split(',')
    .map((member) => member.trim())
    .filter(Boolean);
}

/**
 * Recupera eventos desde el calendario externo público (ICS) para el mes solicitado.
 * @private
 */
function fetchExternalCalendarEvents_(year, month) {
  if (!EXTERNAL_CALENDAR.icsUrl) {
    return [];
  }

  const cacheKey = ['ics', year, month].join('_');
  const cache = CacheService.getScriptCache();

  try {
    const cached = cache.get(cacheKey);
    if (cached) {
      const parsed = JSON.parse(cached);
      if (Array.isArray(parsed)) {
        return parsed;
      }
    }
  } catch (error) {
    console.warn('No se pudo leer la caché de eventos externos:', error);
  }

  try {
    const response = UrlFetchApp.fetch(EXTERNAL_CALENDAR.icsUrl, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) {
      console.warn('No se pudo obtener el calendario externo. Código:', response.getResponseCode());
      return [];
    }

    const ics = response.getContentText();
    const events = parseExternalCalendarEvents_(ics, year, month);

    try {
      cache.put(cacheKey, JSON.stringify(events), ICS_CACHE_TTL_SECONDS);
    } catch (cacheError) {
      console.warn('No se pudo almacenar la caché de eventos externos:', cacheError);
    }

    return events;
  } catch (error) {
    console.warn('No se pudieron cargar los eventos del calendario externo:', error);
    return [];
  }
}

/**
 * Convierte un archivo ICS en eventos de un mes concreto.
 * @private
 */
function parseExternalCalendarEvents_(ics, year, month) {
  if (!ics) {
    return [];
  }

  const unfolded = unfoldIcsContent_(ics);
  const chunks = unfolded.split('BEGIN:VEVENT').slice(1);
  const events = [];
  const timeZone = Session.getScriptTimeZone();
  const monthStart = new Date(year, month - 1, 1);
  const monthEnd = new Date(year, month, 0);
  monthEnd.setHours(23, 59, 59, 999);
  const monthStartTime = monthStart.getTime();
  const monthEndTime = monthEnd.getTime();

  chunks.forEach((chunk) => {
    const body = chunk.split('END:VEVENT')[0];
    const lines = body
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter(Boolean);

    if (!lines.length) {
      return;
    }

    const dtStartLine = extractIcsLine_(lines, 'DTSTART');
    const dtEndLine = extractIcsLine_(lines, 'DTEND');
    const startDate = parseIcsDateValue_(dtStartLine);

    if (!startDate) {
      return;
    }

    const rawEndDate = parseIcsDateValue_(dtEndLine);
    const endDate = rawEndDate || startDate;
    const eventStartTime = startDate.getTime();
    const eventEndExclusiveTime = computeIcsExclusiveEnd_(startDate, rawEndDate).getTime();

    if (eventEndExclusiveTime <= monthStartTime || eventStartTime > monthEndTime) {
      return;
    }

    const effectiveTimestamp = Math.max(eventStartTime, monthStartTime);
    const effectiveDate = new Date(effectiveTimestamp);

    if (effectiveDate.getFullYear() !== year || effectiveDate.getMonth() + 1 !== month) {
      return;
    }

    const summary = decodeIcsText_(getIcsValue_(extractIcsLine_(lines, 'SUMMARY')));
    const description = decodeIcsText_(getIcsValue_(extractIcsLine_(lines, 'DESCRIPTION')));
    const location = decodeIcsText_(getIcsValue_(extractIcsLine_(lines, 'LOCATION')));
    const organizer = parseIcsOrganizer_(extractIcsLine_(lines, 'ORGANIZER'));

    const startTime = hasTimeComponent_(dtStartLine) ? formatTimeString_(startDate, timeZone) : '';
    const endTime = hasTimeComponent_(dtEndLine) ? formatTimeString_(endDate, timeZone) : '';

    events.push({
      day: effectiveDate.getDate(),
      title: summary || 'Evento del calendario Google',
      description,
      location,
      owner: organizer || 'Calendario Google',
      startTime,
      endTime
    });
  });

  return events;
}

/**
 * Une las líneas dobladas dentro de un ICS.
 * @private
 */
function unfoldIcsContent_(content) {
  return content.replace(/\r\n[ \t]/g, '');
}

/**
 * Extrae la línea completa de un campo ICS.
 * @private
 */
function extractIcsLine_(lines, key) {
  const upperKey = key.toUpperCase();
  for (let i = 0; i < lines.length; i += 1) {
    const line = lines[i];
    const upperLine = line.toUpperCase();
    if (upperLine.startsWith(upperKey + ':') || upperLine.startsWith(upperKey + ';')) {
      return line;
    }
  }
  return '';
}

/**
 * Obtiene el valor (tras los dos puntos) de una línea ICS.
 * @private
 */
function getIcsValue_(line) {
  if (!line) {
    return '';
  }
  const index = line.indexOf(':');
  if (index === -1) {
    return '';
  }
  return line.substring(index + 1).trim();
}

/**
 * Verifica si la línea de fecha incluye componente horario.
 * @private
 */
function hasTimeComponent_(line) {
  if (!line) {
    return false;
  }
  return /T\d{6}/.test(line);
}

/**
 * Convierte el valor de fecha de ICS en objeto Date.
 * @private
 */
function parseIcsDateValue_(line) {
  if (!line) {
    return null;
  }

  const colonIndex = line.indexOf(':');
  if (colonIndex === -1) {
    return null;
  }

  const meta = line.substring(0, colonIndex);
  const value = line.substring(colonIndex + 1).trim();
  const tzMatch = /TZID=([^;:]+)/i.exec(meta);
  const timeZone = tzMatch ? tzMatch[1] : Session.getScriptTimeZone();

  try {
    if (/^\d{8}T\d{6}Z$/.test(value)) {
      return Utilities.parseDate(value, 'UTC', "yyyyMMdd'T'HHmmss'Z'");
    }
    if (/^\d{8}T\d{6}$/.test(value)) {
      return Utilities.parseDate(value, timeZone, "yyyyMMdd'T'HHmmss");
    }
    if (/^\d{8}$/.test(value)) {
      return Utilities.parseDate(value, timeZone, 'yyyyMMdd');
    }
  } catch (error) {
    console.warn('No se pudo interpretar la fecha del calendario externo:', value, error);
  }

  return null;
}

/**
 * Calcula la marca de tiempo exclusiva de finalización para un evento ICS.
 * @private
 */
function computeIcsExclusiveEnd_(startDate, rawEndDate) {
  if (!(startDate instanceof Date)) {
    return new Date(0);
  }

  if (rawEndDate instanceof Date) {
    const ensuredEnd = new Date(rawEndDate.getTime());
    if (ensuredEnd.getTime() <= startDate.getTime()) {
      return new Date(startDate.getTime() + 1);
    }
    return ensuredEnd;
  }

  return new Date(startDate.getTime() + 1);
}

/**
 * Formatea un Date a HH:mm en la zona horaria indicada.
 * @private
 */
function formatTimeString_(date, timeZone) {
  if (!(date instanceof Date)) {
    return '';
  }
  return Utilities.formatDate(date, timeZone || Session.getScriptTimeZone(), 'HH:mm');
}

/**
 * Decodifica caracteres escapados en texto ICS.
 * @private
 */
function decodeIcsText_(value) {
  if (!value) {
    return '';
  }
  return value
    .replace(/\\n/gi, '\n')
    .replace(/\\,/g, ',')
    .replace(/\\;/g, ';')
    .replace(/\\\\/g, '\\')
    .trim();
}

/**
 * Extrae el nombre del organizador desde una línea ICS.
 * @private
 */
function parseIcsOrganizer_(line) {
  if (!line) {
    return '';
  }
  const cnMatch = /CN=([^;:]+)/i.exec(line);
  if (cnMatch) {
    return decodeIcsText_(cnMatch[1]);
  }
  const value = getIcsValue_(line);
  if (/^mailto:/i.test(value)) {
    return value.replace(/^mailto:/i, '');
  }
  return decodeIcsText_(value);
}

/**
 * Convierte una cadena HH:mm en minutos para ordenar.
 * @private
 */
function minutesFromTimeString_(value) {
  if (!value) {
    return 24 * 60;
  }
  const match = /^(\d{1,2}):(\d{2})$/.exec(value);
  if (!match) {
    return 24 * 60;
  }
  const hours = Number(match[1]);
  const minutes = Number(match[2]);
  return hours * 60 + minutes;
}

// ---------------------------------------------------------------------------
// Funciones CRUD para Administradores (Fase 4)
// ---------------------------------------------------------------------------

/**
 * Elimina un evento del calendario
 * @param {number} rowIndex - Índice de la fila en la hoja (1-based)
 * @return {Object} Resultado de la operación
 */
function deleteEvent(rowIndex) {
  try {
    ensureCanManage_();

    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName(SHEET_NAMES.events);

    if (!sheet) {
      throw new Error('No se encontró la hoja de eventos');
    }

    if (rowIndex < 2 || rowIndex > sheet.getLastRow()) {
      throw new Error('Índice de fila inválido');
    }

    sheet.deleteRow(rowIndex);

    return {
      success: true,
      message: 'Evento eliminado correctamente'
    };
  } catch (error) {
    console.error('Error al eliminar evento:', error);
    return {
      success: false,
      message: 'Error al eliminar el evento: ' + error.message
    };
  }
}

/**
 * Actualiza un evento existente
 * @param {number} rowIndex - Índice de la fila en la hoja (1-based)
 * @param {Object} payload - Datos del evento actualizado
 * @return {Object} Resultado de la operación
 */
function updateEvent(rowIndex, payload) {
  try {
    ensureCanManage_();

    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName(SHEET_NAMES.events);

    if (!sheet) {
      throw new Error('No se encontró la hoja de eventos');
    }

    if (rowIndex < 2 || rowIndex > sheet.getLastRow()) {
      throw new Error('Índice de fila inválido');
    }

    const eventData = sanitizeEventPayload_(payload);

    // Actualizar las celdas con los nuevos datos (ahora son 8 columnas)
    const range = sheet.getRange(rowIndex, 1, 1, 8);
    range.setValues([[
      eventData.dateStart,
      eventData.dateEnd,
      eventData.startTime,
      eventData.endTime,
      eventData.title,
      eventData.description,
      eventData.location,
      eventData.owner
    ]]);

    return {
      success: true,
      message: 'Evento actualizado correctamente'
    };
  } catch (error) {
    console.error('Error al actualizar evento:', error);
    return {
      success: false,
      message: 'Error al actualizar el evento: ' + error.message
    };
  }
}

/**
 * Elimina un aviso
 * @param {number} rowIndex - Índice de la fila en la hoja (1-based)
 * @return {Object} Resultado de la operación
 */
function deleteNotice(rowIndex) {
  try {
    ensureCanManage_();

    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName(SHEET_NAMES.notices);

    if (!sheet) {
      throw new Error('No se encontró la hoja de avisos');
    }

    if (rowIndex < 2 || rowIndex > sheet.getLastRow()) {
      throw new Error('Índice de fila inválido');
    }

    sheet.deleteRow(rowIndex);

    return {
      success: true,
      message: 'Aviso eliminado correctamente'
    };
  } catch (error) {
    console.error('Error al eliminar aviso:', error);
    return {
      success: false,
      message: 'Error al eliminar el aviso: ' + error.message
    };
  }
}

/**
 * Actualiza un aviso existente
 * @param {number} rowIndex - Índice de la fila en la hoja (1-based)
 * @param {Object} payload - Datos del aviso actualizado
 * @return {Object} Resultado de la operación
 */
function updateNotice(rowIndex, payload) {
  try {
    ensureCanManage_();

    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName(SHEET_NAMES.notices);

    if (!sheet) {
      throw new Error('No se encontró la hoja de avisos');
    }

    if (rowIndex < 2 || rowIndex > sheet.getLastRow()) {
      throw new Error('Índice de fila inválido');
    }

    const notice = sanitizeNoticePayload_(payload);

    // Actualizar las celdas con los nuevos datos
    const range = sheet.getRange(rowIndex, 1, 1, 8);
    range.setValues([[
      notice.type,
      notice.title,
      notice.description,
      notice.date,
      notice.contact,
      notice.priority,
      notice.visibility,
      notice.attachments
    ]]);

    return {
      success: true,
      message: 'Aviso actualizado correctamente'
    };
  } catch (error) {
    console.error('Error al actualizar aviso:', error);
    return {
      success: false,
      message: 'Error al actualizar el aviso: ' + error.message
    };
  }
}

/**
 * Elimina un grupo
 * @param {number} rowIndex - Índice de la fila en la hoja (1-based)
 * @return {Object} Resultado de la operación
 */
function deleteGroup(rowIndex) {
  try {
    ensureCanManage_();

    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName(SHEET_NAMES.groups);

    if (!sheet) {
      throw new Error('No se encontró la hoja de grupos');
    }

    if (rowIndex < 2 || rowIndex > sheet.getLastRow()) {
      throw new Error('Índice de fila inválido');
    }

    sheet.deleteRow(rowIndex);

    return {
      success: true,
      message: 'Grupo eliminado correctamente'
    };
  } catch (error) {
    console.error('Error al eliminar grupo:', error);
    return {
      success: false,
      message: 'Error al eliminar el grupo: ' + error.message
    };
  }
}

/**
 * Actualiza un grupo existente
 * @param {number} rowIndex - Índice de la fila en la hoja (1-based)
 * @param {Object} payload - Datos del grupo actualizado
 * @return {Object} Resultado de la operación
 */
function updateGroup(rowIndex, payload) {
  try {
    ensureCanManage_();

    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName(SHEET_NAMES.groups);

    if (!sheet) {
      throw new Error('No se encontró la hoja de grupos');
    }

    if (rowIndex < 2 || rowIndex > sheet.getLastRow()) {
      throw new Error('Índice de fila inválido');
    }

    const group = sanitizeGroupPayload_(payload);

    // Actualizar las celdas con los nuevos datos
    const range = sheet.getRange(rowIndex, 1, 1, 6);
    range.setValues([[
      group.name,
      group.lead,
      group.contact,
      group.members,
      group.nextMeeting,
      group.notes
    ]]);

    return {
      success: true,
      message: 'Grupo actualizado correctamente'
    };
  } catch (error) {
    console.error('Error al actualizar grupo:', error);
    return {
      success: false,
      message: 'Error al actualizar el grupo: ' + error.message
    };
  }
}
