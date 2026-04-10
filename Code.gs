const CONFIG = {
  APP_NAME: 'Mesa de Soporte TIC',
  USERS_SHEET: 'Usuarios',
  LOG_SHEETS: ['Soporte', 'Sistemas', 'Otros'],
  TZ: Session.getScriptTimeZone() || 'America/Managua',
  DEFAULT_EMAIL_DOMAIN: 'UCC.EDU.NI',
  SIGNATURE_FOLDER_ID: '10yZgZY3MfnCQAhh4RtTuusQ-RjDWUUu-',
  CAREERS: [
    'Ing. de Sistemas',
    'Ing. Civil',
    'Agronomía',
    'Arquitectura',
    'Diseño gráfico',
    'Marketing y publicidad',
    'Derecho',
    'Contabilidad',
    'Contaduría',
    'Ing. Industrial'
  ],
  AREAS: [
    'Ingeniería Civil',
    'CCEEyJJ',
    'Investigación',
    'Posgrado y EC',
    'BE - Proy. Social',
    'Biblioteca',
    'Supervisión Metodológica',
    'Dirección Académica',
    'Comunicación Institucional',
    'Recursos Humanos',
    'Registro Académico',
    'Ingeniería Agronómica',
    'Diseño Gráfico y Arq',
    'TIC'
  ]
};

const USER_HEADERS = [
  'referencia',
  'tipo_usuario',
  'nombre_apellido',
  'genero',
  'carrera_area',
  'pin',
  'firma_file_id',
  'firma_url',
  'fecha_creacion',
  'fecha_actualizacion'
];

const LOG_HEADERS = ['Fecha', 'Nombre y apellido', 'Género', 'Referencia', 'Motivo', 'Observación', 'Firma'];

function doGet() {
  initializeSheets_();
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle(CONFIG.APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getAppBootstrap() {
  initializeSheets_();
  return {
    appName: CONFIG.APP_NAME,
    careers: CONFIG.CAREERS,
    areas: CONFIG.AREAS
  };
}

function findUserAndHistory(reference) {
  initializeSheets_();
  const cleanRef = normalizeInputReference(reference);
  if (!cleanRef) {
    return { found: false, error: 'Debes escribir un usuario.' };
  }
  Logger.log('BUSCAR: ' + cleanRef);

  const user = findUserByReference_(cleanRef);
  const history = getHistoryByReference_(cleanRef);

  return {
    found: Boolean(user),
    reference: cleanRef,
    inferredType: inferUserType_(cleanRef),
    user,
    history
  };
}

function saveVisit(payload) {
  initializeSheets_();
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    const cleanRef = normalizeInputReference(payload.reference);
    if (!cleanRef) throw new Error('Usuario inválido.');
    Logger.log('SAVE: ' + cleanRef);

    const category = String(payload.category || '').trim();
    const reason = String(payload.reason || '').trim();
    if (!category) throw new Error('Debes seleccionar la categoría.');
    if (!reason) throw new Error('Debes seleccionar el motivo de asistencia.');

    const observation = String(payload.observation || '').trim();
    const inputPin = String(payload.pin || '').trim();
    if (!inputPin) throw new Error('Debes ingresar tu clave para continuar.');

    const existingUser = findUserByReference_(cleanRef);
    let resolvedUser;

    if (existingUser) {
      if (String(existingUser.pin || '') !== inputPin) {
        throw new Error('La clave no coincide con el registro existente.');
      }
      resolvedUser = updateExistingUser_(existingUser, payload);
    } else {
      resolvedUser = createUser_(cleanRef, payload, inputPin);
    }

    if (!normalizeGender_(resolvedUser.gender)) {
      throw new Error('Este usuario no tiene género registrado. Activa edición y selecciona M o F.');
    }

    appendLog_(resolvedUser, {
      category,
      reason,
      observation
    });

    return {
      success: true,
      message: 'Registro guardado correctamente.',
      user: resolvedUser,
      history: getHistoryByReference_(cleanRef)
    };
  } finally {
    lock.releaseLock();
  }
}

function initializeSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheetHeaders_(ss, CONFIG.USERS_SHEET, USER_HEADERS);
  CONFIG.LOG_SHEETS.forEach(sheetName => ensureSheetHeaders_(ss, sheetName, LOG_HEADERS));
}

function ensureSheetHeaders_(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);

  const currentHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const needsHeader = headers.some((h, i) => String(currentHeaders[i] || '') !== h);

  if (needsHeader) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}

function createUser_(reference, payload, inputPin) {
  const cleanName = String(payload.fullName || '').trim();
  if (!cleanName) throw new Error('Debes ingresar el nombre y apellido.');
  const gender = normalizeGender_(payload.gender);
  if (!gender) throw new Error('Debes seleccionar el género (M o F).');

  const userType = inferUserType_(reference);
  const careerArea = String(payload.careerArea || '').trim();
  if (!careerArea) {
    throw new Error(userType === 'alumno' ? 'Debes seleccionar la carrera.' : 'Debes seleccionar el área.');
  }

  if (!isValidCareerArea_(userType, careerArea)) {
    throw new Error('La carrera o área seleccionada no es válida.');
  }

  const signatureDataUrl = String(payload.signatureDataUrl || '').trim();
  if (!signatureDataUrl) throw new Error('Debes dibujar la firma para crear el usuario.');

  const now = new Date();
  const signature = saveOrReuseSignature_(reference, null, signatureDataUrl);

  const userObject = {
    reference,
    userType,
    fullName: cleanName,
    gender,
    careerArea,
    pin: inputPin,
    signatureFileId: signature.fileId,
    signatureUrl: signature.url,
    createdAt: now,
    updatedAt: now
  };

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.USERS_SHEET);
  sheet.appendRow([
    userObject.reference,
    userObject.userType,
    userObject.fullName,
    userObject.gender,
    userObject.careerArea,
    userObject.pin,
    userObject.signatureFileId,
    userObject.signatureUrl,
    userObject.createdAt,
    userObject.updatedAt
  ]);

  return userObject;
}

function updateExistingUser_(existingUser, payload) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.USERS_SHEET);
  const row = existingUser._row;
  const allowProfileEdit = Boolean(payload.allowProfileEdit);

  let signature = {
    fileId: existingUser.signatureFileId,
    url: existingUser.signatureUrl
  };

  const incomingSignature = String(payload.signatureDataUrl || '').trim();
  if (allowProfileEdit && incomingSignature) {
    signature = saveOrReuseSignature_(existingUser.reference, existingUser.signatureFileId, incomingSignature);
  } else if (!signature.fileId && !signature.url) {
    throw new Error('Este usuario no tiene firma registrada. Activa edición y dibuja una firma.');
  }

  const updatedAt = new Date();
  const latest = {
    ...existingUser,
    signatureFileId: signature.fileId,
    signatureUrl: signature.url,
    updatedAt
  };

  if (allowProfileEdit) {
    const incomingName = String(payload.fullName || '').trim();
    const incomingGender = normalizeGender_(payload.gender);
    const incomingCareerArea = String(payload.careerArea || '').trim();

    if (!incomingName) throw new Error('Debes ingresar el nombre y apellido.');
    if (!incomingGender) throw new Error('Debes seleccionar el género (M o F).');
    if (!incomingCareerArea) throw new Error('Debes seleccionar carrera o área.');
    if (!isValidCareerArea_(latest.userType, incomingCareerArea)) {
      throw new Error('La carrera o área seleccionada no es válida.');
    }

    latest.fullName = incomingName;
    latest.gender = incomingGender;
    latest.careerArea = incomingCareerArea;

    sheet.getRange(row, 3).setValue(latest.fullName);
    sheet.getRange(row, 4).setValue(latest.gender);
    sheet.getRange(row, 5).setValue(latest.careerArea);
  }

  sheet.getRange(row, 7).setValue(latest.signatureFileId || '');
  sheet.getRange(row, 8).setValue(latest.signatureUrl || '');
  sheet.getRange(row, 10).setValue(updatedAt);

  return latest;
}

function appendLog_(user, details) {
  const now = new Date();
  const formattedDateTime = Utilities.formatDate(now, CONFIG.TZ, 'yyyy-MM-dd HH:mm:ss');

  const signatureUrl = String(user.signatureUrl || '').trim();
  const signatureCellValue = signatureUrl ? buildImageFormula_(signatureUrl) : '';

  const categorySheet = normalizeCategorySheet_(details.category);
  if (!categorySheet) throw new Error('Categoría inválida.');

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(categorySheet);
  sheet.appendRow([
    formattedDateTime,
    user.fullName,
    user.gender || '',
    user.reference,
    details.reason,
    details.observation,
    signatureCellValue
  ]);
}

function normalizeCategorySheet_(category) {
  const raw = String(category || '').trim().toLowerCase();
  if (!raw) return '';
  if (raw === 'soporte' || raw === 'soporte técnico') return 'Soporte';
  if (raw === 'sistemas') return 'Sistemas';
  if (raw === 'otros') return 'Otros';
  return '';
}

function buildImageFormula_(url) {
  const safeUrl = String(url || '').replace(/"/g, '""');
  return `=IMAGE("${safeUrl}")`;
}

function getHistoryByReference_(reference) {
  const currentHistory = getHistoryByReferenceFromCategorySheets_(reference);
  if (currentHistory.length) return currentHistory;
  return getLegacyLogHistoryByReference_(reference);
}

function getHistoryByReferenceFromCategorySheets_(reference) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allItems = [];
  const normalizedReference = normalizeReferenceKey_(reference);

  CONFIG.LOG_SHEETS.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const rows = sheet.getDataRange().getValues();
    if (rows.length <= 1) return;

    const headers = rows.shift();
    const headerMap = buildHeaderIndexMap_(headers);
    const idxRef = findHeaderIndex_(headerMap, ['Referencia', 'referencia', 'usuario']);
    const idxDateTime = findHeaderIndex_(headerMap, ['Fecha', 'fecha', 'fecha_hora', 'datetime']);
    const idxDate = findHeaderIndex_(headerMap, ['fecha', 'date']);
    const idxTime = findHeaderIndex_(headerMap, ['hora', 'time']);
    const idxReason = findHeaderIndex_(headerMap, ['Motivo', 'motivo', 'razon', 'reason']);
    const idxComment = findHeaderIndex_(headerMap, ['Comentarios', 'comentarios', 'comentario_otros', 'comentario']);
    const idxObs = findHeaderIndex_(headerMap, ['Observación', 'observacion', 'obs']);

    if (idxRef === -1 || idxReason === -1) return;

    rows
      .filter(r => normalizeReferenceKey_(r[idxRef]) === normalizedReference)
      .forEach(r => {
        const reason = String(r[idxReason] || '');
        const extraComment = String(r[idxComment] || '');
        const reasonText = reason === 'Otros' && extraComment ? `${reason}: ${extraComment}` : reason;
        const dateTimeValue = idxDateTime !== -1
          ? r[idxDateTime]
          : `${String(r[idxDate] || '').trim()} ${String(r[idxTime] || '').trim()}`.trim();
        allItems.push({
          fechaHora: dateTimeValue,
          categoria: sheetName,
          motivo: reasonText,
          observacion: r[idxObs]
        });
      });
  });

  return allItems.reverse();
}

function getLegacyLogHistoryByReference_(reference) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Registros');
  if (!sheet) return [];
  const normalizedReference = normalizeReferenceKey_(reference);

  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];

  const headers = rows.shift();
  const headerMap = buildHeaderIndexMap_(headers);
  const idxRef = findHeaderIndex_(headerMap, ['referencia', 'usuario']);
  const idxDate = findHeaderIndex_(headerMap, ['fecha', 'date']);
  const idxTime = findHeaderIndex_(headerMap, ['hora', 'time']);
  const idxCategory = findHeaderIndex_(headerMap, ['categoria', 'categoría', 'category']);
  const idxReason = findHeaderIndex_(headerMap, ['motivo', 'razon', 'reason']);
  const idxComment = findHeaderIndex_(headerMap, ['comentario_otros', 'comentarios', 'comentario']);
  const idxObs = findHeaderIndex_(headerMap, ['observacion', 'observación', 'obs']);

  return rows
    .filter(r => normalizeReferenceKey_(r[idxRef]) === normalizedReference)
    .map(r => {
      const reason = String(r[idxReason] || '');
      const extraComment = String(r[idxComment] || '');
      const reasonText = reason === 'Otros' && extraComment ? `${reason}: ${extraComment}` : reason;
      return {
        fecha: r[idxDate],
        hora: r[idxTime],
        categoria: r[idxCategory],
        motivo: reasonText,
        observacion: r[idxObs]
      };
    })
    .reverse();
}


function buildHeaderIndexMap_(headers) {
  const map = {};
  headers.forEach((h, i) => {
    map[normalizeHeaderKey_(h)] = i;
  });
  return map;
}

function findHeaderIndex_(headerMap, aliases) {
  for (let i = 0; i < aliases.length; i += 1) {
    const idx = headerMap[normalizeHeaderKey_(aliases[i])];
    if (typeof idx === 'number') return idx;
  }
  return -1;
}

function normalizeHeaderKey_(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .replace(/\s+/g, '_');
}

function findUserByReference_(reference) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.USERS_SHEET);
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return null;

  const headers = rows.shift();
  const headerMap = buildHeaderIndexMap_(headers);
  const idx = {
    reference: findHeaderIndex_(headerMap, ['referencia', 'Referencia', 'usuario']),
    userType: findHeaderIndex_(headerMap, ['tipo_usuario', 'tipo de usuario']),
    fullName: findHeaderIndex_(headerMap, ['nombre_apellido', 'nombre y apellido']),
    gender: findHeaderIndex_(headerMap, ['genero', 'género']),
    careerArea: findHeaderIndex_(headerMap, ['carrera_area', 'carrera/area', 'carrera_area']),
    pin: findHeaderIndex_(headerMap, ['pin', 'clave', 'contraseña', 'contrasena']),
    signatureFileId: findHeaderIndex_(headerMap, ['firma_file_id', 'firma id', 'firma_file']),
    signatureUrl: findHeaderIndex_(headerMap, ['firma_url', 'firma url'])
  };

  const referenceColumn = idx.reference !== -1 ? idx.reference : 0;

  const normalizedReference = normalizeReferenceKey_(reference);
  const rowIndex = rows.findIndex(r => normalizeReferenceKey_(String(r[referenceColumn])) === normalizedReference);
  if (rowIndex === -1) return null;

  const row = rows[rowIndex];
  return {
    reference: String(row[idx.reference] || ''),
    userType: String(row[idx.userType] || ''),
    fullName: String(row[idx.fullName] || ''),
    gender: String(row[idx.gender] || ''),
    careerArea: String(row[idx.careerArea] || ''),
    pin: String(row[idx.pin] || ''),
    signatureFileId: String(row[idx.signatureFileId] || ''),
    signatureUrl: String(row[idx.signatureUrl] || ''),
    _row: rowIndex + 2
  };
}

function normalizeInputReference(ref) {
  return sanitizeReference_(ref);
}

function inferUserType_(reference) {
  return looksLikeInstitutionalEmail_(reference) ? 'personal' : 'alumno';
}

function sanitizeReference_(reference) {
  if (reference === null || reference === undefined) return '';

  const raw = String(reference)
    .replace(/\u00A0/g, ' ')
    .trim()
    .toUpperCase();

  return normalizeInstitutionalEmail_(raw);
}

function normalizeReferenceKey_(reference) {
  return normalizeInstitutionalEmail_(String(reference || ''))
    .trim()
    .toUpperCase()
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .replace(/^'+/, '')
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .replace(/\s+/g, '');
}

function normalizeGender_(value) {
  const gender = String(value || '').trim().toUpperCase();
  return gender === 'M' || gender === 'F' ? gender : '';
}

function isInstitutionalEmail_(value) {
  return /^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}$/i.test(String(value || '').trim());
}

function looksLikeInstitutionalEmail_(value) {
  const raw = String(value || '').trim();
  if (!raw) return false;
  if (isInstitutionalEmail_(raw)) return true;
  return /^[A-Z0-9._%+-]+@$/i.test(raw);
}

function normalizeInstitutionalEmail_(value) {
  const raw = String(value || '')
    .replace(/\u00A0/g, ' ')
    .trim()
    .toUpperCase();

  if (!raw.includes('@')) return raw;

  const parts = raw.split('@');
  const localPart = String(parts.shift() || '').trim();
  const domainPart = String(parts.join('@') || '').trim();
  if (!localPart) return raw;

  if (!domainPart) return `${localPart}@${CONFIG.DEFAULT_EMAIL_DOMAIN}`;
  return `${localPart}@${domainPart}`;
}

function isValidCareerArea_(userType, value) {
  const list = userType === 'alumno' ? CONFIG.CAREERS : CONFIG.AREAS;
  return list.indexOf(value) !== -1;
}

function saveOrReuseSignature_(reference, existingFileId, signatureDataUrl) {
  if (!signatureDataUrl) {
    return {
      fileId: existingFileId || '',
      url: existingFileId ? drivePublicUrl_(existingFileId) : ''
    };
  }

  const fileId = saveSignatureImage_(reference, signatureDataUrl, existingFileId);
  return { fileId, url: drivePublicUrl_(fileId) };
}

function saveSignatureImage_(reference, signatureDataUrl, existingFileId) {
  const base64 = String(signatureDataUrl || '').split(',')[1];
  if (!base64) throw new Error('No se pudo procesar la firma.');

  const blob = Utilities.newBlob(
    Utilities.base64Decode(base64),
    'image/png',
    `firma_${reference}_${Date.now()}.png`
  );

  const folder = getSignatureFolder_();

  if (existingFileId) {
    try {
      DriveApp.getFileById(existingFileId).setTrashed(true);
    } catch (error) {
      Logger.log('No se pudo reemplazar firma anterior: ' + error.message);
    }
  }

  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getId();
}

function getSignatureFolder_() {
  const configuredFolderId = extractDriveId_(CONFIG.SIGNATURE_FOLDER_ID);
  if (configuredFolderId) {
    return DriveApp.getFolderById(configuredFolderId);
  }

  const folders = DriveApp.getFoldersByName('Firmas_GestionTIC');
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder('Firmas_GestionTIC');
}

function drivePublicUrl_(fileId) {
  return fileId ? `https://drive.google.com/uc?export=view&id=${fileId}` : '';
}

function extractDriveId_(value) {
  const raw = String(value || '').trim();
  if (!raw) return '';
  const matched = raw.match(/[-\w]{25,}/);
  return matched ? matched[0] : raw;
}
