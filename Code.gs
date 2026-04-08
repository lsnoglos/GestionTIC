const CONFIG = {
  APP_NAME: 'Mesa de Soporte TIC',
  USERS_SHEET: 'Usuarios',
  LOG_SHEETS: ['Soporte', 'Sistemas', 'Otros'],
  TZ: Session.getScriptTimeZone() || 'America/Managua',
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
  'carrera_area',
  'pin',
  'firma_file_id',
  'firma_url',
  'fecha_creacion',
  'fecha_actualizacion'
];

const LOG_HEADERS = [
  'usuario',
  'Nombre',
  'Motivo',
  'Observación',
  'Comentarios',
  'Área',
  'Firma',
  'fecha',
  'hora'
];

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
  const cleanRef = sanitizeReference_(reference);
  if (!cleanRef) {
    return { found: false, error: 'Debes escribir un usuario.' };
  }

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
    const cleanRef = sanitizeReference_(payload.reference);
    if (!cleanRef) throw new Error('Usuario inválido.');

    const category = String(payload.category || '').trim();
    const reason = String(payload.reason || '').trim();
    const otherComment = String(payload.otherComment || '').trim();
    if (!category) throw new Error('Debes seleccionar la categoría.');
    if (!reason) throw new Error('Debes seleccionar el motivo de asistencia.');
    if (reason === 'Otros' && !otherComment) {
      throw new Error('Debes escribir el comentario para "Otros".');
    }

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

    appendLog_(resolvedUser, {
      category,
      reason,
      otherComment,
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
    const incomingCareerArea = String(payload.careerArea || '').trim();

    if (!incomingName) throw new Error('Debes ingresar el nombre y apellido.');
    if (!incomingCareerArea) throw new Error('Debes seleccionar carrera o área.');
    if (!isValidCareerArea_(latest.userType, incomingCareerArea)) {
      throw new Error('La carrera o área seleccionada no es válida.');
    }

    latest.fullName = incomingName;
    latest.careerArea = incomingCareerArea;

    sheet.getRange(row, 3).setValue(latest.fullName);
    sheet.getRange(row, 4).setValue(latest.careerArea);
  }

  sheet.getRange(row, 6).setValue(latest.signatureFileId || '');
  sheet.getRange(row, 7).setValue(latest.signatureUrl || '');
  sheet.getRange(row, 9).setValue(updatedAt);

  return latest;
}

function appendLog_(user, details) {
  const now = new Date();
  const formattedDate = Utilities.formatDate(now, CONFIG.TZ, 'yyyy-MM-dd');
  const formattedTime = Utilities.formatDate(now, CONFIG.TZ, 'HH:mm:ss');

  const signatureUrl = String(user.signatureUrl || '').trim();
  const signatureCellValue = signatureUrl ? buildImageFormula_(signatureUrl) : '';

  const categorySheet = normalizeCategorySheet_(details.category);
  if (!categorySheet) throw new Error('Categoría inválida.');

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(categorySheet);
  sheet.appendRow([
    user.reference,
    user.fullName,
    details.reason,
    details.observation,
    details.reason === 'Otros' ? details.otherComment : '',
    user.careerArea,
    signatureCellValue,
    formattedDate,
    formattedTime
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

  CONFIG.LOG_SHEETS.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const rows = sheet.getDataRange().getValues();
    if (rows.length <= 1) return;

    const headers = rows.shift().map(h => String(h || '').trim());
    const idxRef = headers.indexOf('usuario');
    const idxDate = headers.indexOf('fecha');
    const idxTime = headers.indexOf('hora');
    const idxReason = headers.indexOf('Motivo');
    const idxComment = headers.indexOf('Comentarios');
    const idxObs = headers.indexOf('Observación');

    rows
      .filter(r => String(r[idxRef] || '').trim().toUpperCase() === reference.toUpperCase())
      .forEach(r => {
        const reason = String(r[idxReason] || '');
        const extraComment = String(r[idxComment] || '');
        const reasonText = reason === 'Otros' && extraComment ? `${reason}: ${extraComment}` : reason;
        allItems.push({
          fecha: r[idxDate],
          hora: r[idxTime],
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

  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];

  const headers = rows.shift().map(h => String(h || '').trim());
  const idxRef = headers.indexOf('referencia');
  const idxDate = headers.indexOf('fecha');
  const idxTime = headers.indexOf('hora');
  const idxCategory = headers.indexOf('categoria');
  const idxReason = headers.indexOf('motivo');
  const idxComment = headers.indexOf('comentario_otros');
  const idxObs = headers.indexOf('observacion');

  return rows
    .filter(r => String(r[idxRef] || '').trim().toUpperCase() === reference.toUpperCase())
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

function findUserByReference_(reference) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.USERS_SHEET);
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return null;

  const headers = rows.shift().map(h => String(h || '').trim());
  const idx = {
    reference: headers.indexOf('referencia'),
    userType: headers.indexOf('tipo_usuario'),
    fullName: headers.indexOf('nombre_apellido'),
    careerArea: headers.indexOf('carrera_area'),
    pin: headers.indexOf('pin'),
    signatureFileId: headers.indexOf('firma_file_id'),
    signatureUrl: headers.indexOf('firma_url')
  };

  const rowIndex = rows.findIndex(
    r => String(r[idx.reference] || '').trim().toUpperCase() === reference.toUpperCase()
  );
  if (rowIndex === -1) return null;

  const row = rows[rowIndex];
  return {
    reference: String(row[idx.reference] || ''),
    userType: String(row[idx.userType] || ''),
    fullName: String(row[idx.fullName] || ''),
    careerArea: String(row[idx.careerArea] || ''),
    pin: String(row[idx.pin] || ''),
    signatureFileId: String(row[idx.signatureFileId] || ''),
    signatureUrl: String(row[idx.signatureUrl] || ''),
    _row: rowIndex + 2
  };
}

function inferUserType_(reference) {
  return isInstitutionalEmail_(reference) ? 'personal' : 'alumno';
}

function sanitizeReference_(reference) {
  return String(reference || '').trim().toUpperCase();
}

function isInstitutionalEmail_(value) {
  return /^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}$/i.test(String(value || '').trim());
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
