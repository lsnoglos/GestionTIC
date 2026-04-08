const CONFIG = {
  APP_NAME: 'Mesa de Soporte TIC',
  USERS_SHEET: 'Usuarios',
  LOGS_SHEET: 'Registros',
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
  ]
};

const USER_HEADERS = [
  'referencia',
  'tipo_usuario',
  'nombre_apellido',
  'carrera',
  'celular',
  'fecha_nacimiento',
  'pin',
  'firma_file_id',
  'firma_url',
  'fecha_creacion',
  'fecha_actualizacion'
];

const LOG_HEADERS = [
  'fecha',
  'hora',
  'descripcion',
  'nombre_apellido',
  'referencia',
  'observacion',
  'firma_url',
  'fecha_hora_iso'
];

const USER_TYPE_MAP = {
  A: 'alumno',
  ALUMNO: 'alumno',
  ALUMNA: 'alumno',
  E: 'alumno',
  ESTUDIANTE: 'alumno',
  M: 'maestro',
  MAESTRO: 'maestro',
  MAESTRA: 'maestro',
  DOCENTE: 'maestro',
  PROFESOR: 'maestro'
};

function doGet() {
  initializeSheets_();
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle(CONFIG.APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getAppBootstrap() {
  initializeSheets_();
  return {
    appName: CONFIG.APP_NAME,
    careers: CONFIG.CAREERS
  };
}

function findUserAndHistory(reference) {
  initializeSheets_();
  const cleanRef = sanitizeReference_(reference);
  if (!cleanRef) {
    return { found: false, error: 'Debes escribir un carnet o cédula.' };
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
    if (!cleanRef) throw new Error('Referencia inválida.');

    const reason = String(payload.reason || '').trim();
    const otherReason = String(payload.otherReason || '').trim();
    const finalReason = reason === 'Otros' ? otherReason : reason;
    if (!finalReason) throw new Error('Debes seleccionar o escribir el motivo.');

    const observation = String(payload.observation || '').trim();
    const inputPin = String(payload.pin || '').trim();

    const existingUser = findUserByReference_(cleanRef);
    let resolvedUser;

    if (existingUser) {
      if (!inputPin) throw new Error('Debes ingresar tu clave para continuar.');
      if (String(existingUser.pin || '') !== inputPin) {
        throw new Error('La clave no coincide con el registro existente.');
      }
      resolvedUser = updateExistingUser_(existingUser, payload);
    } else {
      resolvedUser = createUser_(cleanRef, payload);
    }

    appendLog_(resolvedUser, finalReason, observation);

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
  ensureSheetHeaders_(ss, CONFIG.LOGS_SHEET, LOG_HEADERS);
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

function createUser_(reference, payload) {
  const cleanName = String(payload.fullName || '').trim();
  if (!cleanName) throw new Error('Debes ingresar el nombre y apellido.');

  const userType = inferUserType_(reference);
  const career = userType === 'alumno' ? String(payload.career || '').trim() : '';
  if (userType === 'alumno' && !career) {
    throw new Error('Debes seleccionar la carrera del alumno.');
  }

  const birthDate = resolveBirthDate_(reference, payload.birthDate);
  const phone = String(payload.phone || '').trim();
  const pin = String(payload.pin || '').trim();
  const signatureDataUrl = String(payload.signatureDataUrl || '').trim();

  if (!pin) throw new Error('Debes crear una clave.');
  if (!signatureDataUrl) throw new Error('Debes dibujar y guardar la firma para crear el usuario.');

  const now = new Date();
  const signature = saveOrReuseSignature_(reference, null, signatureDataUrl);

  const userObject = {
    reference,
    userType,
    fullName: cleanName,
    career,
    phone,
    birthDate,
    pin,
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
    userObject.career,
    userObject.phone,
    userObject.birthDate,
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

  let signature = {
    fileId: existingUser.signatureFileId,
    url: existingUser.signatureUrl
  };

  const incomingSignature = String(payload.signatureDataUrl || '').trim();
  const wantsUpdateSignature = Boolean(incomingSignature);
  if (wantsUpdateSignature) {
    signature = saveOrReuseSignature_(existingUser.reference, existingUser.signatureFileId, incomingSignature);
  } else if (!signature.fileId && !signature.url) {
    throw new Error('Este usuario no tiene firma registrada. Dibuja una firma para continuar.');
  }

  const updatedAt = new Date();
  const latest = {
    ...existingUser,
    userType: normalizeUserType_(existingUser.userType, existingUser.reference),
    signatureFileId: signature.fileId,
    signatureUrl: signature.url,
    updatedAt
  };

  sheet.getRange(row, 2).setValue(latest.userType);
  sheet.getRange(row, 8).setValue(latest.signatureFileId || '');
  sheet.getRange(row, 9).setValue(latest.signatureUrl || '');
  sheet.getRange(row, 11).setValue(updatedAt);

  return latest;
}

function appendLog_(user, reason, observation) {
  const now = new Date();
  const formattedDate = Utilities.formatDate(now, CONFIG.TZ, 'yyyy-MM-dd');
  const formattedTime = Utilities.formatDate(now, CONFIG.TZ, 'HH:mm:ss');

  const signatureUrl = String(user.signatureUrl || '').trim();
  const signatureCellValue = signatureUrl ? buildImageFormula_(signatureUrl) : '';

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.LOGS_SHEET);
  sheet.appendRow([
    formattedDate,
    formattedTime,
    reason,
    user.fullName,
    user.reference,
    observation,
    signatureCellValue,
    now.toISOString()
  ]);
}

function buildImageFormula_(url) {
  const safeUrl = String(url || '').replace(/"/g, '""');
  return `=IMAGE("${safeUrl}")`;
}

function getHistoryByReference_(reference) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.LOGS_SHEET);
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];

  const headers = rows.shift().map(h => String(h || '').trim());
  const idxRef = headers.indexOf('referencia');
  const idxDate = headers.indexOf('fecha');
  const idxTime = headers.indexOf('hora');
  const idxDesc = headers.indexOf('descripcion');
  const idxObs = headers.indexOf('observacion');

  return rows
    .filter(r => String(r[idxRef] || '').trim().toUpperCase() === reference.toUpperCase())
    .map(r => ({
      fecha: r[idxDate],
      hora: r[idxTime],
      descripcion: r[idxDesc],
      observacion: r[idxObs]
    }))
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
    career: headers.indexOf('carrera'),
    phone: headers.indexOf('celular'),
    birthDate: headers.indexOf('fecha_nacimiento'),
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
    userType: normalizeUserType_(row[idx.userType], row[idx.reference]),
    fullName: String(row[idx.fullName] || ''),
    career: String(row[idx.career] || ''),
    phone: String(row[idx.phone] || ''),
    birthDate: String(row[idx.birthDate] || ''),
    pin: String(row[idx.pin] || ''),
    signatureFileId: String(row[idx.signatureFileId] || ''),
    signatureUrl: String(row[idx.signatureUrl] || ''),
    _row: rowIndex + 2
  };
}

function inferUserType_(reference) {
  return isCedula_(reference) ? 'maestro' : 'alumno';
}

function normalizeUserType_(rawType, reference) {
  const key = String(rawType || '').trim().toUpperCase();
  if (USER_TYPE_MAP[key]) return USER_TYPE_MAP[key];
  return inferUserType_(sanitizeReference_(reference));
}

function sanitizeReference_(reference) {
  return String(reference || '').trim().toUpperCase();
}

function isCedula_(reference) {
  return /^\d{3}-\d{6}-\d{4}[A-Z]$/i.test(String(reference || '').trim());
}

function resolveBirthDate_(reference, inputBirthDate) {
  if (isCedula_(reference)) {
    return birthDateFromCedula_(reference);
  }

  const manual = String(inputBirthDate || '').trim();
  if (!manual) throw new Error('Debes ingresar la fecha de nacimiento para carnet.');
  return manual;
}

function birthDateFromCedula_(cedula) {
  const parts = cedula.split('-');
  const block = parts[1] || '';
  const day = block.substring(0, 2);
  const month = block.substring(2, 4);
  const year2 = Number(block.substring(4, 6));

  if (!day || !month || Number.isNaN(year2)) return '';

  const currentYear2 = Number(Utilities.formatDate(new Date(), CONFIG.TZ, 'yy'));
  const year = year2 > currentYear2 ? 1900 + year2 : 2000 + year2;

  return `${year}-${month}-${day}`;
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

function activarPermisosDrive() {
  DriveApp.getRootFolder();
}

function testDriveDirecto() {
  const folder = DriveApp.getFolderById("10yZgZY3MfnCQAhh4RtTuusQ-RjDWUUu-");
  folder.createFile("test.txt", "ok");
}
