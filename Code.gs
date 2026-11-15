/**
 * ============================================================================
 * HEART MATE WEB APP — GOOGLE APPS SCRIPT API
 * ============================================================================
 * Version: 3.0.0
 * Runtime: Google Apps Script V8
 * Database: Google Sheets
 * Image hosting: ImgBB API v1
 * Music search: YouTube Data API v3
 * Frontend: GitHub Pages (index.html / dashboard.html)
 *
 * IMPORTANT SETUP
 * 1) Paste this entire file into Code.gs in a Google Apps Script project.
 * 2) Run setupSystem() once and grant the requested permissions.
 *    - If this is a spreadsheet-bound script, it uses that spreadsheet.
 *    - If this is a standalone script, it creates the database spreadsheet.
 * 3) Save secrets in Script Properties:
 *    GIFT_MEMORY_YOUTUBE_API_KEY = your YouTube Data API v3 key
 *    GIFT_MEMORY_IMGBB_API_KEY = your ImgBB API key
 * 4) Deploy > New deployment > Web app:
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5) Use the deployed /exec URL in the frontend as API_URL.
 *
 * FRONTEND TRANSPORT
 * GitHub Pages and Apps Script are different origins. This API intentionally
 * supports JSONP through `callback=` so the frontend can call it without a
 * browser CORS preflight. The index/dashboard integration should use the
 * JSONP helper described in the API CONTRACT near the bottom of this file.
 * For image uploads, the browser sends small JSONP chunks to temporary
 * Apps Script CacheService storage. Apps Script then POSTs the assembled image
 * to ImgBB; only ImgBB image URLs are saved in Google Sheets. No image files
 * are stored in Google Drive by this version.
 *
 * SECURITY NOTES
 * - Admin password and YouTube API key are stored in Script Properties.
 * - Admin actions require a short-lived session token.
 * - Login uses a one-time challenge proof; the raw password is not sent in the
 *   request URL.
 * - Public gift URLs are unguessable IDs, not recipient accounts/passwords.
 * - Anyone with a published gift URL can view that gift. Do not store private
 *   or highly sensitive information in a public gift.
 * - JSONP is suitable for this small personal project. Admin API calls still
 *   require a token and the frontend should keep that token in memory only.
 * - Public creation is intentionally open. Nonce, honeypot and throttles are
 *   best-effort abuse controls; use CAPTCHA / edge protection for high traffic.
 * ============================================================================
 */

/* -------------------------------------------------------------------------- */
/* 01. CONFIGURATION                                                          */
/* -------------------------------------------------------------------------- */

const APP = Object.freeze({
  NAME: 'Heart Mate',
  VERSION: '3.0.0',
  DEFAULT_ADMIN_PASSWORD: 'ASSD55290',
  SESSION_TTL_SECONDS: 21600, // 6 hours (CacheService maximum)
  LOGIN_CHALLENGE_TTL_SECONDS: 300,
  LOGIN_FAIL_TTL_SECONDS: 900,
  MAX_LOGIN_FAILURES: 10,
  YOUTUBE_CACHE_TTL_SECONDS: 1800,
  YOUTUBE_MAX_RESULTS: 10,
  MAX_IMAGE_BYTES: 1000000, // Compress images in the browser before upload.
  IMAGE_CHUNK_SIZE: 4000,  // Base64 characters per JSONP request; keep in sync with both HTML files.
  MAX_IMAGE_CHUNKS: 340,   // Enough for a 1 MB image encoded as Base64 at 4,000 chars/chunk.
  IMAGE_UPLOAD_CACHE_TTL_SECONDS: 21600, // CacheService maximum TTL: 6 hours.
  IMAGE_UPLOAD_SESSION_TTL_MS: 21600000, // Upload sessions expire after 6 hours.
  MAX_TEXT_LENGTH: 5000,
  SHEETS: Object.freeze({
    GIFTS: 'Gifts',
    CARDS: 'Cards',
    SONGS: 'Songs',
    PHOTOS: 'Photos',
    MEMORIES: 'Memories',
    SETTINGS: 'Settings'
  }),
  HEADERS: Object.freeze({
    Gifts: [
      'gift_id', 'gift_type', 'title', 'recipient_name', 'sender_name',
      'status', 'theme', 'message', 'payload_json', 'created_at',
      'updated_at', 'view_count', 'last_viewed_at'
    ],
    Cards: [
      'gift_id', 'design_id', 'greeting', 'message', 'closing',
      'created_at', 'updated_at'
    ],
    Songs: [
      'gift_id', 'youtube_video_id', 'video_title', 'channel_title',
      'thumbnail_url', 'watch_url', 'embed_url', 'created_at', 'updated_at'
    ],
    Photos: [
      'photo_id', 'gift_id', 'image_url', 'thumbnail_url', 'caption',
      'photo_type', 'created_at', 'imgbb_image_id', 'imgbb_url_viewer'
    ],
    Memories: [
      'memory_id', 'gift_id', 'memory_title', 'memory_text',
      'memory_date', 'photo_url', 'created_at'
    ],
    Settings: ['setting_key', 'setting_value', 'updated_at']
  })
});

const PROP = Object.freeze({
  SPREADSHEET_ID: 'GIFT_MEMORY_SPREADSHEET_ID',
  IMGBB_API_KEY: 'GIFT_MEMORY_IMGBB_API_KEY',
  ADMIN_PASSWORD: 'GIFT_MEMORY_ADMIN_PASSWORD',
  LOGIN_SALT: 'GIFT_MEMORY_LOGIN_SALT',
  YOUTUBE_API_KEY: 'GIFT_MEMORY_YOUTUBE_API_KEY',
  SITE_URL: 'GIFT_MEMORY_SITE_URL'
});

/* -------------------------------------------------------------------------- */
/* 02. FIRST-TIME SETUP                                                       */
/* -------------------------------------------------------------------------- */

/**
 * Run this function manually once from the Apps Script editor.
 * Creates the database tabs and their headers automatically. Image files are
 * hosted by ImgBB; upload chunks are held temporarily in Apps Script CacheService.
 */
function setupSystem() {
  const props = PropertiesService.getScriptProperties();
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  if (!spreadsheet) {
    const savedId = props.getProperty(PROP.SPREADSHEET_ID);
    if (savedId) {
      spreadsheet = SpreadsheetApp.openById(savedId);
    } else {
      spreadsheet = SpreadsheetApp.create('Heart Mate — Database');
    }
  }

  // Reuse the default empty tab as Gifts so a new spreadsheet has no stray Sheet1.
  const defaultSheet = spreadsheet.getSheetByName('Sheet1');
  if (defaultSheet && spreadsheet.getSheets().length === 1 && defaultSheet.getLastRow() === 0) {
    defaultSheet.setName(APP.SHEETS.GIFTS);
  }

  props.setProperty(PROP.SPREADSHEET_ID, spreadsheet.getId());

  Object.keys(APP.HEADERS).forEach(function (sheetName) {
    ensureSheet_(spreadsheet, sheetName, APP.HEADERS[sheetName]);
  });


  if (!props.getProperty(PROP.ADMIN_PASSWORD)) {
    props.setProperty(PROP.ADMIN_PASSWORD, APP.DEFAULT_ADMIN_PASSWORD);
  }
  if (!props.getProperty(PROP.LOGIN_SALT)) {
    props.setProperty(PROP.LOGIN_SALT, Utilities.getUuid());
  }

  // Keep the site URL in both Settings and Script Properties once configured.
  const storedSiteUrl = props.getProperty(PROP.SITE_URL);
  if (storedSiteUrl) {
    upsertSetting_('site_url', storedSiteUrl);
  }

  const result = {
    ok: true,
    message: 'ตั้งค่าระบบและสร้างหัวตารางเรียบร้อยแล้ว',
    spreadsheet_name: spreadsheet.getName(),
    spreadsheet_url: spreadsheet.getUrl(),
    sheets: Object.keys(APP.HEADERS),
    image_host: 'ImgBB',
    imgbb_api_key_configured: Boolean(props.getProperty(PROP.IMGBB_API_KEY)),
    admin_password_initialized: true
  };

  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

/**
 * Save the YouTube Data API v3 key in Script Properties.
 * Run manually from the editor, for example:
 * saveYouTubeApiKey('AIza...');
 * Do not put the key in index.html, dashboard.html, GitHub, or a Sheet cell.
 */
function saveYouTubeApiKey(apiKey) {
  const key = String(apiKey || '').trim();
  if (key.length < 20) {
    throw new Error('กรุณาระบุ YouTube Data API Key ที่ถูกต้อง');
  }
  PropertiesService.getScriptProperties().setProperty(PROP.YOUTUBE_API_KEY, key);
  return 'บันทึก YouTube API Key ลง Script Properties แล้ว';
}

/**
 * Optional: set the GitHub Pages base URL used to generate full gift links.
 * Example: saveSiteUrl('https://username.github.io/repository-name');
 */
function saveSiteUrl(siteUrl) {
  const normalized = normalizeSiteUrl_(siteUrl);
  PropertiesService.getScriptProperties().setProperty(PROP.SITE_URL, normalized);
  upsertSetting_('site_url', normalized);
  return 'บันทึก URL เว็บไซต์แล้ว: ' + normalized;
}

/** Returns the database spreadsheet URL for setup/maintenance. */
function getDatabaseUrl() {
  const id = PropertiesService.getScriptProperties().getProperty(PROP.SPREADSHEET_ID);
  if (!id) throw new Error('ยังไม่ได้ตั้งค่าระบบ กรุณารัน setupSystem() ก่อน');
  return SpreadsheetApp.openById(id).getUrl();
}

function ensureSheet_(spreadsheet, sheetName, requiredHeaders) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) sheet = spreadsheet.insertSheet(sheetName);

  const lastColumn = Math.max(sheet.getLastColumn(), 1);
  const existing = sheet.getLastRow() > 0
    ? sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0]
    : [];
  const existingHeaders = existing.map(function (value) { return String(value).trim(); });
  const knownHeaders = existingHeaders.filter(function (value) { return value !== ''; });
  const lastHeaderColumn = existingHeaders.reduce(function (last, value, index) {
    return value ? index + 1 : last;
  }, 0);

  if (knownHeaders.length === 0) {
    sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
  } else {
    const missing = requiredHeaders.filter(function (header) {
      return knownHeaders.indexOf(header) === -1;
    });
    if (missing.length) {
      sheet.getRange(1, lastHeaderColumn + 1, 1, missing.length).setValues([missing]);
    }
  }

  const finalColumns = sheet.getLastColumn();
  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 30);
  sheet.getRange(1, 1, 1, finalColumns)
    .setBackground('#FCE4EC')
    .setFontColor('#6D4355')
    .setFontWeight('bold')
    .setWrap(true);
  if (!sheet.getFilter() && sheet.getLastRow() > 1) {
    sheet.getRange(1, 1, sheet.getLastRow(), finalColumns).createFilter();
  }
  sheet.autoResizeColumns(1, Math.min(finalColumns, 12));
  return sheet;
}

/**
 * Store the ImgBB API key in Script Properties. Run manually in Apps Script:
 * saveImgBBApiKey('YOUR_IMGBB_API_KEY');
 * Never put this key in index.html, dashboard.html, GitHub, or a Sheet cell.
 */
function saveImgBBApiKey(apiKey) {
  const key = String(apiKey || '').trim();
  if (key.length < 10 || /\s/.test(key)) {
    throw new Error('กรุณาระบุ ImgBB API Key ที่ถูกต้อง');
  }
  PropertiesService.getScriptProperties().setProperty(PROP.IMGBB_API_KEY, key);
  return 'บันทึก ImgBB API Key ลง Script Properties แล้ว';
}

/** Run once manually to authorize UrlFetchApp for ImgBB uploads. */
function authorizeImgBBAccess() {
  const response = UrlFetchApp.fetch('https://api.imgbb.com/', { muteHttpExceptions: true });
  return {
    ok: true,
    message: 'อนุญาตการเชื่อมต่อภายนอกแล้ว',
    http_status: response.getResponseCode()
  };
}


/* -------------------------------------------------------------------------- */
/* 03. WEB APP ENTRY POINT AND JSONP RESPONSE                                 */
/* -------------------------------------------------------------------------- */

/** Main browser API entry point. All frontend requests use GET + JSONP. */
function doGet(event) {
  const params = event && event.parameter ? event.parameter : {};
  const action = String(params.action || 'ping').trim();
  let callback = '';

  try {
    callback = normalizeCallback_(params.callback || '');
    const result = routeAction_(action, params);
    return makeResponse_(result, callback);
  } catch (error) {
    console.error(error && error.stack ? error.stack : error);
    const message = error && error.message ? error.message : 'เกิดข้อผิดพลาดภายในระบบ';
    return makeResponse_({ ok: false, error: message }, callback);
  }
}

function routeAction_(action, params) {
  switch (action) {
    case 'ping':
      return { ok: true, app: APP.NAME, version: APP.VERSION, timestamp: nowIso_() };

    case 'getPublicConfig':
      return getPublicConfig_();

    case 'getPublicNonce':
      return getPublicNonce_();

    case 'createPublicGift':
      return createPublicGift_(params);

    case 'getLoginChallenge':
      return getLoginChallenge_();

    case 'login':
      return login_(params);

    case 'logout':
      return logout_(params.token);

    case 'listGifts':
      return listGifts_(params);

    case 'getGift':
      return getGift_(params);

    case 'createGift':
      return createGift_(params);

    case 'updateGift':
      return updateGift_(params);

    case 'deleteGift':
      return deleteGift_(params);

    case 'searchYouTube':
      return searchYouTube_(params);

    case 'saveSettings':
      return saveSettings_(params);

    case 'startImageUpload':
      return startImageUpload_(params);

    case 'uploadImageChunk':
      return uploadImageChunk_(params);

    case 'finishImageUpload':
      return finishImageUpload_(params);

    case 'deletePhoto':
      return deletePhoto_(params);

    default:
      throw new Error('ไม่รู้จัก API action: ' + action);
  }
}

function normalizeCallback_(callback) {
  const value = String(callback || '').trim();
  // Allow a simple dotted callback path, e.g. GiftApi.callbacks.cb123.
  if (!value) return '';
  const valid = /^[A-Za-z_$][A-Za-z0-9_$]*(\.[A-Za-z_$][A-Za-z0-9_$]*){0,3}$/;
  if (!valid.test(value)) throw new Error('ชื่อ JSONP callback ไม่ถูกต้อง');
  return value;
}

function makeResponse_(data, callback) {
  let body = JSON.stringify(data);
  body = body.replace(/\u2028/g, '\\u2028').replace(/\u2029/g, '\\u2029');

  if (callback) {
    return ContentService.createTextOutput('/**/' + callback + '(' + body + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}

function parsePayload_(params) {
  let raw = '';
  try {
    if (params.payload_gzip64) {
      // Preferred transport for the public creator: gzip UTF-8 JSON, then
      // Base64URL it so Thai characters do not expand inside the query string.
      let encoded = String(params.payload_gzip64).replace(/-/g, '+').replace(/_/g, '/');
      while (encoded.length % 4) encoded += '=';
      const compressedBytes = Utilities.base64Decode(encoded);
      raw = Utilities.ungzip(Utilities.newBlob(compressedBytes, 'application/gzip', 'payload.json.gz'))
        .getDataAsString('UTF-8');
    } else if (params.payload64) {
      // Uncompressed Base64URL fallback for older browsers.
      let encoded = String(params.payload64).replace(/-/g, '+').replace(/_/g, '/');
      while (encoded.length % 4) encoded += '=';
      const bytes = Utilities.base64Decode(encoded);
      raw = Utilities.newBlob(bytes).getDataAsString('UTF-8');
    } else if (params.payload) {
      raw = String(params.payload);
    } else {
      return {};
    }
    const parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
      throw new Error('payload ต้องเป็น JSON object');
    }
    return parsed;
  } catch (error) {
    throw new Error('ข้อมูล payload ไม่ถูกต้อง: ' + error.message);
  }
}

/* -------------------------------------------------------------------------- */
/* 04. AUTHENTICATION AND SESSION MANAGEMENT                                  */
/* -------------------------------------------------------------------------- */

function getLoginChallenge_() {
  const cache = CacheService.getScriptCache();
  const challengeId = Utilities.getUuid().replace(/-/g, '');
  const nonce = Utilities.getUuid() + Utilities.getUuid();
  const salt = PropertiesService.getScriptProperties().getProperty(PROP.LOGIN_SALT);
  if (!salt) throw new Error('ยังไม่ได้ตั้งค่าระบบ กรุณารัน setupSystem()');

  cache.put('LOGIN_CHALLENGE_' + challengeId, nonce, APP.LOGIN_CHALLENGE_TTL_SECONDS);
  return {
    ok: true,
    challenge_id: challengeId,
    nonce: nonce,
    salt: salt,
    algorithm: 'SHA-256(password + ":" + salt + ":" + nonce)',
    expires_in: APP.LOGIN_CHALLENGE_TTL_SECONDS
  };
}

/**
 * The frontend calculates SHA-256(password + ':' + salt + ':' + nonce) using
 * Web Crypto and sends only that proof. Each challenge can be used once.
 */
function login_(params) {
  const cache = CacheService.getScriptCache();
  const failuresKey = 'LOGIN_FAILURE_COUNT';
  const currentFailures = Number(cache.get(failuresKey) || 0);
  if (currentFailures >= APP.MAX_LOGIN_FAILURES) {
    throw new Error('พยายามเข้าสู่ระบบมากเกินไป กรุณารอประมาณ 15 นาที');
  }

  const challengeId = String(params.challenge_id || '').trim();
  const suppliedProof = String(params.proof || '').trim().toLowerCase();
  if (!challengeId || !/^[a-f0-9]{32}$/.test(challengeId)) {
    throw new Error('ข้อมูลยืนยันการเข้าสู่ระบบไม่ถูกต้อง');
  }

  const cacheKey = 'LOGIN_CHALLENGE_' + challengeId;
  const nonce = cache.get(cacheKey);
  cache.remove(cacheKey); // One-time challenge: prevent replay.
  if (!nonce) return loginFailed_('คำขอเข้าสู่ระบบหมดอายุ กรุณาลองใหม่');

  const props = PropertiesService.getScriptProperties();
  const password = props.getProperty(PROP.ADMIN_PASSWORD);
  const salt = props.getProperty(PROP.LOGIN_SALT);
  const expectedProof = sha256Hex_(password + ':' + salt + ':' + nonce);

  if (!constantTimeEquals_(expectedProof, suppliedProof)) {
    return loginFailed_('รหัสเข้าใช้งานไม่ถูกต้อง');
  }

  cache.remove(failuresKey);
  const token = Utilities.getUuid() + Utilities.getUuid();
  const sessionKey = 'ADMIN_SESSION_' + sha256Hex_(token);
  cache.put(sessionKey, JSON.stringify({ created_at: nowIso_(), role: 'admin' }), APP.SESSION_TTL_SECONDS);

  return {
    ok: true,
    token: token,
    token_type: 'Bearer',
    expires_in: APP.SESSION_TTL_SECONDS,
    message: 'เข้าสู่ระบบสำเร็จ'
  };
}

function loginFailed_(message) {
  const cache = CacheService.getScriptCache();
  const failuresKey = 'LOGIN_FAILURE_COUNT';
  const failures = Number(cache.get(failuresKey) || 0) + 1;
  cache.put(failuresKey, String(failures), APP.LOGIN_FAIL_TTL_SECONDS);
  throw new Error(message);
}

function requireAdmin_(token) {
  const value = String(token || '').trim();
  if (value.length < 20) throw new Error('กรุณาเข้าสู่ระบบก่อนใช้งาน');

  const key = 'ADMIN_SESSION_' + sha256Hex_(value);
  const session = CacheService.getScriptCache().get(key);
  if (!session) throw new Error('เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่');
  return JSON.parse(session);
}

function logout_(token) {
  const value = String(token || '').trim();
  if (value) CacheService.getScriptCache().remove('ADMIN_SESSION_' + sha256Hex_(value));
  return { ok: true, message: 'ออกจากระบบแล้ว' };
}


/* -------------------------------------------------------------------------- */
/* 05. DATABASE HELPERS                                                       */
/* -------------------------------------------------------------------------- */

function getSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty(PROP.SPREADSHEET_ID);
  if (!id) throw new Error('ยังไม่ได้ตั้งค่าฐานข้อมูล กรุณารัน setupSystem()');
  return SpreadsheetApp.openById(id);
}

function getSheet_(sheetName) {
  const sheet = getSpreadsheet_().getSheetByName(sheetName);
  if (!sheet) throw new Error('ไม่พบชีต ' + sheetName + ' กรุณารัน setupSystem()');
  return sheet;
}

function getHeaders_(sheet) {
  if (sheet.getLastColumn() < 1) throw new Error('ชีต ' + sheet.getName() + ' ยังไม่มีหัวตาราง');
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0]
    .map(function (value) { return String(value).trim(); });
}

function appendObject_(sheetName, record) {
  const sheet = getSheet_(sheetName);
  const headers = getHeaders_(sheet);
  const row = headers.map(function (header) {
    return Object.prototype.hasOwnProperty.call(record, header) ? record[header] : '';
  });
  sheet.appendRow(row);
  return sheet.getLastRow();
}

function readObjects_(sheetName) {
  const sheet = getSheet_(sheetName);
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || lastColumn < 1) return [];

  const headers = getHeaders_(sheet);
  const values = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
  return values.map(function (row) {
    const item = {};
    headers.forEach(function (header, index) {
      item[header] = row[index] instanceof Date ? row[index].toISOString() : row[index];
    });
    return item;
  });
}

function findRowByValue_(sheetName, columnName, value) {
  const sheet = getSheet_(sheetName);
  const headers = getHeaders_(sheet);
  const columnIndex = headers.indexOf(columnName);
  if (columnIndex === -1 || sheet.getLastRow() < 2) return -1;

  const values = sheet.getRange(2, columnIndex + 1, sheet.getLastRow() - 1, 1).getDisplayValues();
  for (let index = 0; index < values.length; index += 1) {
    if (String(values[index][0]) === String(value)) return index + 2;
  }
  return -1;
}

function getRecordById_(sheetName, idColumn, idValue) {
  const rowNumber = findRowByValue_(sheetName, idColumn, idValue);
  if (rowNumber === -1) return null;
  const sheet = getSheet_(sheetName);
  const headers = getHeaders_(sheet);
  const values = sheet.getRange(rowNumber, 1, 1, headers.length).getValues()[0];
  const record = {};
  headers.forEach(function (header, index) {
    record[header] = values[index] instanceof Date ? values[index].toISOString() : values[index];
  });
  return { row_number: rowNumber, record: record };
}

function updateRecordById_(sheetName, idColumn, idValue, updates) {
  const found = getRecordById_(sheetName, idColumn, idValue);
  if (!found) throw new Error('ไม่พบข้อมูลที่ต้องการแก้ไข');

  const sheet = getSheet_(sheetName);
  const headers = getHeaders_(sheet);
  Object.keys(updates).forEach(function (key) {
    const index = headers.indexOf(key);
    if (index !== -1) sheet.getRange(found.row_number, index + 1).setValue(updates[key]);
  });
  return found.row_number;
}

function deleteRowsByValue_(sheetName, columnName, value) {
  const sheet = getSheet_(sheetName);
  const rowNumbers = [];
  const columnIndex = getHeaders_(sheet).indexOf(columnName);
  if (columnIndex === -1 || sheet.getLastRow() < 2) return;
  const values = sheet.getRange(2, columnIndex + 1, sheet.getLastRow() - 1, 1).getDisplayValues();
  values.forEach(function (row, index) {
    if (String(row[0]) === String(value)) rowNumbers.push(index + 2);
  });
  rowNumbers.reverse().forEach(function (rowNumber) { sheet.deleteRow(rowNumber); });
}

function withScriptLock_(callback) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    return callback();
  } finally {
    lock.releaseLock();
  }
}

/* -------------------------------------------------------------------------- */
/* 06. GIFT CRUD                                                              */
/* -------------------------------------------------------------------------- */

function listGifts_(params) {
  requireAdmin_(params.token);
  const typeFilter = String(params.type || '').trim();
  const statusFilter = String(params.status || '').trim();
  const keyword = String(params.q || '').trim().toLowerCase();

  let gifts = readObjects_(APP.SHEETS.GIFTS).map(giftFromRow_);
  if (typeFilter) gifts = gifts.filter(function (gift) { return gift.gift_type === typeFilter; });
  if (statusFilter) gifts = gifts.filter(function (gift) { return gift.status === statusFilter; });
  if (keyword) {
    gifts = gifts.filter(function (gift) {
      const text = [gift.title, gift.recipient_name, gift.sender_name, gift.gift_id]
        .join(' ').toLowerCase();
      return text.indexOf(keyword) !== -1;
    });
  }
  gifts.sort(function (a, b) { return String(b.created_at).localeCompare(String(a.created_at)); });
  return { ok: true, gifts: gifts, count: gifts.length };
}

function publicPhoto_(photo) {
  return {
    photo_id: String(photo.photo_id || ''),
    gift_id: String(photo.gift_id || ''),
    image_url: safeHttpUrl_(photo.image_url),
    thumbnail_url: safeHttpUrl_(photo.thumbnail_url || photo.image_url),
    caption: String(photo.caption || ''),
    photo_type: String(photo.photo_type || ''),
    created_at: String(photo.created_at || '')
  };
}

function getGift_(params) {
  const giftId = sanitizeText_(params.id || params.gift_id, 80);
  if (!giftId) throw new Error('กรุณาระบุรหัสของขวัญ');

  const found = getRecordById_(APP.SHEETS.GIFTS, 'gift_id', giftId);
  if (!found) throw new Error('ไม่พบของขวัญนี้');
  const gift = giftFromRow_(found.record);

  if (gift.status !== 'published') {
    requireAdmin_(params.token);
  }

  const relatedPhotos = readObjects_(APP.SHEETS.PHOTOS).filter(function (photo) {
    return String(photo.gift_id) === giftId;
  });
  const relatedMemories = readObjects_(APP.SHEETS.MEMORIES).filter(function (memory) {
    return String(memory.gift_id) === giftId;
  });
  const card = getRecordById_(APP.SHEETS.CARDS, 'gift_id', giftId);
  const song = getRecordById_(APP.SHEETS.SONGS, 'gift_id', giftId);

  gift.photos = relatedPhotos.map(publicPhoto_);
  gift.memories = relatedMemories;
  gift.card = card ? card.record : null;
  gift.song = song ? song.record : null;

  // Public view counter. Failures to count a view should not block the gift.
  try {
    withScriptLock_(function () {
      const latest = getRecordById_(APP.SHEETS.GIFTS, 'gift_id', giftId);
      if (latest) {
        const nextCount = Number(latest.record.view_count || 0) + 1;
        updateRecordById_(APP.SHEETS.GIFTS, 'gift_id', giftId, {
          view_count: nextCount,
          last_viewed_at: nowIso_()
        });
        gift.view_count = nextCount;
      }
    });
  } catch (error) {
    console.warn('Unable to increment view counter: ' + error.message);
  }

  return { ok: true, gift: gift };
}

/**
 * Issues a short-lived, one-time nonce for the public gift creator.
 * This is an anti-CSRF / replay barrier, not a CAPTCHA.
 */
function getPublicNonce_() {
  const nonce = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').substring(0, 16);
  CacheService.getScriptCache().put('PUBLIC_CREATE_NONCE_' + nonce, '1', 600);
  return { ok: true, nonce: nonce, expires_in: 600, min_form_time_ms: 1200 };
}

/**
 * Public creation endpoint. Visitors can create and publish their own gift
 * without an admin account. Public gifts are link-accessible by design.
 * Best-effort abuse controls: one-time nonce, honeypot, elapsed-time check,
 * per-browser throttle, and a global daily ceiling.
 */
function createPublicGift_(params) {
  const nonce = String(params.nonce || '').trim();
  if (!/^[a-f0-9]{48}$/.test(nonce)) throw new Error('คำขอสร้างของขวัญไม่ถูกต้อง กรุณาลองใหม่');
  const nonceKey = 'PUBLIC_CREATE_NONCE_' + nonce;
  const nonceCache = CacheService.getScriptCache();
  if (!nonceCache.get(nonceKey)) throw new Error('คำขอสร้างของขวัญหมดอายุ กรุณาลองใหม่');
  nonceCache.remove(nonceKey);

  // Hidden fields should remain empty. They discourage simple automated spam.
  if (String(params.website || '').trim() || String(params.company || '').trim()) {
    throw new Error('ไม่สามารถส่งคำขอนี้ได้');
  }
  const startedAt = Number(params.form_started_at || 0);
  if (startedAt && (Date.now() - startedAt < 1200 || Date.now() - startedAt > 86400000)) {
    throw new Error('กรุณาตรวจสอบข้อมูลก่อนส่งอีกครั้ง');
  }

  const clientId = normalizePublicClientId_(params.client_id);
  incrementCacheLimit_('PUBLIC_CREATE_' + clientId, 4, 3600, 'คุณสร้างของขวัญบ่อยเกินไป กรุณารอสักครู่แล้วลองใหม่');
  incrementDailyLimit_('PUBLIC_GIFT_CREATE', 300, 'ขณะนี้มีการสร้างของขวัญจำนวนมาก กรุณาลองใหม่ภายหลัง');

  const input = parsePayload_(params);
  input.status = 'published'; // Publicly-created gifts are published immediately.
  const payload = normalizeGiftPayload_(input);

  return withScriptLock_(function () {
    const giftId = createUniqueGiftId_();
    const timestamp = nowIso_();
    const record = {
      gift_id: giftId,
      gift_type: payload.gift_type,
      title: payload.title,
      recipient_name: payload.recipient_name,
      sender_name: payload.sender_name,
      status: 'published',
      theme: payload.theme,
      message: payload.message,
      payload_json: JSON.stringify(payload.content),
      created_at: timestamp,
      updated_at: timestamp,
      view_count: 0,
      last_viewed_at: ''
    };
    appendObject_(APP.SHEETS.GIFTS, record);
    syncGiftDetailSheets_(giftId, payload, timestamp);

    // This one-time-returned secret lets the creator attach their own photos
    // to any gift type without exposing an editable public endpoint to link viewers.
    const ownerUploadToken = Utilities.getUuid() + Utilities.getUuid();
    PropertiesService.getScriptProperties().setProperty(
      'PUBLIC_OWNER_UPLOAD_' + giftId,
      JSON.stringify({ hash: sha256Hex_(ownerUploadToken), created_at_ms: Date.now() })
    );

    const publicUrl = buildGiftUrl_(giftId);
    return {
      ok: true,
      gift_id: giftId,
      owner_upload_token: ownerUploadToken,
      relative_url: '?id=' + encodeURIComponent(giftId),
      public_url: publicUrl,
      gift: Object.assign({}, record, { content: payload.content, public_url: publicUrl }),
      message: 'สร้างของขวัญและเผยแพร่เรียบร้อยแล้ว'
    };
  });
}

function normalizePublicClientId_(value) {
  const id = String(value || '').trim();
  if (!/^[A-Za-z0-9_-]{16,64}$/.test(id)) throw new Error('ไม่พบรหัสเซสชันของเบราว์เซอร์ กรุณารีเฟรชหน้าเว็บ');
  return id;
}

function incrementCacheLimit_(key, limit, ttlSeconds, errorMessage) {
  return withScriptLock_(function () {
    const cache = CacheService.getScriptCache();
    const count = Number(cache.get(key) || 0);
    if (count >= limit) throw new Error(errorMessage);
    cache.put(key, String(count + 1), ttlSeconds);
    return count + 1;
  });
}

function incrementDailyLimit_(prefix, limit, errorMessage) {
  return withScriptLock_(function () {
    const dateKey = Utilities.formatDate(new Date(), 'Etc/UTC', 'yyyyMMdd');
    const key = prefix + '_' + dateKey;
    const props = PropertiesService.getScriptProperties();
    const count = Number(props.getProperty(key) || 0);
    if (count >= limit) throw new Error(errorMessage);
    props.setProperty(key, String(count + 1));
    return count + 1;
  });
}

function createGift_(params) {
  requireAdmin_(params.token);
  const payload = normalizeGiftPayload_(parsePayload_(params));
  return withScriptLock_(function () {
    const giftId = createUniqueGiftId_();
    const timestamp = nowIso_();
    const record = {
      gift_id: giftId,
      gift_type: payload.gift_type,
      title: payload.title,
      recipient_name: payload.recipient_name,
      sender_name: payload.sender_name,
      status: payload.status,
      theme: payload.theme,
      message: payload.message,
      payload_json: JSON.stringify(payload.content),
      created_at: timestamp,
      updated_at: timestamp,
      view_count: 0,
      last_viewed_at: ''
    };
    appendObject_(APP.SHEETS.GIFTS, record);
    syncGiftDetailSheets_(giftId, payload, timestamp);

    const publicUrl = buildGiftUrl_(giftId);
    return {
      ok: true,
      gift_id: giftId,
      relative_url: '?id=' + encodeURIComponent(giftId),
      public_url: publicUrl,
      gift: Object.assign({}, record, { content: payload.content, public_url: publicUrl }),
      message: 'สร้างของขวัญเรียบร้อยแล้ว'
    };
  });
}

function updateGift_(params) {
  requireAdmin_(params.token);
  const giftId = sanitizeText_(params.id || params.gift_id, 80);
  if (!giftId) throw new Error('กรุณาระบุรหัสของขวัญ');
  const patch = parsePayload_(params);

  return withScriptLock_(function () {
    const found = getRecordById_(APP.SHEETS.GIFTS, 'gift_id', giftId);
    if (!found) throw new Error('ไม่พบของขวัญที่ต้องการแก้ไข');
    const oldGift = giftFromRow_(found.record);
    const merged = Object.assign({}, oldGift, patch);
    merged.content = Object.assign({}, oldGift.content || {}, patch.content || {});
    const payload = normalizeGiftPayload_(merged);
    const timestamp = nowIso_();

    updateRecordById_(APP.SHEETS.GIFTS, 'gift_id', giftId, {
      gift_type: payload.gift_type,
      title: payload.title,
      recipient_name: payload.recipient_name,
      sender_name: payload.sender_name,
      status: payload.status,
      theme: payload.theme,
      message: payload.message,
      payload_json: JSON.stringify(payload.content),
      updated_at: timestamp
    });
    syncGiftDetailSheets_(giftId, payload, timestamp);

    return {
      ok: true,
      gift_id: giftId,
      public_url: buildGiftUrl_(giftId),
      message: 'บันทึกการแก้ไขเรียบร้อยแล้ว'
    };
  });
}

function deleteGift_(params) {
  requireAdmin_(params.token);
  const giftId = sanitizeText_(params.id || params.gift_id, 80);
  if (!giftId) throw new Error('กรุณาระบุรหัสของขวัญ');

  return withScriptLock_(function () {
    const found = getRecordById_(APP.SHEETS.GIFTS, 'gift_id', giftId);
    if (!found) throw new Error('ไม่พบของขวัญที่ต้องการลบ');

    [APP.SHEETS.CARDS, APP.SHEETS.SONGS, APP.SHEETS.PHOTOS, APP.SHEETS.MEMORIES]
      .forEach(function (sheetName) { deleteRowsByValue_(sheetName, 'gift_id', giftId); });
    deleteRowsByValue_(APP.SHEETS.GIFTS, 'gift_id', giftId);
    PropertiesService.getScriptProperties().deleteProperty('PUBLIC_OWNER_UPLOAD_' + giftId);
    return { ok: true, message: 'ลบข้อมูลของขวัญออกจาก Google Sheets แล้ว รูปภาพที่เคยฝากไว้บน ImgBB ยังคงอยู่' };
  });
}

function normalizeGiftPayload_(input) {
  const allowedTypes = ['flower', 'card', 'song', 'photobooth', 'memory'];
  const giftType = String(input.gift_type || input.type || '').trim().toLowerCase();
  if (allowedTypes.indexOf(giftType) === -1) {
    throw new Error('ประเภทของขวัญไม่ถูกต้อง ต้องเป็น flower, card, song, photobooth หรือ memory');
  }

  const status = String(input.status || 'published').toLowerCase();
  if (['draft', 'published', 'archived'].indexOf(status) === -1) {
    throw new Error('สถานะของขวัญไม่ถูกต้อง');
  }

  const content = input.content && typeof input.content === 'object' && !Array.isArray(input.content)
    ? input.content
    : (input.data && typeof input.data === 'object' && !Array.isArray(input.data) ? input.data : {});

  if (giftType === 'song') {
    const videoId = String(content.youtube_video_id || content.video_id || '').trim();
    if (!/^[A-Za-z0-9_-]{11}$/.test(videoId)) {
      throw new Error('กรุณาเลือกเพลง YouTube ที่มี Video ID ถูกต้องก่อนบันทึก');
    }
    content.youtube_video_id = videoId;
  }

  return {
    gift_type: giftType,
    title: sanitizeText_(input.title || defaultGiftTitle_(giftType), 160),
    recipient_name: sanitizeText_(input.recipient_name, 120),
    sender_name: sanitizeText_(input.sender_name, 120),
    status: status,
    theme: sanitizeText_(input.theme || 'pastel', 80),
    message: sanitizeText_(input.message, APP.MAX_TEXT_LENGTH),
    content: content
  };
}

function giftFromRow_(row) {
  let content = {};
  try { content = row.payload_json ? JSON.parse(String(row.payload_json)) : {}; }
  catch (error) { content = {}; }

  return {
    gift_id: String(row.gift_id || ''),
    gift_type: String(row.gift_type || ''),
    title: String(row.title || ''),
    recipient_name: String(row.recipient_name || ''),
    sender_name: String(row.sender_name || ''),
    status: String(row.status || 'draft'),
    theme: String(row.theme || 'pastel'),
    message: String(row.message || ''),
    content: content,
    created_at: String(row.created_at || ''),
    updated_at: String(row.updated_at || ''),
    view_count: Number(row.view_count || 0),
    last_viewed_at: String(row.last_viewed_at || '')
  };
}

function syncGiftDetailSheets_(giftId, payload, timestamp) {
  const content = payload.content || {};
  if (payload.gift_type !== 'card') deleteRowsByValue_(APP.SHEETS.CARDS, 'gift_id', giftId);
  if (payload.gift_type !== 'song') deleteRowsByValue_(APP.SHEETS.SONGS, 'gift_id', giftId);
  if (payload.gift_type !== 'memory') deleteRowsByValue_(APP.SHEETS.MEMORIES, 'gift_id', giftId);

  if (payload.gift_type === 'card') {
    upsertByGiftId_(APP.SHEETS.CARDS, giftId, {
      gift_id: giftId,
      design_id: sanitizeText_(content.design_id || content.design || 'classic', 80),
      greeting: sanitizeText_(content.greeting, 300),
      message: sanitizeText_(content.message || payload.message, APP.MAX_TEXT_LENGTH),
      closing: sanitizeText_(content.closing, 300),
      created_at: timestamp,
      updated_at: timestamp
    });
  }

  if (payload.gift_type === 'song') {
    const videoId = String(content.youtube_video_id || '');
    upsertByGiftId_(APP.SHEETS.SONGS, giftId, {
      gift_id: giftId,
      youtube_video_id: videoId,
      video_title: sanitizeText_(content.video_title || content.title, 300),
      channel_title: sanitizeText_(content.channel_title, 200),
      thumbnail_url: safeHttpUrl_(content.thumbnail_url),
      watch_url: 'https://www.youtube.com/watch?v=' + videoId,
      embed_url: 'https://www.youtube.com/embed/' + videoId,
      created_at: timestamp,
      updated_at: timestamp
    });
  }

  if (payload.gift_type === 'memory' && Array.isArray(content.entries)) {
    deleteRowsByValue_(APP.SHEETS.MEMORIES, 'gift_id', giftId);
    content.entries.slice(0, 100).forEach(function (entry) {
      appendObject_(APP.SHEETS.MEMORIES, {
        memory_id: createId_('mem'),
        gift_id: giftId,
        memory_title: sanitizeText_(entry.title, 200),
        memory_text: sanitizeText_(entry.text, 2000),
        memory_date: sanitizeText_(entry.date, 40),
        photo_url: safeHttpUrl_(entry.photo_url),
        created_at: timestamp
      });
    });
  }
}

function upsertByGiftId_(sheetName, giftId, record) {
  const existing = getRecordById_(sheetName, 'gift_id', giftId);
  if (existing) {
    updateRecordById_(sheetName, 'gift_id', giftId, record);
  } else {
    appendObject_(sheetName, record);
  }
}

function createUniqueGiftId_() {
  let id;
  do { id = 'gift-' + Utilities.getUuid().replace(/-/g, '').substring(0, 12).toLowerCase(); }
  while (getRecordById_(APP.SHEETS.GIFTS, 'gift_id', id));
  return id;
}

function defaultGiftTitle_(type) {
  const titles = {
    flower: 'ช่อดอกไม้ของขวัญ',
    card: 'การ์ดพิเศษ',
    song: 'เพลงที่อยากส่งให้',
    photobooth: 'โฟโต้บูธความทรงจำ',
    memory: 'กล่องความทรงจำ'
  };
  return titles[type] || 'ของขวัญพิเศษ';
}

function buildGiftUrl_(giftId) {
  const base = PropertiesService.getScriptProperties().getProperty(PROP.SITE_URL);
  if (!base) return '?id=' + encodeURIComponent(giftId);
  return base.replace(/\/+$/, '') + '/?id=' + encodeURIComponent(giftId);
}

/* -------------------------------------------------------------------------- */
/* 07. YOUTUBE SEARCH                                                         */
/* -------------------------------------------------------------------------- */

function searchYouTube_(params) {
  if (params.token) {
    requireAdmin_(params.token);
  } else {
    const publicClientId = normalizePublicClientId_(params.client_id);
    incrementCacheLimit_('PUBLIC_YT_CLIENT_' + publicClientId, 30, 3600, 'ค้นหาเพลงบ่อยเกินไป กรุณารอสักครู่');
  }
  const query = sanitizeText_(params.q || params.query, 100);
  if (query.length < 2) throw new Error('กรุณาพิมพ์ชื่อเพลงหรือศิลปินอย่างน้อย 2 ตัวอักษร');

  const maxResults = Math.max(1, Math.min(Number(params.max_results || APP.YOUTUBE_MAX_RESULTS), APP.YOUTUBE_MAX_RESULTS));
  const cache = CacheService.getScriptCache();
  const cacheKey = 'YT_SEARCH_' + sha256Hex_(query.toLowerCase() + '|' + maxResults);
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  if (!params.token) {
    incrementDailyLimit_('PUBLIC_YOUTUBE_SEARCH', 85, 'ระบบค้นหาเพลงใช้งานครบโควตาประจำวันแล้ว กรุณาลองใหม่ในวันถัดไป');
  }

  const apiKey = PropertiesService.getScriptProperties().getProperty(PROP.YOUTUBE_API_KEY);
  if (!apiKey) throw new Error('ยังไม่ได้ตั้งค่า YouTube API Key กรุณารัน saveYouTubeApiKey() ก่อน');

  const paramsMap = {
    part: 'snippet',
    q: query,
    type: 'video',
    videoEmbeddable: 'true',
    maxResults: String(maxResults),
    order: 'relevance',
    regionCode: 'TH',
    relevanceLanguage: 'th',
    safeSearch: 'moderate',
    key: apiKey
  };
  const queryString = Object.keys(paramsMap).map(function (key) {
    return encodeURIComponent(key) + '=' + encodeURIComponent(paramsMap[key]);
  }).join('&');
  const endpoint = 'https://www.googleapis.com/youtube/v3/search?' + queryString;
  const response = UrlFetchApp.fetch(endpoint, { muteHttpExceptions: true });
  const statusCode = response.getResponseCode();
  const responseText = response.getContentText();
  let data;
  try { data = JSON.parse(responseText); }
  catch (error) { throw new Error('YouTube ส่งข้อมูลกลับมาไม่ถูกต้อง'); }

  if (statusCode < 200 || statusCode >= 300 || data.error) {
    const apiMessage = data.error && data.error.message ? data.error.message : 'ไม่สามารถค้นหาเพลงได้';
    throw new Error('YouTube API: ' + apiMessage);
  }

  const results = (data.items || []).map(function (item) {
    const snippet = item.snippet || {};
    const thumbnails = snippet.thumbnails || {};
    const videoId = item.id && item.id.videoId ? item.id.videoId : '';
    return {
      video_id: videoId,
      title: snippet.title || '',
      channel_title: snippet.channelTitle || '',
      description: snippet.description || '',
      published_at: snippet.publishedAt || '',
      thumbnail_url: (thumbnails.high && thumbnails.high.url) ||
        (thumbnails.medium && thumbnails.medium.url) ||
        (thumbnails.default && thumbnails.default.url) || '',
      watch_url: videoId ? 'https://www.youtube.com/watch?v=' + videoId : '',
      embed_url: videoId ? 'https://www.youtube.com/embed/' + videoId : ''
    };
  }).filter(function (item) { return /^[A-Za-z0-9_-]{11}$/.test(item.video_id); });

  const result = { ok: true, query: query, results: results, count: results.length };
  cache.put(cacheKey, JSON.stringify(result), APP.YOUTUBE_CACHE_TTL_SECONDS);
  return result;
}

/* -------------------------------------------------------------------------- */
/* 08. IMGBB IMAGE UPLOADS                                                    */
/* -------------------------------------------------------------------------- */

/**
 * Starts one image upload. The client compresses to <= 1 MB, converts to
 * base64 and sends 4,000-character JSONP chunks. Chunks are held temporarily
 * in Script Cache, then finishImageUpload_ sends one multipart POST to ImgBB.
 * No image files are created in Google Drive.
 */
function startImageUpload_(params) {
  const payload = parsePayload_(params);
  const giftId = sanitizeText_(payload.gift_id, 80);
  if (!giftId) throw new Error('กรุณาระบุ gift_id สำหรับรูปภาพ');

  const imgbbKey = PropertiesService.getScriptProperties().getProperty(PROP.IMGBB_API_KEY);
  if (!imgbbKey) throw new Error('ยังไม่ได้ตั้งค่า ImgBB API Key กรุณาเพิ่ม GIFT_MEMORY_IMGBB_API_KEY ใน Script Properties');

  const isAdmin = Boolean(params.token);
  if (isAdmin) requireAdmin_(params.token);
  const giftFound = getRecordById_(APP.SHEETS.GIFTS, 'gift_id', giftId);
  if (!giftFound) throw new Error('ไม่พบของขวัญที่ต้องการแนบรูป');
  if (!isAdmin && String(giftFound.record.status) !== 'published') {
    throw new Error('ของขวัญนี้ยังไม่เปิดให้เพิ่มรูปภาพ');
  }

  let isOwnerUpload = false;
  if (!isAdmin && payload.owner_upload_token) {
    const ownerProperty = PropertiesService.getScriptProperties().getProperty('PUBLIC_OWNER_UPLOAD_' + giftId);
    if (ownerProperty) {
      try {
        const ownerRecord = JSON.parse(ownerProperty);
        const stillValid = Date.now() - Number(ownerRecord.created_at_ms || 0) <= 7 * 86400000;
        isOwnerUpload = stillValid && constantTimeEquals_(ownerRecord.hash, sha256Hex_(payload.owner_upload_token));
      } catch (error) {
        isOwnerUpload = false;
      }
    }
  }

  // Link viewers may add community snapshots only to Photo Booth gifts.
  // The original creator's secret allows attaching images to any gift type.
  if (!isAdmin && !isOwnerUpload && String(giftFound.record.gift_type) !== 'photobooth') {
    throw new Error('การเพิ่มรูปภาพประเภทนี้ต้องใช้ลิงก์สิทธิ์ของผู้สร้างของขวัญ');
  }
  if (!isAdmin && !isOwnerUpload) {
    const publicUploadKey = 'PUBLIC_UPLOADS_' + giftId;
    const currentUploads = Number(CacheService.getScriptCache().get(publicUploadKey) || 0);
    if (currentUploads >= 20) throw new Error('ของขวัญนี้มีการอัปโหลดรูปครบจำนวนชั่วคราวแล้ว');
    CacheService.getScriptCache().put(publicUploadKey, String(currentUploads + 1), 3600);
  }

  const mimeType = String(payload.mime_type || '').toLowerCase();
  if (['image/jpeg', 'image/png', 'image/webp'].indexOf(mimeType) === -1) {
    throw new Error('รองรับเฉพาะไฟล์ JPG, PNG และ WebP');
  }

  const expectedBytes = Number(payload.expected_bytes || 0);
  const totalChunks = Number(payload.total_chunks || 0);
  if (!Number.isFinite(expectedBytes) || expectedBytes < 1 || expectedBytes > APP.MAX_IMAGE_BYTES) {
    throw new Error('รูปภาพต้องมีขนาดไม่เกิน 1 MB หลังบีบอัด');
  }
  const expectedBase64Length = Math.ceil(expectedBytes / 3) * 4;
  const expectedChunkCount = Math.ceil(expectedBase64Length / APP.IMAGE_CHUNK_SIZE);
  if (!Number.isInteger(totalChunks) || totalChunks !== expectedChunkCount || totalChunks > APP.MAX_IMAGE_CHUNKS) {
    throw new Error('จำนวนชิ้นข้อมูลรูปภาพไม่ถูกต้อง กรุณาลองเลือกภาพใหม่');
  }

  const uploadId = Utilities.getUuid().replace(/-/g, '');
  const uploadToken = Utilities.getUuid() + Utilities.getUuid();
  const session = {
    upload_id: uploadId,
    upload_token_hash: sha256Hex_(uploadToken),
    cache_prefix: 'HM_IMG_' + uploadId + '_',
    gift_id: giftId,
    file_name: sanitizeFileName_(payload.file_name || 'heart-mate-photo.jpg'),
    mime_type: mimeType,
    expected_bytes: expectedBytes,
    expected_base64_length: expectedBase64Length,
    total_chunks: totalChunks,
    caption: sanitizeText_(payload.caption, 300),
    photo_type: sanitizeText_(payload.photo_type || 'photobooth', 40),
    created_at_ms: Date.now()
  };
  PropertiesService.getScriptProperties().setProperty('UPLOAD_' + uploadId, JSON.stringify(session));

  return {
    ok: true,
    upload_id: uploadId,
    upload_token: uploadToken,
    chunk_size: APP.IMAGE_CHUNK_SIZE,
    total_chunks: totalChunks,
    max_bytes: APP.MAX_IMAGE_BYTES,
    storage: 'imgbb',
    message: 'เริ่มอัปโหลดรูปภาพไปยัง ImgBB แล้ว'
  };
}

/** Stores one small Base64 part temporarily in Apps Script CacheService. */
function uploadImageChunk_(params) {
  const uploadId = String(params.upload_id || '').trim();
  const uploadToken = String(params.upload_token || '').trim();
  const index = Number(params.index);
  const chunk = String(params.chunk || '');
  const session = getUploadSession_(uploadId, uploadToken);

  if (!Number.isInteger(index) || index < 0 || index >= session.total_chunks) {
    throw new Error('หมายเลขชิ้นข้อมูลไม่ถูกต้อง');
  }
  if (!chunk || chunk.length > APP.IMAGE_CHUNK_SIZE || !/^[A-Za-z0-9+/=]+$/.test(chunk)) {
    throw new Error('ข้อมูลรูปภาพส่วนนี้ไม่ถูกต้อง');
  }
  if (index < session.total_chunks - 1 && chunk.length !== APP.IMAGE_CHUNK_SIZE) {
    throw new Error('ขนาดข้อมูลของชุดนี้ไม่ถูกต้อง กรุณาส่งชุดข้อมูลใหม่');
  }

  const cache = CacheService.getScriptCache();
  const key = session.cache_prefix + String(index).padStart(3, '0');
  cache.put(key, chunk, APP.IMAGE_UPLOAD_CACHE_TTL_SECONDS);
  return { ok: true, upload_id: uploadId, received_index: index };
}

/**
 * Reassembles cached chunks, validates the image, uploads it to ImgBB using
 * UrlFetchApp POST (multipart/form-data), then saves the returned URLs in
 * Google Sheets. A short completion cache makes retries idempotent.
 */
function finishImageUpload_(params) {
  const uploadId = String(params.upload_id || '').trim();
  const uploadToken = String(params.upload_token || '').trim();
  if (!/^[a-f0-9]{32}$/.test(uploadId)) throw new Error('รหัสอัปโหลดไม่ถูกต้อง');

  const cache = CacheService.getScriptCache();
  const doneKey = 'HM_IMG_DONE_' + uploadId;
  const completedRaw = cache.get(doneKey);
  if (completedRaw) {
    const completed = JSON.parse(completedRaw);
    if (!constantTimeEquals_(completed.upload_token_hash, sha256Hex_(uploadToken))) {
      throw new Error('ไม่มีสิทธิ์อัปโหลดรูปภาพนี้');
    }
    return completed.result;
  }

  const session = getUploadSession_(uploadId, uploadToken);
  const keys = [];
  for (let index = 0; index < session.total_chunks; index += 1) {
    keys.push(session.cache_prefix + String(index).padStart(3, '0'));
  }
  const chunksByKey = cache.getAll(keys);
  let base64 = '';
  for (let index = 0; index < keys.length; index += 1) {
    const chunk = chunksByKey[keys[index]];
    if (typeof chunk !== 'string' || !chunk) {
      throw new Error('ข้อมูลอัปโหลดหมดอายุหรือขาดชุดที่ ' + (index + 1) + ' จาก ' + keys.length + ' กรุณาเริ่มอัปโหลดรูปนี้ใหม่');
    }
    base64 += chunk;
  }

  if (base64.length !== session.expected_base64_length || !/^[A-Za-z0-9+/]+={0,2}$/.test(base64)) {
    throw new Error('ข้อมูลรูปภาพไม่ครบหรือไม่ถูกต้อง กรุณาเริ่มอัปโหลดใหม่');
  }
  const bytes = Utilities.base64Decode(base64);
  if (bytes.length !== session.expected_bytes || bytes.length > APP.MAX_IMAGE_BYTES) {
    throw new Error('ขนาดรูปภาพไม่ตรงกับข้อมูลเริ่มต้น กรุณาเริ่มอัปโหลดใหม่');
  }
  validateImageSignature_(bytes, session.mime_type);

  const apiKey = PropertiesService.getScriptProperties().getProperty(PROP.IMGBB_API_KEY);
  if (!apiKey) throw new Error('ไม่พบ ImgBB API Key กรุณาตั้งค่า GIFT_MEMORY_IMGBB_API_KEY');

  const blob = Utilities.newBlob(bytes, session.mime_type, session.file_name);
  const endpoint = 'https://api.imgbb.com/1/upload?key=' + encodeURIComponent(apiKey);
  let response;
  try {
    response = UrlFetchApp.fetch(endpoint, {
      method: 'post',
      payload: { image: blob, name: session.file_name },
      muteHttpExceptions: true,
      followRedirects: true
    });
  } catch (error) {
    throw new Error('เชื่อมต่อ ImgBB ไม่สำเร็จ: ' + (error && error.message ? error.message : 'กรุณาลองใหม่'));
  }

  const httpCode = response.getResponseCode();
  let imgbbResult;
  try {
    imgbbResult = JSON.parse(response.getContentText());
  } catch (error) {
    throw new Error('ImgBB ส่งคำตอบกลับมาไม่ถูกต้อง (HTTP ' + httpCode + ')');
  }
  if (httpCode < 200 || httpCode >= 300 || !imgbbResult || imgbbResult.success !== true || !imgbbResult.data) {
    const apiMessage = imgbbResult && imgbbResult.error && imgbbResult.error.message
      ? String(imgbbResult.error.message)
      : 'ไม่สามารถอัปโหลดรูปภาพได้';
    throw new Error('ImgBB อัปโหลดไม่สำเร็จ: ' + apiMessage + ' (HTTP ' + httpCode + ')');
  }

  const data = imgbbResult.data;
  const imageUrl = safeHttpUrl_(data.url || (data.image && data.image.url));
  const thumbnailUrl = safeHttpUrl_(data.display_url || (data.medium && data.medium.url) || (data.thumb && data.thumb.url) || imageUrl);
  const viewerUrl = safeHttpUrl_(data.url_viewer);
  if (!imageUrl) throw new Error('อัปโหลดสำเร็จแต่ ImgBB ไม่ส่ง URL รูปภาพกลับมา');

  const photoId = createId_('photo');
  const photoRecord = {
    photo_id: photoId,
    gift_id: session.gift_id,
    image_url: imageUrl,
    thumbnail_url: thumbnailUrl,
    caption: session.caption,
    photo_type: session.photo_type,
    created_at: nowIso_(),
    imgbb_image_id: sanitizeText_(data.id, 100),
    imgbb_url_viewer: viewerUrl
  };
  appendObject_(APP.SHEETS.PHOTOS, photoRecord);

  const result = {
    ok: true,
    photo: publicPhoto_(photoRecord),
    storage: 'imgbb',
    message: 'อัปโหลดรูปไปยัง ImgBB และบันทึกลิงก์ลง Google Sheets เรียบร้อยแล้ว'
  };
  cache.put(doneKey, JSON.stringify({ upload_token_hash: session.upload_token_hash, result: result }), 3600);
  cleanupUploadSession_(uploadId, session);
  return result;
}

function validateImageSignature_(bytes, mimeType) {
  const isJpeg = bytes.length >= 3 && (bytes[0] & 255) === 255 && (bytes[1] & 255) === 216 && (bytes[2] & 255) === 255;
  const isPng = bytes.length >= 8 && (bytes[0] & 255) === 137 && String.fromCharCode.apply(null, bytes.slice(1, 4)) === 'PNG';
  const isWebp = bytes.length >= 12 && String.fromCharCode.apply(null, bytes.slice(0, 4)) === 'RIFF' && String.fromCharCode.apply(null, bytes.slice(8, 12)) === 'WEBP';
  const valid = (mimeType === 'image/jpeg' && isJpeg) ||
    (mimeType === 'image/png' && isPng) ||
    (mimeType === 'image/webp' && isWebp);
  if (!valid) throw new Error('ชนิดไฟล์ไม่ตรงกับข้อมูลรูปภาพ กรุณาเลือก JPG, PNG หรือ WebP ใหม่');
}

function getUploadSession_(uploadId, uploadToken) {
  if (!/^[a-f0-9]{32}$/.test(uploadId)) throw new Error('รหัสอัปโหลดไม่ถูกต้อง');
  const raw = PropertiesService.getScriptProperties().getProperty('UPLOAD_' + uploadId);
  if (!raw) throw new Error('ไม่พบเซสชันอัปโหลด หรือเซสชันหมดอายุแล้ว');
  const session = JSON.parse(raw);
  if (!constantTimeEquals_(session.upload_token_hash, sha256Hex_(uploadToken))) {
    throw new Error('ไม่มีสิทธิ์อัปโหลดรูปภาพนี้');
  }
  if (Date.now() - Number(session.created_at_ms) > APP.IMAGE_UPLOAD_SESSION_TTL_MS) {
    cleanupUploadSession_(uploadId, session);
    throw new Error('เซสชันอัปโหลดหมดอายุ กรุณาเริ่มใหม่');
  }
  return session;
}

function cleanupUploadSession_(uploadId, session) {
  if (session && session.cache_prefix) {
    const keys = [];
    for (let index = 0; index < Number(session.total_chunks || 0); index += 1) {
      keys.push(session.cache_prefix + String(index).padStart(3, '0'));
    }
    if (keys.length) CacheService.getScriptCache().removeAll(keys);
  }
  PropertiesService.getScriptProperties().deleteProperty('UPLOAD_' + uploadId);
}

/** Run manually or from a time-driven trigger to clean abandoned uploads. */
function cleanupExpiredUploads() {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  let cleaned = 0;
  Object.keys(all).forEach(function (key) {
    if (key.indexOf('UPLOAD_') !== 0) return;
    try {
      const session = JSON.parse(all[key]);
      if (Date.now() - Number(session.created_at_ms) > APP.IMAGE_UPLOAD_SESSION_TTL_MS) {
        cleanupUploadSession_(key.substring('UPLOAD_'.length), session);
        cleaned += 1;
      }
    } catch (error) {
      props.deleteProperty(key);
    }
  });
  return { ok: true, cleaned_uploads: cleaned };
}

function deletePhoto_(params) {
  requireAdmin_(params.token);
  const photoId = sanitizeText_(params.photo_id || params.id, 100);
  if (!photoId) throw new Error('กรุณาระบุ photo_id');
  const found = getRecordById_(APP.SHEETS.PHOTOS, 'photo_id', photoId);
  if (!found) throw new Error('ไม่พบรูปภาพ');
  getSheet_(APP.SHEETS.PHOTOS).deleteRow(found.row_number);
  return {
    ok: true,
    remote_image_deleted: false,
    message: 'ลบรายการรูปภาพออกจาก Google Sheets แล้ว รูปต้นฉบับบน ImgBB ยังคงอยู่'
  };
}

/* -------------------------------------------------------------------------- */
/* 09. SETTINGS AND PUBLIC CONFIG                                             */
/* -------------------------------------------------------------------------- */

function getPublicConfig_() {
  const props = PropertiesService.getScriptProperties();
  return {
    ok: true,
    app_name: APP.NAME,
    version: APP.VERSION,
    public_creation_enabled: true,
    public_gift_types: ['flower', 'card', 'song', 'photobooth', 'memory'],
    site_url: props.getProperty(PROP.SITE_URL) || '',
    image_storage: 'imgbb',
    imgbb_api_key_configured: Boolean(props.getProperty(PROP.IMGBB_API_KEY)),
    max_image_bytes: APP.MAX_IMAGE_BYTES,
    image_chunk_size: APP.IMAGE_CHUNK_SIZE
  };
}

function saveSettings_(params) {
  requireAdmin_(params.token);
  const payload = parsePayload_(params);
  if (Object.prototype.hasOwnProperty.call(payload, 'site_url')) {
    const url = normalizeSiteUrl_(payload.site_url);
    PropertiesService.getScriptProperties().setProperty(PROP.SITE_URL, url);
    upsertSetting_('site_url', url);
  }
  return { ok: true, settings: { site_url: PropertiesService.getScriptProperties().getProperty(PROP.SITE_URL) || '' } };
}

function upsertSetting_(key, value) {
  const record = {
    setting_key: String(key),
    setting_value: String(value),
    updated_at: nowIso_()
  };
  const existing = getRecordById_(APP.SHEETS.SETTINGS, 'setting_key', key);
  if (existing) updateRecordById_(APP.SHEETS.SETTINGS, 'setting_key', key, record);
  else appendObject_(APP.SHEETS.SETTINGS, record);
}

function normalizeSiteUrl_(siteUrl) {
  const value = String(siteUrl || '').trim().replace(/\/+$/, '');
  if (!/^https:\/\/[A-Za-z0-9.-]+(?::\d+)?(?:\/[A-Za-z0-9._~!$&'()*+,;=:@%-]*)*$/.test(value)) {
    throw new Error('กรุณาระบุ URL เว็บไซต์แบบ https:// ที่ถูกต้อง');
  }
  return value;
}

/* -------------------------------------------------------------------------- */
/* 10. GENERAL UTILITIES                                                      */
/* -------------------------------------------------------------------------- */

function nowIso_() {
  return new Date().toISOString();
}

function sanitizeText_(value, maxLength) {
  if (value === null || value === undefined) return '';
  return String(value)
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '')
    .trim()
    .substring(0, maxLength || APP.MAX_TEXT_LENGTH);
}

function sanitizeFileName_(fileName) {
  const name = String(fileName || 'image.jpg').replace(/[\\/:*?"<>|\u0000-\u001F]/g, '_').trim();
  return (name || 'image.jpg').substring(0, 120);
}

function safeHttpUrl_(url) {
  const value = String(url || '').trim();
  return /^https:\/\//i.test(value) ? value.substring(0, 2000) : '';
}

function createId_(prefix) {
  return String(prefix || 'id') + '-' + Utilities.getUuid().replace(/-/g, '').substring(0, 16).toLowerCase();
}

function sha256Hex_(value) {
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(value),
    Utilities.Charset.UTF_8
  );
  return digest.map(function (byte) {
    const unsigned = (byte + 256) % 256;
    return ('0' + unsigned.toString(16)).slice(-2);
  }).join('');
}

function constantTimeEquals_(a, b) {
  const left = String(a || '');
  const right = String(b || '');
  if (left.length !== right.length) return false;
  let result = 0;
  for (let index = 0; index < left.length; index += 1) {
    result |= left.charCodeAt(index) ^ right.charCodeAt(index);
  }
  return result === 0;
}

/* -------------------------------------------------------------------------- */
/* 11. API CONTRACT FOR index.html / dashboard.html                           */
/* -------------------------------------------------------------------------- */
/**
 * Base URL:
 *   https://script.google.com/macros/s/DEPLOYMENT_ID/exec
 *
 * JSONP request format:
 *   <API_URL>?action=ping&callback=GiftApi.callbacks.cb1
 *
 * Every frontend request should use a unique callback name and remove its
 * script element after the callback runs. Use encodeURIComponent() for every
 * query-string value. Do not fetch() this endpoint from GitHub Pages.
 *
 * Public actions:
 *   ping
 *   getPublicConfig
 *   getPublicNonce
 *   createPublicGift        params: nonce, client_id, form_started_at, payload JSON
 *   searchYouTube           params: q, client_id (cached / public quota-limited)
 *   getGift                 params: id
 *   startImageUpload        payload: {gift_id, file_name, mime_type,
 *                             expected_bytes, total_chunks, caption, photo_type}
 *   uploadImageChunk        params: upload_id, upload_token, index, chunk
 *   finishImageUpload       params: upload_id, upload_token
 *
 * Authentication:
 *   1) getLoginChallenge -> {challenge_id, nonce, salt}
 *   2) Browser computes SHA-256(password + ':' + salt + ':' + nonce) as lowercase hex
 *      using crypto.subtle.digest('SHA-256', new TextEncoder().encode(text)).
 *   3) login params: challenge_id, proof -> {token, expires_in}
 *   4) Keep token in memory and send token=... with admin requests.
 *
 * Admin actions:
 *   logout                  params: token
 *   listGifts               params: token, optional type/status/q
 *   createGift              params: token, payload JSON string
 *   updateGift              params: token, id, payload JSON string
 *   deleteGift              params: token, id
 *   searchYouTube           params: token, q, optional max_results (admin bypasses public cap)
 *   saveSettings            params: token, payload JSON string
 *   startImageUpload        params: token + payload (admin may upload to drafts)
 *   deletePhoto             params: token, photo_id
 *
 * createGift payload example:
 * {
 *   "gift_type": "song",
 *   "title": "เพลงที่อยากส่งให้",
 *   "recipient_name": "คนพิเศษ",
 *   "sender_name": "คนส่ง",
 *   "status": "published",
 *   "theme": "pastel-pink",
 *   "message": "เพลงนี้ทำให้นึกถึงเธอ",
 *   "content": {
 *     "youtube_video_id": "VIDEO_ID_11_CHARS",
 *     "video_title": "ชื่อเพลง",
 *     "channel_title": "ชื่อช่อง",
 *     "thumbnail_url": "https://..."
 *   }
 * }
 *
 * Supported gift_type values:
 *   flower | card | song | photobooth | memory
 *
 * Note: Images are compressed to JPG/PNG/WebP and <= 1 MB in the browser.
 * The client sends base64 in 4,000-character chunks (IMAGE_CHUNK_SIZE).
 * finishImageUpload uploads the assembled image to ImgBB and saves the direct
 * image URL, thumbnail URL and ImgBB image id in the Photos sheet.
 */

/* -------------------------------------------------------------------------- */
/* 12. MAINTENANCE / OPTIONAL MANUAL TESTS                                    */
/* -------------------------------------------------------------------------- */

/** Run manually to verify the basic services after setup. */
function testSystemHealth() {
  const spreadsheet = getSpreadsheet_();
  const report = {
    ok: true,
    app: APP.NAME,
    version: APP.VERSION,
    spreadsheet_url: spreadsheet.getUrl(),
    sheets: Object.keys(APP.HEADERS).map(function (name) {
      const sheet = spreadsheet.getSheetByName(name);
      return { name: name, exists: Boolean(sheet), rows: sheet ? Math.max(0, sheet.getLastRow() - 1) : 0 };
    }),
    youtube_api_key_configured: Boolean(PropertiesService.getScriptProperties().getProperty(PROP.YOUTUBE_API_KEY)),
    imgbb_api_key_configured: Boolean(PropertiesService.getScriptProperties().getProperty(PROP.IMGBB_API_KEY)),
    image_storage: 'ImgBB',
    public_creation_enabled: true
  };
  Logger.log(JSON.stringify(report, null, 2));
  return report;
}
