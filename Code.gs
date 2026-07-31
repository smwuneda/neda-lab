/**
 * NEDA Lab News backend
 * Public: list posts
 * Admin: create/update/delete only after Firebase ID-token verification
 *
 * Sheet columns (row 1 header)
 * A id
 * B category
 * C title
 * D date
 * E description
 * F driveFileId
 * G createdAt
 * H galleryJson
 */
const SPREADSHEET_ID = '1cCqsib6TeVwZI91VdF_vXjAg2mDEWE9pJhwB96mXuC0';
const FOLDER_ID = '12SkDaPvMGQEZ2KqhOnSEG3mHBx1TAhMo';
const DATA_SHEET_NAME = 'Photos';
const FIREBASE_API_KEY = 'AIzaSyAuPJNrBUPqlbMXzAtJ6Xrd1DYL-1gm4r8';
const ALLOWED_EMAILS = ['smwu.neda.lab@gmail.com'];
const MAX_IMAGES = 30;

function doGet(e) {
  try {
    const action = String((e && e.parameter && e.parameter.action) || 'list');
    if (action !== 'list') return json_({ ok: false, error: 'Unsupported action.' });
    const category = String((e.parameter && e.parameter.category) || '').trim();
    const offset = Math.max(0, Number(e.parameter.offset || 0));
    const limit = Math.min(100, Math.max(1, Number(e.parameter.limit || 12)));
    return json_(listPosts_(offset, limit, category));
  } catch (error) {
    return json_({ ok: false, error: String((error && error.message) || error) });
  }
}

function doPost(e) {
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const action = String(data.action || '');
    if (['create', 'update', 'delete'].indexOf(action) !== -1) verifyAdmin_(data.idToken);
    if (action === 'create') return json_(createPost_(data));
    if (action === 'update') return json_(updatePost_(data));
    if (action === 'delete') return json_(deletePost_(data));
    return json_({ ok: false, error: 'Unsupported action.' });
  } catch (error) {
    return json_({ ok: false, error: String((error && error.message) || error) });
  }
}

function verifyAdmin_(idToken) {
  if (!idToken) throw new Error('Authentication token is missing.');
  const response = UrlFetchApp.fetch(
    'https://identitytoolkit.googleapis.com/v1/accounts:lookup?key=' + encodeURIComponent(FIREBASE_API_KEY),
    {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ idToken: String(idToken) }),
      muteHttpExceptions: true
    }
  );
  if (response.getResponseCode() !== 200) throw new Error('Login verification failed. Please sign in again.');
  const body = JSON.parse(response.getContentText() || '{}');
  const user = body.users && body.users[0];
  const email = String((user && user.email) || '').trim().toLowerCase();
  const verified = Boolean(user && user.emailVerified);
  const allow = ALLOWED_EMAILS.map(function (v) { return String(v).toLowerCase(); });
  if (!verified || allow.indexOf(email) === -1) throw new Error('This Google account is not authorized.');
  return email;
}

function getDataSheet_() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(DATA_SHEET_NAME);
  if (!sheet) throw new Error('"Photos" sheet not found.');
  ensureHeaders_(sheet);
  return sheet;
}

function ensureHeaders_(sheet) {
  const headers = ['id', 'category', 'title', 'date', 'description', 'driveFileId', 'createdAt', 'galleryJson'];
  const firstRow = sheet.getRange(1, 1, 1, headers.length).getDisplayValues()[0];
  let needs = false;
  for (var i = 0; i < headers.length; i++) {
    if (String(firstRow[i] || '').trim() !== headers[i]) {
      needs = true;
      break;
    }
  }
  if (needs) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
}

function getFolder_() {
  return DriveApp.getFolderById(FOLDER_ID);
}

function normalizeImages_(images) {
  if (!images) return [];
  if (!Array.isArray(images)) throw new Error('Images payload must be an array.');
  if (!images.length) return [];
  if (images.length > MAX_IMAGES) throw new Error('You can upload up to ' + MAX_IMAGES + ' images at once.');
  return images.map(function (item) {
    return {
      imageData: String(item.imageData || ''),
      mimeType: String(item.mimeType || 'image/jpeg'),
      fileName: String(item.fileName || ('neda-news-' + Date.now() + '.jpg'))
    };
  }).filter(function (item) { return item.imageData; });
}

function saveImages_(images) {
  const folder = getFolder_();
  return images.map(function (item) {
    const blob = Utilities.newBlob(Utilities.base64Decode(item.imageData), item.mimeType, item.fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { id: file.getId(), url: 'https://drive.google.com/uc?export=view&id=' + file.getId() };
  });
}

function parseGalleryFromRow_(row) {
  const galleryRaw = row[7];
  const primaryId = String(row[5] || '').trim();
  var gallery = [];
  if (galleryRaw) {
    try { gallery = JSON.parse(String(galleryRaw)); } catch (ignored) { gallery = []; }
  }
  if (!Array.isArray(gallery)) gallery = [];
  gallery = gallery.filter(function (item) { return item && String(item.id || '').trim(); }).map(function (item) {
    const id = String(item.id).trim();
    return { id: id, url: 'https://drive.google.com/uc?export=view&id=' + id };
  });
  if (!gallery.length && primaryId) gallery = [{ id: primaryId, url: 'https://drive.google.com/uc?export=view&id=' + primaryId }];
  return gallery;
}

function serializeGallery_(gallery) {
  return JSON.stringify((gallery || []).map(function (item) { return { id: String(item.id || '').trim() }; }));
}

function createPost_(data) {
  const title = String(data.title || '').trim();
  if (!title) throw new Error('A title is required.');
  const images = normalizeImages_(data.images);
  if (!images.length) throw new Error('At least one image is required.');
  const gallery = saveImages_(images);
  const id = Utilities.getUuid();
  const createdAt = new Date();
  const category = String(data.category || 'News').trim();
  const date = String(data.date || '').trim();
  const description = String(data.description || '').trim();
  getDataSheet_().appendRow([id, category, title, date, description, gallery[0].id, createdAt, serializeGallery_(gallery)]);
  return { ok: true, post: postToObject_([id, category, title, date, description, gallery[0].id, createdAt, serializeGallery_(gallery)]) };
}

function updatePost_(data) {
  const id = String(data.id || '').trim();
  if (!id) throw new Error('Missing post id.');
  const sheet = getDataSheet_();
  const values = sheet.getDataRange().getValues();
  let rowIndex = -1;
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === id) { rowIndex = i; break; }
  }
  if (rowIndex === -1) throw new Error('Post not found.');

  const row = values[rowIndex];
  const category = data.category !== undefined ? String(data.category).trim() : row[1];
  const title = data.title !== undefined ? String(data.title).trim() : row[2];
  const date = data.date !== undefined ? String(data.date).trim() : row[3];
  const description = data.description !== undefined ? String(data.description).trim() : row[4];
  let gallery = parseGalleryFromRow_(row);

  const newImages = normalizeImages_(data.images);
  if (newImages.length) {
    trashGallery_(gallery);
    gallery = saveImages_(newImages);
  }
  if (!gallery.length) throw new Error('At least one image is required.');

  const sheetRow = rowIndex + 1;
  sheet.getRange(sheetRow, 1, 1, 8).setValues([[id, category, title, date, description, gallery[0].id, row[6], serializeGallery_(gallery)]]);
  return { ok: true, post: postToObject_([id, category, title, date, description, gallery[0].id, row[6], serializeGallery_(gallery)]) };
}

function deletePost_(data) {
  const id = String(data.id || '').trim();
  if (!id) throw new Error('Missing post id.');
  const sheet = getDataSheet_();
  const values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === id) {
      trashGallery_(parseGalleryFromRow_(values[i]));
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }
  throw new Error('Post not found.');
}

function trashGallery_(gallery) {
  (gallery || []).forEach(function (item) {
    try { DriveApp.getFileById(item.id).setTrashed(true); } catch (ignored) {}
  });
}

function listPosts_(offset, limit, category) {
  const values = getDataSheet_().getDataRange().getValues();
  let rows = values.slice(1).filter(function (row) { return row[0]; });
  if (category) rows = rows.filter(function (row) { return String(row[1]).trim() === category; });
  rows.reverse();
  const total = rows.length;
  const selected = rows.slice(offset, offset + limit);
  return { ok: true, posts: selected.map(postToObject_), hasMore: offset + selected.length < total };
}

function postToObject_(row) {
  const gallery = parseGalleryFromRow_(row);
  const first = gallery[0] || null;
  return {
    id: row[0],
    category: row[1],
    title: row[2],
    date: row[3],
    description: row[4],
    imageFileId: first ? first.id : '',
    imageUrl: first ? first.url : '',
    imageUrls: gallery.map(function (item) { return item.url; }),
    gallery: gallery,
    photoCount: gallery.length,
    createdAt: row[6] instanceof Date ? row[6].toISOString() : String(row[6] || '')
  };
}

function json_(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function authorizeExternalRequest() {
  UrlFetchApp.fetch('https://www.googleapis.com/');
}
