/*
  Kingdom Bible for Jesus
  © 2026 Taiwo Adegoke. All rights reserved.
  This source is the property of Taiwo Adegoke. Unauthorized copying, reuse,
  modification, or redistribution is prohibited.
*/
/**
 * God's Favor Collection — backend (Google Apps Script)
 * Retail micro-app: product catalog (Native Attire + Body Oil), order/deposit tracking,
 * private supplier/cost, Drive photo upload, WhatsApp share, weekly Friday reminders.
 *
 * Built on the Cornerstone Studio engine (same as Vickie Jay Clothiers).
 * ALL critical fixes applied from line one:
 *   - Every API function returns a JSON string or plain id string (never a raw object).
 *   - No appendRow; setValues + SpreadsheetApp.flush().
 *   - Every value coerced; every list-builder guarded (one bad row can't blank a list).
 *   - Dates handled as ISO strings across the wire.
 *   - Photos returned as Chrome/iOS-reliable lh3 URLs.
 *
 * Glory to God alone.
 */

var PROP       = PropertiesService.getScriptProperties();
var SS_KEY     = 'GFC_SHEET_ID';
var FOLDER_KEY = 'GFC_FOLDER_ID';

// 'extraCost' is appended LAST so pre-existing 11-column product rows still map correctly.
var PRODUCTS_HEADERS = ['id','category','name','description','price','supplier','costPrice','photos','active','createdAt','updatedAt','extraCost','extraNote'];
var ORDERS_HEADERS   = ['id','customerName','phone','items','total','deposit','balance','orderDate','dueDate','status','remindersOn','notes','createdAt','updatedAt','photos','lineItems','depositDate'];
var SETTINGS_HEADERS = ['key','value'];
var LOG_HEADERS      = ['timestamp','orderId','customerName','message'];
var EXPENSES_HEADERS = ['id','date','description','amount','productId','category','createdAt'];
// Customer interest requests from the public storefront. Append-only at END.
var REQUESTS_HEADERS = ['id','createdAt','productId','productName','price','customerName','phone','offer','note','status','handledAt'];

var DEFAULT_SETTINGS = {
  businessName: "God's Favor Collection",
  ownerEmail: 'Godfavorcollection@gmail.com',
  remindersOn: 'yes',
  reminderDay: 'Friday',
  defaultTermDays: '21',
  pin: '',
  businessPhone: '',
  categories: '["Native Attire","Body Oil"]',
  tagline: 'Native attire & body oil, made with favor'
};

/* =========================================================================
 *  SETUP  (run once from the editor: Run > setup)
 * ========================================================================= */
function setup_(){
  var ss = getSS_();                 // creates the spreadsheet if missing
  sheet_(ss, 'Products', PRODUCTS_HEADERS);
  sheet_(ss, 'Orders',   ORDERS_HEADERS);
  var settingsSheet = sheet_(ss, 'Settings', SETTINGS_HEADERS);
  sheet_(ss, 'ReminderLog', LOG_HEADERS);
  sheet_(ss, 'Expenses', EXPENSES_HEADERS);
  sheet_(ss, 'Requests', REQUESTS_HEADERS);
  // keep the Products header row in sync if the schema grew (e.g. extraCost added later)
  var psh = ss.getSheetByName('Products');
  if (psh) psh.getRange(1, 1, 1, PRODUCTS_HEADERS.length).setValues([PRODUCTS_HEADERS]);
  var osh = ss.getSheetByName('Orders');
  if (osh) osh.getRange(1, 1, 1, ORDERS_HEADERS.length).setValues([ORDERS_HEADERS]);

  // seed default settings only if the Settings tab is empty (first run)
  if (settingsSheet.getLastRow() < 2){
    var rows = [];
    for (var k in DEFAULT_SETTINGS){ if (DEFAULT_SETTINGS.hasOwnProperty(k)) rows.push([k, DEFAULT_SETTINGS[k]]); }
    settingsSheet.getRange(2, 1, rows.length, 2).setValues(rows);
  }

  ensureFolder_();                   // creates the Drive photo folder if missing
  installReminderTrigger_();          // weekly Friday digest
  SpreadsheetApp.flush();

  return 'Setup complete.\nSpreadsheet: ' + ss.getUrl() +
         '\nPhoto folder: ' + DriveApp.getFolderById(PROP.getProperty(FOLDER_KEY)).getUrl();
}

/* =========================================================================
 *  WEB APP ENTRY
 * ========================================================================= */
function doGet(e){
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle("God's Favor Collection")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* =========================================================================
 *  STORAGE HELPERS
 * ========================================================================= */
function getSS_(){
  var id = PROP.getProperty(SS_KEY);
  if (id){
    try { return SpreadsheetApp.openById(id); } catch (err){ /* fall through, recreate */ }
  }
  var ss = SpreadsheetApp.create("God's Favor Collection — Data");
  PROP.setProperty(SS_KEY, ss.getId());
  return ss;
}

function sheet_(ss, name, headers){
  var sh = ss.getSheetByName(name);
  if (!sh){
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  } else if (sh.getLastRow() === 0){
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function ensureFolder_(){
  var id = PROP.getProperty(FOLDER_KEY);
  if (id){
    try { return DriveApp.getFolderById(id); } catch (err){ /* recreate */ }
  }
  var folder = DriveApp.createFolder("God's Favor Collection — Photos");
  PROP.setProperty(FOLDER_KEY, folder.getId());
  return folder;
}

function coerce_(v){
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return v.toISOString();
  if (typeof v === 'string' || typeof v === 'number' || typeof v === 'boolean') return v;
  return String(v);
}

function genId_(prefix){
  return prefix + '_' + Date.now().toString(36) + Math.floor(Math.random() * 1e4).toString(36);
}

// Read all data rows of a tab as {row:<sheetRowNumber>, obj:<record>}
function readAll_(name, headers){
  var out = [];
  try {
    var sh = sheet_(getSS_(), name, headers);
    var last = sh.getLastRow();
    if (last < 2) return out;
    var values = sh.getRange(2, 1, last - 1, headers.length).getValues();
    for (var i = 0; i < values.length; i++){
      try {
        var obj = {};
        for (var c = 0; c < headers.length; c++){ obj[headers[c]] = coerce_(values[i][c]); }
        if (String(obj.id) === '') continue;   // skip blank rows
        out.push({ row: i + 2, obj: obj });
      } catch (rowErr){ /* skip one malformed row, keep the rest */ }
    }
  } catch (err){ /* return whatever we have */ }
  return out;
}

function findRow_(name, headers, id){
  var all = readAll_(name, headers);
  for (var i = 0; i < all.length; i++){ if (String(all[i].obj.id) === String(id)) return all[i]; }
  return null;
}

/* =========================================================================
 *  SETTINGS
 * ========================================================================= */
function getSettingsObj_(){
  var obj = {};
  for (var k in DEFAULT_SETTINGS){ if (DEFAULT_SETTINGS.hasOwnProperty(k)) obj[k] = DEFAULT_SETTINGS[k]; }
  var rows = readAllSettings_();
  for (var i = 0; i < rows.length; i++){ obj[rows[i].key] = rows[i].value; }
  return obj;
}

function readAllSettings_(){
  var out = [];
  try {
    var sh = sheet_(getSS_(), 'Settings', SETTINGS_HEADERS);
    var last = sh.getLastRow();
    if (last < 2) return out;
    var vals = sh.getRange(2, 1, last - 1, 2).getValues();
    for (var i = 0; i < vals.length; i++){
      var key = String(coerce_(vals[i][0])).trim();
      if (key) out.push({ row: i + 2, key: key, value: coerce_(vals[i][1]) });
    }
  } catch (err){}
  return out;
}

/* =========================================================================
 *  OWNER AUTHENTICATION  (PIN)
 *  The PIN stored in Settings is the owner credential. The public storefront
 *  (apiShop) needs none; every private endpoint below requires it. Anonymous
 *  visitors — the shared link — can therefore never read cost/profit/orders,
 *  even by calling the API directly from the console.
 * ========================================================================= */
function ownerPin_(){ return String(getSettingsObj_().pin || '').trim(); }

// True only when a PIN is set AND the supplied one matches. Fails closed.
function requireOwner_(pin){
  var stored = ownerPin_();
  return stored !== '' && String(pin == null ? '' : pin).trim() === stored;
}

// Owner sign-in. First-ever login (no PIN set yet) provisions the PIN she types.
function apiOwnerLogin(pin){
  try {
    pin = String(pin == null ? '' : pin).trim();
    if (!pin) return JSON.stringify({ ok: false, err: 'Enter your PIN' });
    var stored = ownerPin_();
    if (stored === ''){
      if (pin.length < 4) return JSON.stringify({ ok: false, err: 'Choose a PIN of at least 4 digits' });
      setSettingValue_('pin', pin);
      SpreadsheetApp.flush();
      return JSON.stringify({ ok: true, provisioned: true });
    }
    if (pin === stored) return JSON.stringify({ ok: true });
    return JSON.stringify({ ok: false, err: 'Wrong PIN' });
  } catch (err){ return JSON.stringify({ ok: false, err: 'Please try again' }); }
}

function apiGetSettings(pin){
  if (!requireOwner_(pin)) return '{}';
  return JSON.stringify(getSettingsObj_());
}

function apiSaveSettings(payloadJson, pin){
  if (!requireOwner_(pin)) return 'error:auth';
  try {
    var payload = JSON.parse(payloadJson || '{}');
    var sh = sheet_(getSS_(), 'Settings', SETTINGS_HEADERS);
    var existing = readAllSettings_();
    var index = {};
    for (var i = 0; i < existing.length; i++){ index[existing[i].key] = existing[i].row; }
    for (var key in payload){
      if (!payload.hasOwnProperty(key)) continue;
      var val = coerce_(payload[key]);
      var writeRow = index[key] ? index[key] : sh.getLastRow() + 1;
      sh.getRange(writeRow, 1, 1, 2).setValues([[key, val]]);
      index[key] = writeRow;
    }
    // reminders may have been toggled -> refresh the trigger
    installReminderTrigger_();
    SpreadsheetApp.flush();
    return 'ok';
  } catch (err){ return 'error:' + String(err); }
}

// Write a single Settings key/value (used by category save).
function setSettingValue_(key, val){
  var sh = sheet_(getSS_(), 'Settings', SETTINGS_HEADERS);
  var existing = readAllSettings_();
  var row = 0;
  for (var i = 0; i < existing.length; i++){ if (existing[i].key === key){ row = existing[i].row; break; } }
  var writeRow = row ? row : sh.getLastRow() + 1;
  sh.getRange(writeRow, 1, 1, 2).setValues([[key, coerce_(val)]]);
}

// Save the product-category list, and apply any renames to existing products so
// nothing is orphaned. payload = { categories: [names...], renames: [{from,to}...] }.
function apiSaveCategories(payloadJson, pin){
  if (!requireOwner_(pin)) return 'error:auth';
  try {
    var p = JSON.parse(payloadJson || '{}');
    var cats = (p.categories instanceof Array) ? p.categories : [];
    var renames = (p.renames instanceof Array) ? p.renames : [];

    // apply renames to existing products (category column)
    var map = {};
    for (var i = 0; i < renames.length; i++){
      var f = String(renames[i].from || ''), t = String(renames[i].to || '');
      if (f && t && f !== t) map[f] = t;
    }
    var didRename = false;
    for (var k in map){ if (map.hasOwnProperty(k)){ didRename = true; break; } }
    if (didRename){
      var sh = sheet_(getSS_(), 'Products', PRODUCTS_HEADERS);
      var all = readAll_('Products', PRODUCTS_HEADERS);
      var catCol = PRODUCTS_HEADERS.indexOf('category') + 1;   // = 2
      for (var j = 0; j < all.length; j++){
        var c = String(all[j].obj.category || '');
        if (map.hasOwnProperty(c)){
          var nv = map[c];
          if (/^[=+\-@]/.test(nv)) nv = "'" + nv;   // keep a formula-leading name as literal text
          sh.getRange(all[j].row, catCol, 1, 1).setValues([[nv]]);
        }
      }
    }

    setSettingValue_('categories', JSON.stringify(cats));
    SpreadsheetApp.flush();
    return 'ok';
  } catch (err){ return 'error:' + String(err); }
}

/* =========================================================================
 *  PRODUCTS
 * ========================================================================= */
function apiListProducts(pin){
  if (!requireOwner_(pin)) return '[]';
  var all = readAll_('Products', PRODUCTS_HEADERS);
  var out = [];
  for (var i = 0; i < all.length; i++){
    try { out.push(shapeProduct_(all[i].obj)); } catch (e){}
  }
  return JSON.stringify(out);
}

function apiGetProduct(id, pin){
  if (!requireOwner_(pin)) return 'null';
  var found = findRow_('Products', PRODUCTS_HEADERS, id);
  if (!found) return 'null';
  return JSON.stringify(shapeProduct_(found.obj));
}

/* =========================================================================
 *  PUBLIC STOREFRONT  (safe to share — no cost, no supplier, no profit)
 *  This is the ONLY product endpoint the shared link uses. It returns just
 *  the customer-facing fields for ACTIVE products, plus a small settings
 *  subset (name, tagline, WhatsApp, deposit term). Nothing private crosses.
 * ========================================================================= */
function shapeShopProduct_(o){
  var photos = [];
  try { photos = o.photos ? JSON.parse(o.photos) : []; } catch (e){ photos = []; }
  if (!(photos instanceof Array)) photos = [];
  var safe = [];
  for (var i = 0; i < photos.length; i++){
    var ph = photos[i] || {};
    safe.push({ url: String(ph.url || ''), thumb: String(ph.thumb || ''), openUrl: String(ph.openUrl || '') });
  }
  return {
    id: String(o.id),
    category: String(o.category || ''),
    name: String(o.name || ''),
    description: String(o.description || ''),
    price: Number(o.price) || 0,
    photos: safe
  };
}

function apiShop(){
  try {
    var all = readAll_('Products', PRODUCTS_HEADERS);
    var items = [];
    for (var i = 0; i < all.length; i++){
      try {
        var o = all[i].obj;
        if (String(o.active) === 'false') continue;   // hidden products stay hidden
        items.push(shapeShopProduct_(o));
      } catch (e){}
    }
    var s = getSettingsObj_();
    return JSON.stringify({
      products: items,
      settings: {
        businessName:  String(s.businessName || "God's Favor Collection"),
        tagline:       String(s.tagline || ''),
        businessPhone: String(s.businessPhone || ''),
        defaultTermDays: String(s.defaultTermDays || '21'),
        categories:    String(s.categories || '[]')
      }
    });
  } catch (err){ return JSON.stringify({ products: [], settings: {} }); }
}

// First configured category (used as the fallback when a product has no category).
function firstCategory_(){
  try {
    var a = JSON.parse(getSettingsObj_().categories || '[]');
    if (a instanceof Array && a.length) return String(a[0]);
  } catch (e){}
  return '';
}

function shapeProduct_(o){
  var photos = [];
  try { photos = o.photos ? JSON.parse(o.photos) : []; } catch (e){ photos = []; }
  if (!(photos instanceof Array)) photos = [];
  var price = Number(o.price) || 0;
  var cost  = Number(o.costPrice) || 0;
  var extra = Number(o.extraCost) || 0;
  return {
    id: String(o.id),
    category: String(o.category || ''),
    name: String(o.name || ''),
    description: String(o.description || ''),
    price: price,
    supplier: String(o.supplier || ''),
    costPrice: cost,
    extraCost: extra,
    extraNote: String(o.extraNote || ''),
    margin: Math.max(0, price - cost),                 // gross (price − supplier cost)
    netPerItem: Math.max(0, price - cost - extra),      // after per-item other costs
    photos: photos,
    active: String(o.active) === 'false' ? false : true,
    createdAt: String(o.createdAt || ''),
    updatedAt: String(o.updatedAt || '')
  };
}

function apiSaveProduct(payloadJson, pin){
  if (!requireOwner_(pin)) return 'error:auth';
  try {
    var p = JSON.parse(payloadJson || '{}');
    var sh = sheet_(getSS_(), 'Products', PRODUCTS_HEADERS);
    var found = p.id ? findRow_('Products', PRODUCTS_HEADERS, p.id) : null;
    var now = new Date().toISOString();
    var id = found ? found.obj.id : genId_('prod');

    // preserve existing photos unless payload explicitly carries them
    var photos = p.photos;
    if (!(photos instanceof Array) && found){
      try { photos = JSON.parse(found.obj.photos || '[]'); } catch (e){ photos = []; }
    }
    if (!(photos instanceof Array)) photos = [];

    var rec = {
      id: id,
      category: coerce_(p.category || firstCategory_()),
      name: coerce_(p.name || ''),
      description: coerce_(p.description || ''),
      price: Number(p.price) || 0,
      supplier: coerce_(p.supplier || ''),
      costPrice: Number(p.costPrice) || 0,
      extraCost: Number(p.extraCost) || 0,
      extraNote: coerce_(p.extraNote || ''),
      photos: JSON.stringify(photos),
      active: (p.active === false || p.active === 'false') ? 'false' : 'true',
      createdAt: found ? coerce_(found.obj.createdAt || now) : now,
      updatedAt: now
    };

    var row = [];
    for (var c = 0; c < PRODUCTS_HEADERS.length; c++){ row.push(coerce_(rec[PRODUCTS_HEADERS[c]])); }
    var writeRow = found ? found.row : sh.getLastRow() + 1;
    sh.getRange(writeRow, 1, 1, PRODUCTS_HEADERS.length).setValues([row]);
    SpreadsheetApp.flush();
    return id;
  } catch (err){ return 'error:' + String(err); }
}

function apiDeleteProduct(id, pin){
  if (!requireOwner_(pin)) return 'error:auth';
  try {
    var found = findRow_('Products', PRODUCTS_HEADERS, id);
    if (!found) return 'ok';
    sheet_(getSS_(), 'Products', PRODUCTS_HEADERS).deleteRow(found.row);
    SpreadsheetApp.flush();
    return 'ok';
  } catch (err){ return 'error:' + String(err); }
}

/* ---- photo upload / remove ---- */
// recordType is 'Products' (default) or 'Orders'. Both keep a `photos` JSON array,
// so a product or an order can each hold multiple photos.
function photoTarget_(recordType){
  if (String(recordType) === 'Orders') return { sheet: 'Orders', headers: ORDERS_HEADERS };
  return { sheet: 'Products', headers: PRODUCTS_HEADERS };
}
function attachPhoto_(cfg, recordId, photo){
  var found = findRow_(cfg.sheet, cfg.headers, recordId);
  if (!found) return;
  var photos = [];
  try { photos = JSON.parse(found.obj.photos || '[]'); } catch (e){ photos = []; }
  if (!(photos instanceof Array)) photos = [];
  photos.push(photo);
  var col = cfg.headers.indexOf('photos') + 1;
  sheet_(getSS_(), cfg.sheet, cfg.headers).getRange(found.row, col, 1, 1).setValues([[JSON.stringify(photos)]]);
  SpreadsheetApp.flush();
}

function apiUploadPhoto(recordType, recordId, recordName, base64, mimeType, fileName, pin){
  if (!requireOwner_(pin)) return JSON.stringify({ error: 'auth' });
  try {
    var cfg = photoTarget_(recordType);
    var folder = ensureFolder_();
    var subName = (String(recordName || 'item')).replace(/[^\w\- ]/g, '').trim() + ' (' + recordId + ')';
    var it = folder.getFoldersByName(subName);
    var sub = it.hasNext() ? it.next() : folder.createFolder(subName);
    var bytes = Utilities.base64Decode(base64);                 // NO charset arg — it corrupts images
    var blob = Utilities.newBlob(bytes, mimeType || 'image/jpeg', fileName || (Date.now() + '.jpg'));
    var file = sub.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var fid = file.getId();
    var photo = {
      id: fid,
      name: file.getName(),
      url: 'https://lh3.googleusercontent.com/d/' + fid + '=w1000',
      thumb: 'https://drive.google.com/thumbnail?id=' + fid + '&sz=w1000',
      openUrl: 'https://drive.google.com/file/d/' + fid + '/view',
      addedAt: new Date().toISOString()
    };
    if (recordId) attachPhoto_(cfg, recordId, photo);
    return JSON.stringify(photo);
  } catch (err){ return JSON.stringify({ error: String(err) }); }
}

// Return a photo's raw bytes as base64 so the phone can attach the real image to a
// WhatsApp share (fetching the Google image URL directly is blocked by CORS; this goes
// through google.script.run instead).
function apiGetPhotoBytes(fileId, pin){
  if (!requireOwner_(pin)) return JSON.stringify({ error: 'auth' });
  try{
    var file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    return JSON.stringify({
      b64: Utilities.base64Encode(blob.getBytes()),
      mime: blob.getContentType() || 'image/jpeg',
      name: file.getName() || 'product.jpg'
    });
  }catch(err){ return JSON.stringify({ error: String(err) }); }
}

function apiRemovePhoto(recordType, recordId, fileId, pin){
  if (!requireOwner_(pin)) return 'error:auth';
  try {
    var cfg = photoTarget_(recordType);
    var found = findRow_(cfg.sheet, cfg.headers, recordId);
    if (!found) return 'ok';
    var photos = [];
    try { photos = JSON.parse(found.obj.photos || '[]'); } catch (e){ photos = []; }
    var kept = [];
    for (var i = 0; i < photos.length; i++){ if (String(photos[i].id) !== String(fileId)) kept.push(photos[i]); }
    var col = cfg.headers.indexOf('photos') + 1;
    sheet_(getSS_(), cfg.sheet, cfg.headers).getRange(found.row, col, 1, 1).setValues([[JSON.stringify(kept)]]);
    SpreadsheetApp.flush();
    try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e){}
    return 'ok';
  } catch (err){ return 'error:' + String(err); }
}

/* =========================================================================
 *  ORDERS
 * ========================================================================= */
function apiListOrders(pin){
  if (!requireOwner_(pin)) return '[]';
  var all = readAll_('Orders', ORDERS_HEADERS);
  var out = [];
  for (var i = 0; i < all.length; i++){
    try { out.push(shapeOrder_(all[i].obj)); } catch (e){}
  }
  return JSON.stringify(out);
}

function apiGetOrder(id, pin){
  if (!requireOwner_(pin)) return 'null';
  var found = findRow_('Orders', ORDERS_HEADERS, id);
  if (!found) return 'null';
  return JSON.stringify(shapeOrder_(found.obj));
}

function shapeOrder_(o){
  var total   = Number(o.total) || 0;
  var deposit = Number(o.deposit) || 0;
  var balance = Math.max(0, total - deposit);       // never negative
  var photos = [];
  try { photos = o.photos ? JSON.parse(o.photos) : []; } catch (e){ photos = []; }
  if (!(photos instanceof Array)) photos = [];
  var lineItems = [];
  try { lineItems = o.lineItems ? JSON.parse(o.lineItems) : []; } catch (e){ lineItems = []; }
  if (!(lineItems instanceof Array)) lineItems = [];
  return {
    id: String(o.id),
    photos: photos,
    lineItems: lineItems,
    customerName: String(o.customerName || ''),
    phone: String(o.phone || ''),
    items: String(o.items || ''),
    total: total,
    deposit: deposit,
    balance: balance,
    orderDate: isoDate_(o.orderDate),
    depositDate: isoDate_(o.depositDate),
    dueDate: isoDate_(o.dueDate),
    status: String(o.status || 'Deposit Paid'),
    remindersOn: String(o.remindersOn) === 'false' ? false : true,
    notes: String(o.notes || ''),
    createdAt: String(o.createdAt || ''),
    updatedAt: String(o.updatedAt || '')
  };
}

function apiSaveOrder(payloadJson, pin){
  if (!requireOwner_(pin)) return 'error:auth';
  try {
    var p = JSON.parse(payloadJson || '{}');
    var sh = sheet_(getSS_(), 'Orders', ORDERS_HEADERS);
    var found = p.id ? findRow_('Orders', ORDERS_HEADERS, p.id) : null;
    var now = new Date().toISOString();
    var id = found ? found.obj.id : genId_('ord');

    var total   = Number(p.total) || 0;
    var deposit = Number(p.deposit) || 0;
    var balance = Math.max(0, total - deposit);

    var orderDate = isoDate_(p.orderDate) || isoDate_(new Date());
    var dueDate = isoDate_(p.dueDate);
    if (!dueDate){
      var termDays = parseInt(getSettingsObj_().defaultTermDays, 10) || 21;
      dueDate = addDaysISO_(orderDate, termDays);
    }

    // preserve existing photos unless the payload explicitly carries them
    var photos = p.photos;
    if (!(photos instanceof Array) && found){
      try { photos = JSON.parse(found.obj.photos || '[]'); } catch (e){ photos = []; }
    }
    if (!(photos instanceof Array)) photos = [];

    var lineItems = p.lineItems;
    if (!(lineItems instanceof Array) && found){
      try { lineItems = JSON.parse(found.obj.lineItems || '[]'); } catch (e){ lineItems = []; }
    }
    if (!(lineItems instanceof Array)) lineItems = [];

    var rec = {
      id: id,
      photos: JSON.stringify(photos),
      lineItems: JSON.stringify(lineItems),
      customerName: coerce_(p.customerName || ''),
      phone: coerce_(p.phone || ''),
      items: coerce_(p.items || ''),
      total: total,
      deposit: deposit,
      balance: balance,
      orderDate: orderDate,
      depositDate: isoDate_(p.depositDate) || orderDate,
      dueDate: dueDate,
      status: coerce_(p.status || (balance <= 0 ? 'Paid in Full' : 'Deposit Paid')),
      remindersOn: (p.remindersOn === false || p.remindersOn === 'false') ? 'false' : 'true',
      notes: coerce_(p.notes || ''),
      createdAt: found ? coerce_(found.obj.createdAt || now) : now,
      updatedAt: now
    };

    var row = [];
    for (var c = 0; c < ORDERS_HEADERS.length; c++){ row.push(coerce_(rec[ORDERS_HEADERS[c]])); }
    var writeRow = found ? found.row : sh.getLastRow() + 1;
    sh.getRange(writeRow, 1, 1, ORDERS_HEADERS.length).setValues([row]);
    SpreadsheetApp.flush();
    return id;
  } catch (err){ return 'error:' + String(err); }
}

function apiDeleteOrder(id, pin){
  if (!requireOwner_(pin)) return 'error:auth';
  try {
    var found = findRow_('Orders', ORDERS_HEADERS, id);
    if (!found) return 'ok';
    sheet_(getSS_(), 'Orders', ORDERS_HEADERS).deleteRow(found.row);
    SpreadsheetApp.flush();
    return 'ok';
  } catch (err){ return 'error:' + String(err); }
}

/* =========================================================================
 *  EXPENSES
 * ========================================================================= */
function shapeExpense_(o){
  return {
    id: String(o.id),
    date: isoDate_(o.date),
    description: String(o.description || ''),
    amount: Number(o.amount) || 0,
    productId: String(o.productId || ''),
    category: String(o.category || ''),
    createdAt: String(o.createdAt || '')
  };
}

function apiListExpenses(pin){
  if (!requireOwner_(pin)) return '[]';
  var all = readAll_('Expenses', EXPENSES_HEADERS);
  var out = [];
  for (var i = 0; i < all.length; i++){ try { out.push(shapeExpense_(all[i].obj)); } catch (e){} }
  return JSON.stringify(out);
}

function apiSaveExpense(payloadJson, pin){
  if (!requireOwner_(pin)) return 'error:auth';
  try {
    var p = JSON.parse(payloadJson || '{}');
    var sh = sheet_(getSS_(), 'Expenses', EXPENSES_HEADERS);
    var found = p.id ? findRow_('Expenses', EXPENSES_HEADERS, p.id) : null;
    var now = new Date().toISOString();
    var id = found ? found.obj.id : genId_('exp');
    var rec = {
      id: id,
      date: isoDate_(p.date) || isoDate_(new Date()),
      description: coerce_(p.description || ''),
      amount: Number(p.amount) || 0,
      productId: coerce_(p.productId || ''),
      category: coerce_(p.category || ''),
      createdAt: found ? coerce_(found.obj.createdAt || now) : now
    };
    var row = [];
    for (var c = 0; c < EXPENSES_HEADERS.length; c++){ row.push(coerce_(rec[EXPENSES_HEADERS[c]])); }
    var writeRow = found ? found.row : sh.getLastRow() + 1;
    sh.getRange(writeRow, 1, 1, EXPENSES_HEADERS.length).setValues([row]);
    SpreadsheetApp.flush();
    return id;
  } catch (err){ return 'error:' + String(err); }
}

function apiDeleteExpense(id, pin){
  if (!requireOwner_(pin)) return 'error:auth';
  try {
    var found = findRow_('Expenses', EXPENSES_HEADERS, id);
    if (!found) return 'ok';
    sheet_(getSS_(), 'Expenses', EXPENSES_HEADERS).deleteRow(found.row);
    SpreadsheetApp.flush();
    return 'ok';
  } catch (err){ return 'error:' + String(err); }
}

/* =========================================================================
 *  PROFIT SUMMARY  (business-wide)
 *  netProfit = money actually received − everything logged in Expenses.
 * ========================================================================= */
function apiProfit(pin){
  if (!requireOwner_(pin)) return JSON.stringify({totalSales:0,cogs:0,totalExpenses:0,netProfit:0,moneyReceived:0,outstanding:0,expenseCount:0});
  try {
    // product cost map (supplier cost + other cost per item) for line items lacking a snapshot cost
    var costMap = {};
    var prods = readAll_('Products', PRODUCTS_HEADERS);
    for (var pi = 0; pi < prods.length; pi++){
      var po; try { po = shapeProduct_(prods[pi].obj); } catch (e){ continue; }
      costMap[po.id] = (Number(po.costPrice) || 0) + (Number(po.extraCost) || 0);
    }

    var moneyReceived = 0, outstanding = 0, totalSales = 0, cogs = 0;
    var orders = readAll_('Orders', ORDERS_HEADERS);
    for (var i = 0; i < orders.length; i++){
      var o; try { o = shapeOrder_(orders[i].obj); } catch (e){ continue; }
      if (o.status === 'Cancelled') continue;
      moneyReceived += o.deposit;      // amount paid so far (paid-in-full sets deposit = total)
      outstanding   += o.balance;
      totalSales    += o.total;
      var lis = o.lineItems || [];
      for (var li = 0; li < lis.length; li++){
        var item = lis[li];
        var unit = (item.cost !== undefined && item.cost !== null && item.cost !== '')
                   ? Number(item.cost) : (costMap[item.productId] || 0);
        cogs += (Number(unit) || 0) * (Number(item.qty) || 0);
      }
    }
    var totalExpenses = 0, count = 0;
    var exps = readAll_('Expenses', EXPENSES_HEADERS);
    for (var j = 0; j < exps.length; j++){
      var e; try { e = shapeExpense_(exps[j].obj); } catch (er){ continue; }
      totalExpenses += e.amount; count++;
    }
    // Net profit = booked sales − cost of goods − other expenses (assumes orders get fully paid).
    return JSON.stringify({
      totalSales: totalSales,
      cogs: cogs,
      totalExpenses: totalExpenses,
      netProfit: totalSales - cogs - totalExpenses,
      moneyReceived: moneyReceived,
      outstanding: outstanding,
      expenseCount: count
    });
  } catch (err){
    return JSON.stringify({totalSales:0,cogs:0,totalExpenses:0,netProfit:0,moneyReceived:0,outstanding:0,expenseCount:0});
  }
}

/* =========================================================================
 *  DASHBOARD
 * ========================================================================= */
function apiDashboard(pin){
  if (!requireOwner_(pin)) return JSON.stringify({ products: 0, openOrders: 0, outstanding: 0, dueThisWeek: 0, overdue: 0 });
  try {
    var orders = [];
    var all = readAll_('Orders', ORDERS_HEADERS);
    for (var i = 0; i < all.length; i++){ try { orders.push(shapeOrder_(all[i].obj)); } catch (e){} }
    var products = readAll_('Products', PRODUCTS_HEADERS);

    var outstanding = 0, dueThisWeek = 0, overdue = 0, openOrders = 0;
    var todayISO = isoDate_(new Date());
    for (var j = 0; j < orders.length; j++){
      var o = orders[j];
      if (o.status === 'Completed' || o.status === 'Cancelled') continue;
      if (o.balance > 0){
        outstanding += o.balance;
        openOrders++;
        var days = daysBetweenISO_(todayISO, o.dueDate);
        if (days < 0) overdue++;
        else if (days <= 7) dueThisWeek++;
      }
    }
    return JSON.stringify({
      products: products.length,
      openOrders: openOrders,
      outstanding: outstanding,
      dueThisWeek: dueThisWeek,
      overdue: overdue,
      today: todayISO
    });
  } catch (err){ return JSON.stringify({ products: 0, openOrders: 0, outstanding: 0, dueThisWeek: 0, overdue: 0 }); }
}

/* =========================================================================
 *  CUSTOMER REQUESTS  (interest submitted from the public storefront)
 *  apiSubmitInterest is PUBLIC (a shopper has no PIN). It only ever WRITES a
 *  request row and emails the owner — it can read nothing private. The owner's
 *  Requests list / status update are PIN-gated like everything else.
 * ========================================================================= */
function shapeRequest_(o){
  return {
    id: String(o.id),
    createdAt: String(o.createdAt || ''),
    productId: String(o.productId || ''),
    productName: String(o.productName || ''),
    price: Number(o.price) || 0,
    customerName: String(o.customerName || ''),
    phone: String(o.phone || ''),
    offer: String(o.offer || ''),
    note: String(o.note || ''),
    status: String(o.status || 'New'),
    handledAt: String(o.handledAt || '')
  };
}

// Keep customer-supplied text safe: trim, cap length, and neutralise a leading
// spreadsheet-formula character so a note like "=cmd" can't become a live formula.
function cleanField_(v, max){
  var raw = String(v == null ? '' : v);
  var out = '';
  for (var i = 0; i < raw.length; i++){
    var code = raw.charCodeAt(i);
    out += (code < 32 || code === 127) ? ' ' : raw.charAt(i);   // strip control chars + DEL
  }
  out = out.trim();
  if (out.length > max) out = out.slice(0, max);
  if (/^[=+\-@]/.test(out)) out = "'" + out;   // neutralise a leading spreadsheet-formula char
  return out;
}

function apiSubmitInterest(payloadJson){
  try {
    var p = JSON.parse(payloadJson || '{}');

    // Validate against the live catalog — only a real, active product can be requested.
    var prod = null;
    var all = readAll_('Products', PRODUCTS_HEADERS);
    for (var i = 0; i < all.length; i++){
      if (String(all[i].obj.id) === String(p.productId) && String(all[i].obj.active) !== 'false'){
        prod = shapeProduct_(all[i].obj); break;
      }
    }
    if (!prod) return JSON.stringify({ ok: false, err: 'That item is no longer available.' });

    var name  = cleanField_(p.customerName, 80);
    var phone = cleanField_(p.phone, 40);
    var note  = cleanField_(p.note, 500);
    var offer = cleanField_(p.offer, 20);
    if (!name && !phone) return JSON.stringify({ ok: false, err: 'Please add your name or WhatsApp number.' });

    var id = genId_('req');
    var now = new Date().toISOString();
    var rec = {
      id: id, createdAt: now,
      productId: prod.id, productName: prod.name, price: prod.price,
      customerName: name, phone: phone, offer: offer, note: note,
      status: 'New', handledAt: ''
    };
    var sh = sheet_(getSS_(), 'Requests', REQUESTS_HEADERS);
    var row = [];
    for (var c = 0; c < REQUESTS_HEADERS.length; c++){ row.push(coerce_(rec[REQUESTS_HEADERS[c]])); }
    sh.getRange(sh.getLastRow() + 1, 1, 1, REQUESTS_HEADERS.length).setValues([row]);
    SpreadsheetApp.flush();

    try { emailNewRequest_(shapeRequest_(rec)); } catch (mailErr){ /* record still saved */ }
    return JSON.stringify({ ok: true, id: id });
  } catch (err){ return JSON.stringify({ ok: false, err: 'Please try again.' }); }
}

function apiListRequests(pin){
  if (!requireOwner_(pin)) return '[]';
  var all = readAll_('Requests', REQUESTS_HEADERS);
  var out = [];
  for (var i = 0; i < all.length; i++){ try { out.push(shapeRequest_(all[i].obj)); } catch (e){} }
  out.reverse();   // newest first
  return JSON.stringify(out);
}

function apiUpdateRequestStatus(id, status, pin){
  if (!requireOwner_(pin)) return 'error:auth';
  try {
    var allowed = { 'New': 1, 'Accepted': 1, 'Declined': 1 };
    var st = allowed[String(status)] ? String(status) : 'New';
    var found = findRow_('Requests', REQUESTS_HEADERS, id);
    if (!found) return 'ok';
    var sh = sheet_(getSS_(), 'Requests', REQUESTS_HEADERS);
    var stCol = REQUESTS_HEADERS.indexOf('status') + 1;
    var hCol  = REQUESTS_HEADERS.indexOf('handledAt') + 1;
    sh.getRange(found.row, stCol, 1, 1).setValues([[st]]);
    sh.getRange(found.row, hCol, 1, 1).setValues([[st === 'New' ? '' : new Date().toISOString()]]);
    SpreadsheetApp.flush();
    return 'ok';
  } catch (err){ return 'error:' + String(err); }
}

function emailNewRequest_(r){
  var s = getSettingsObj_();
  var email = String(s.ownerEmail || '').trim();
  if (!email) return;
  var money = function(n){ return '$' + (Number(n) || 0).toLocaleString(); };
  var offerLine = r.offer ? ('<tr><td style="padding:6px 12px;color:#666">Willing to pay</td><td style="padding:6px 12px;font-weight:700">$' + esc_(r.offer) + '</td></tr>') : '';
  var noteLine  = r.note ? ('<tr><td style="padding:6px 12px;color:#666">Note</td><td style="padding:6px 12px">' + esc_(r.note) + '</td></tr>') : '';
  var phoneLine = r.phone ? ('<tr><td style="padding:6px 12px;color:#666">WhatsApp</td><td style="padding:6px 12px;font-weight:700">' + esc_(r.phone) + '</td></tr>') : '';
  var html = '' +
    '<div style="font-family:Arial,Helvetica,sans-serif;max-width:560px;margin:0 auto">' +
      '<div style="background:linear-gradient(135deg,#1B6FA8,#2A93CE);color:#fff;padding:18px 16px;border-radius:12px 12px 0 0">' +
        '<div style="font-size:19px;font-weight:800">God’s Favor Collection</div>' +
        '<div style="opacity:.9;font-size:13px">New customer request 🔔</div>' +
      '</div>' +
      '<div style="border:1px solid #eee;border-top:0;border-radius:0 0 12px 12px;padding:10px 4px">' +
        '<table style="width:100%;border-collapse:collapse;font-size:14px">' +
          '<tr><td style="padding:6px 12px;color:#666">Item</td><td style="padding:6px 12px;font-weight:700">' + esc_(r.productName) + ' (' + money(r.price) + ')</td></tr>' +
          '<tr><td style="padding:6px 12px;color:#666">From</td><td style="padding:6px 12px;font-weight:700">' + esc_(r.customerName || '—') + '</td></tr>' +
          phoneLine + offerLine + noteLine +
        '</table>' +
      '</div>' +
      '<p style="font-size:13px;color:#444;margin:14px 4px">Open your app, go to <b>Requests</b>, and tap <b>Accept</b> to turn this into an order (and send them a WhatsApp confirmation), or <b>Decline</b>.</p>' +
      '<p style="font-size:12px;color:#999;text-align:center;margin-top:16px">Cornerstone Studio · Glory to God alone</p>' +
    '</div>';
  MailApp.sendEmail({ to: email, subject: "New request — " + (r.customerName || 'customer') + " · " + r.productName, htmlBody: html });
}

/* =========================================================================
 *  DATE HELPERS  (server side, ISO strings only)
 * ========================================================================= */
function isoDate_(v){
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var s = String(v);
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  var d = new Date(s);
  if (!isNaN(d.getTime())) return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return '';
}

function addDaysISO_(iso, days){
  var parts = String(iso).slice(0, 10).split('-');
  var d = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
  d.setDate(d.getDate() + days);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function daysBetweenISO_(fromISO, toISO){
  var a = String(fromISO).slice(0, 10).split('-');
  var b = String(toISO).slice(0, 10).split('-');
  var da = new Date(parseInt(a[0], 10), parseInt(a[1], 10) - 1, parseInt(a[2], 10));
  var db = new Date(parseInt(b[0], 10), parseInt(b[1], 10) - 1, parseInt(b[2], 10));
  return Math.round((db.getTime() - da.getTime()) / 86400000);
}

/* =========================================================================
 *  REMINDERS  (weekly Friday digest email)
 * ========================================================================= */
function installReminderTrigger_(){
  // remove existing sendReminders triggers
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++){
    var h = triggers[i].getHandlerFunction();
    if (h === 'sendReminders' || h === 'sendReminders_') ScriptApp.deleteTrigger(triggers[i]);
  }
  var s = getSettingsObj_();
  if (String(s.remindersOn).toLowerCase() === 'no' || String(s.remindersOn).toLowerCase() === 'false') return;

  var dayMap = {
    'sunday': ScriptApp.WeekDay.SUNDAY, 'monday': ScriptApp.WeekDay.MONDAY,
    'tuesday': ScriptApp.WeekDay.TUESDAY, 'wednesday': ScriptApp.WeekDay.WEDNESDAY,
    'thursday': ScriptApp.WeekDay.THURSDAY, 'friday': ScriptApp.WeekDay.FRIDAY,
    'saturday': ScriptApp.WeekDay.SATURDAY
  };
  var wd = dayMap[String(s.reminderDay || 'Friday').toLowerCase()] || ScriptApp.WeekDay.FRIDAY;
  ScriptApp.newTrigger('sendReminders_').timeBased().onWeekDay(wd).atHour(7).create();
}

function sendReminders_(){
  try {
    var s = getSettingsObj_();
    if (String(s.remindersOn).toLowerCase() === 'no' || String(s.remindersOn).toLowerCase() === 'false') return;
    var email = String(s.ownerEmail || '').trim();
    if (!email) return;

    var all = readAll_('Orders', ORDERS_HEADERS);
    var due = [];
    var todayISO = isoDate_(new Date());
    for (var i = 0; i < all.length; i++){
      var o;
      try { o = shapeOrder_(all[i].obj); } catch (e){ continue; }
      if (o.status === 'Completed' || o.status === 'Cancelled') continue;
      if (!o.remindersOn) continue;
      if (o.balance <= 0) continue;
      o._days = daysBetweenISO_(todayISO, o.dueDate);
      due.push(o);
    }
    if (!due.length) return;

    due.sort(function(a, b){ return a._days - b._days; });   // most overdue first

    var html = buildDigestHtml_(due, s);
    var subject = "God's Favor Collection — " + due.length + " customer" + (due.length > 1 ? 's' : '') + " with a balance to follow up";
    MailApp.sendEmail({ to: email, subject: subject, htmlBody: html });

    // log it
    var log = sheet_(getSS_(), 'ReminderLog', LOG_HEADERS);
    var stamp = new Date().toISOString();
    var logRows = [];
    for (var j = 0; j < due.length; j++){
      logRows.push([stamp, due[j].id, due[j].customerName, 'balance ' + due[j].balance + ' due ' + due[j].dueDate]);
    }
    if (logRows.length){
      log.getRange(log.getLastRow() + 1, 1, logRows.length, LOG_HEADERS.length).setValues(logRows);
    }
    SpreadsheetApp.flush();
  } catch (err){ /* never throw from a trigger */ }
}

function buildDigestHtml_(due, s){
  var money = function(n){ return '$' + (Number(n) || 0).toLocaleString(); };  // USD
  var rows = '';
  for (var i = 0; i < due.length; i++){
    var o = due[i];
    var tag = o._days < 0 ? ('<span style="color:#B00020;font-weight:700">' + Math.abs(o._days) + ' day(s) overdue</span>')
            : o._days === 0 ? '<span style="color:#B8860B;font-weight:700">due today</span>'
            : ('due in ' + o._days + ' day(s)');
    rows +=
      '<tr>' +
        '<td style="padding:10px 12px;border-bottom:1px solid #eee">' +
          '<div style="font-weight:700;color:#1B6FA8">' + esc_(o.customerName) + '</div>' +
          '<div style="font-size:12px;color:#666">' + esc_(o.items || '') + '</div>' +
        '</td>' +
        '<td style="padding:10px 12px;border-bottom:1px solid #eee;text-align:right;white-space:nowrap">' +
          '<div style="font-weight:700">' + money(o.balance) + '</div>' +
          '<div style="font-size:12px;color:#666">' + tag + '</div>' +
        '</td>' +
      '</tr>';
  }
  return '' +
    '<div style="font-family:Arial,Helvetica,sans-serif;max-width:600px;margin:0 auto">' +
      '<div style="background:linear-gradient(135deg,#1B6FA8,#2A93CE);color:#fff;padding:20px 16px;border-radius:12px 12px 0 0">' +
        '<div style="font-size:20px;font-weight:800;letter-spacing:.5px">God’s Favor Collection</div>' +
        '<div style="opacity:.9;font-size:13px">Weekly follow-up list</div>' +
      '</div>' +
      '<div style="border:1px solid #eee;border-top:0;border-radius:0 0 12px 12px;padding:8px 0">' +
        '<table style="width:100%;border-collapse:collapse">' + rows + '</table>' +
      '</div>' +
      '<p style="font-size:13px;color:#444;margin:16px 4px">Open the app, go to <b>Orders</b>, and tap <b>WhatsApp</b> next to each customer to send the reminder.</p>' +
      '<p style="font-size:12px;color:#999;text-align:center;margin-top:18px">' +
        '“Commit your work to the LORD, and your plans will be established.” — Proverbs 16:3<br>' +
        'Cornerstone Studio · Glory to God alone' +
      '</p>' +
    '</div>';
}

function esc_(s){
  return String(s || '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// Manual test from the app: send the digest right now (ignores the day-of-week schedule).
function apiSendReminderNow(pin){
  if (!requireOwner_(pin)) return 'error:auth';
  try { sendReminders_(); return 'ok'; } catch (err){ return 'error:' + String(err); }
}
