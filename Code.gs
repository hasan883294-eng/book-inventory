/********************************************************************
 * انبار کتاب — نسخه API برای GitHub Pages (با رمز ورود + ورود گروهی از اکسل)
 * ------------------------------------------------------------------
 * دیتابیس: همین گوگل‌شیت (داخل گوگل درایو شما)
 * شیت‌ها: books | channels | transactions
 *
 * ورودی: خرید / مرجوعی
 * خروجی: حضوری / دیجی‌کالا / کانال‌های دلخواه
 * موجودی = جمع ورودی − جمع خروجی  (هشدار: کمتر از ۵)
 *
 * امنیت: رمز پنل در Script Properties با کلید SITE_PASSWORD نگهداری
 * می‌شود (پیش‌فرض admin1234 — حتماً عوضش کنید). نشست کاربر توکن
 * امضاشده HMAC است که در مرورگرش نگه‌داری می‌شود.
 ********************************************************************/

const SHEETS_ = { BOOKS: "books", CHANNELS: "channels", TX: "transactions" };
const LOW_STOCK_ = 5;                 // آستانه هشدار موجودی
const TZ_ = "Asia/Tehran";            // ساعت رسمی ایران
const SESSION_DAYS_ = 7;              // مدت اعتبار نشست
const IMPORT_MAX_ROWS_ = 2000;

/* ------------------------ وب‌اپ ------------------------ */

function doGet() {
  return json_({ ok: true, message: "API انبار کتاب فعال است. از صفحه سایت استفاده کنید." });
}

/* ---- API: فرانت‌اند (GitHub Pages) با POST و بدنه JSON صدا می‌زند ---- */
const API_METHODS_ = {
  apiLogin: apiLogin, apiCheck: apiCheck, apiGetState: apiGetState,
  apiLookupBarcode: apiLookupBarcode, apiAddBook: apiAddBook, apiUpdateBook: apiUpdateBook,
  apiDeleteBook: apiDeleteBook, apiAddChannel: apiAddChannel, apiDeleteChannel: apiDeleteChannel,
  apiAddTransaction: apiAddTransaction, apiListTransactions: apiListTransactions,
  apiImportBooks: apiImportBooks, apiImportExcel: apiImportExcel, apiGetSampleFile: apiGetSampleFile
};

function doPost(e) {
  try {
    ensureSetup_();
    const body = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    const fn = API_METHODS_[String(body.fn || "")];
    if (!fn) throw new Error("درخواست نامعتبر");
    const args = Array.isArray(body.args) ? body.args : [];
    const result = fn.apply(null, args);
    return json_({ ok: true, result: result === undefined ? null : result });
  } catch (err) {
    return json_({ ok: false, error: (err && err.message) ? err.message : "خطای سرور" });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** یک بار از ادیتور اجرا کنید تا دسترسی‌ها داده شود */
function setup() {
  ensureSetup_();
  const p = PropertiesService.getScriptProperties();
  if (!p.getProperty("SITE_PASSWORD")) {
    p.setProperty("SITE_PASSWORD", "admin1234");
  }
  getSecret_();
  return "OK — شیت‌ها آماده شدند. رمز پیش‌فرض admin1234 است؛ از Project Settings → Script Properties کلید SITE_PASSWORD را تغییر دهید.";
}

/* -------------------------- احراز هویت -------------------------- */

function getPwd_() {
  return PropertiesService.getScriptProperties().getProperty("SITE_PASSWORD") || "admin1234";
}

function getSecret_() {
  const p = PropertiesService.getScriptProperties();
  let s = p.getProperty("SESSION_SECRET");
  if (!s) {
    s = Utilities.getUuid() + Utilities.getUuid();
    p.setProperty("SESSION_SECRET", s);
  }
  return s;
}

function hmac_(msg) {
  const bytes = Utilities.computeHmacSha256Signature(String(msg), getSecret_());
  return Utilities.base64EncodeWebSafe(bytes);
}

function makeToken_() {
  const exp = Date.now() + SESSION_DAYS_ * 24 * 3600 * 1000;
  return exp + "." + hmac_(String(exp));
}

function tokenOk_(t) {
  if (!t) return false;
  const parts = String(t).split(".");
  if (parts.length !== 2) return false;
  const exp = Number(parts[0]);
  if (!isFinite(exp) || exp < Date.now()) return false;
  return parts[1] === hmac_(parts[0]);
}

function assertAuth_(token) {
  if (!tokenOk_(token)) throw new Error("نشست شما منقضی شده؛ لطفاً دوباره وارد شوید");
}

/** ورود با رمز پنل — برمی‌گرداند توکن نشست */
function apiLogin(token, password) {
  if (String(password || "") !== getPwd_()) {
    Utilities.sleep(600); // کندکردن حدس‌زدن رمز
    throw new Error("رمز عبور اشتباه است");
  }
  return { token: makeToken_(), days: SESSION_DAYS_ };
}

/** بررسی اعتبار توکن (برای بوت اولیه رابط) */
function apiCheck(token) {
  return { ok: tokenOk_(token) };
}

/* ---------------------- آماده‌سازی شیت‌ها ---------------------- */

function ensureSetup_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  getOrCreate_(ss, SHEETS_.BOOKS, ["id", "barcode", "title", "author", "publisher", "dkpc", "createdAt"]);
  const ch = getOrCreate_(ss, SHEETS_.CHANNELS, ["id", "name", "direction", "isSystem", "createdAt"]);
  getOrCreate_(ss, SHEETS_.TX, ["id", "ts", "bookId", "barcode", "title", "direction", "channelId", "channelName", "quantity", "note"]);

  if (ch.getLastRow() < 2) {
    const now = Date.now();
    [["خرید", "in"], ["مرجوعی", "in"], ["حضوری", "out"], ["دیجی‌کالا", "out"]].forEach(function (r, i) {
      ch.appendRow([i + 1, r[0], r[1], true, now]);
    });
  }
}

function getOrCreate_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  }
  return sh;
}

/* ------------------------ ابزار خواندن ------------------------ */

function rowsOf_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  return vals
    .filter(function (r) { return String(r[0]) !== ""; })
    .map(function (r) {
      const o = {};
      head.forEach(function (h, i) { o[h] = r[i]; });
      return o;
    });
}

function nextId_(sh) {
  if (sh.getLastRow() < 2) return 1;
  const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().flat().map(Number);
  const max = Math.max.apply(null, ids.concat([0]));
  return max + 1;
}

/** نقشه موجودی: bookId -> {in, out, stock} */
function stockMap_() {
  const map = {};
  rowsOf_(SHEETS_.TX).forEach(function (t) {
    const id = String(t.bookId);
    if (!map[id]) map[id] = { totalIn: 0, totalOut: 0 };
    if (t.direction === "in") map[id].totalIn += Number(t.quantity) || 0;
    else map[id].totalOut += Number(t.quantity) || 0;
  });
  Object.keys(map).forEach(function (k) {
    map[k].stock = map[k].totalIn - map[k].totalOut;
  });
  return map;
}

/* ------------------------ API برای رابط کاربری ------------------------ */

/** وضعیت کامل: کتاب‌ها با موجودی + کانال‌ها + آمار امروز + آخرین گردش‌ها */
function apiGetState(token) {
  assertAuth_(token);
  ensureSetup_();
  const stocks = stockMap_();
  const books = rowsOf_(SHEETS_.BOOKS).map(function (b) {
    const s = stocks[String(b.id)] || { totalIn: 0, totalOut: 0, stock: 0 };
    return {
      id: Number(b.id),
      barcode: String(b.barcode),
      title: String(b.title),
      author: b.author ? String(b.author) : "",
      publisher: b.publisher ? String(b.publisher) : "",
      dkpc: b.dkpc ? String(b.dkpc) : "",
      totalIn: s.totalIn,
      totalOut: s.totalOut,
      stock: s.stock
    };
  });
  books.sort(function (a, b) { return a.title.localeCompare(b.title, "fa"); });

  const channels = rowsOf_(SHEETS_.CHANNELS).map(function (c) {
    return { id: Number(c.id), name: String(c.name), direction: String(c.direction), isSystem: c.isSystem === true || String(c.isSystem) === "TRUE" };
  });

  const tx = rowsOf_(SHEETS_.TX);
  const today = Utilities.formatDate(new Date(), TZ_, "yyyy-MM-dd");
  let todayIn = 0, todayOut = 0;
  tx.forEach(function (t) {
    if (Utilities.formatDate(new Date(Number(t.ts)), TZ_, "yyyy-MM-dd") === today) {
      if (t.direction === "in") todayIn += Number(t.quantity) || 0;
      else todayOut += Number(t.quantity) || 0;
    }
  });

  const recent = tx.slice(-10).reverse().map(txView_);

  return {
    books: books,
    channels: channels,
    todayIn: todayIn,
    todayOut: todayOut,
    lowStockThreshold: LOW_STOCK_,
    recent: recent
  };
}

function txView_(t) {
  return {
    id: Number(t.id),
    ts: Number(t.ts),
    direction: String(t.direction),
    quantity: Number(t.quantity),
    note: t.note ? String(t.note) : "",
    book: { id: Number(t.bookId), title: String(t.title), barcode: String(t.barcode) },
    channel: { id: Number(t.channelId), name: String(t.channelName) }
  };
}

/** جست‌وجوی کتاب با بارکد */
function apiLookupBarcode(token, barcode) {
  assertAuth_(token);
  const code = String(barcode || "").trim();
  if (!code) throw new Error("بارکد ارسال نشده");
  const stocks = stockMap_();
  const found = rowsOf_(SHEETS_.BOOKS).filter(function (b) { return String(b.barcode) === code; })[0];
  if (!found) return { found: false, book: null };
  const s = stocks[String(found.id)] || { totalIn: 0, totalOut: 0, stock: 0 };
  return {
    found: true,
    book: {
      id: Number(found.id),
      barcode: code,
      title: String(found.title),
      author: found.author ? String(found.author) : "",
      publisher: found.publisher ? String(found.publisher) : "",
      dkpc: found.dkpc ? String(found.dkpc) : "",
      totalIn: s.totalIn,
      totalOut: s.totalOut,
      stock: s.stock
    }
  };
}

/** ثبت کتاب جدید */
function apiAddBook(token, p) {
  assertAuth_(token);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    ensureSetup_();
    const barcode = String(p.barcode || "").trim();
    const title = String(p.title || "").trim();
    if (!barcode) throw new Error("بارکد کتاب الزامی است");
    if (!title) throw new Error("عنوان کتاب الزامی است");
    const dup = rowsOf_(SHEETS_.BOOKS).some(function (b) { return String(b.barcode) === barcode; });
    if (dup) throw new Error("این بارکد قبلاً برای کتاب دیگری ثبت شده است");
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS_.BOOKS);
    const id = nextId_(sh);
    sh.appendRow([id, barcode, title, String(p.author || "").trim(), String(p.publisher || "").trim(), String(p.dkpc || "").trim(), Date.now()]);
    return { ok: true, id: id };
  } finally { lock.releaseLock(); }
}

/** ویرایش کتاب */
function apiUpdateBook(token, p) {
  assertAuth_(token);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const id = Number(p.id);
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS_.BOOKS);
    if (!sh || sh.getLastRow() < 2) throw new Error("کتاب یافت نشد");
    const barcode = String(p.barcode || "").trim();
    const title = String(p.title || "").trim();
    if (!barcode || !title) throw new Error("بارکد و عنوان نمی‌توانند خالی باشند");
    const lr = sh.getLastRow();
    const vals = sh.getRange(2, 1, lr - 1, 6).getValues();
    for (let i = 0; i < vals.length; i++) {
      if (Number(vals[i][0]) === id) {
        // یکتایی بارکد
        for (let j = 0; j < vals.length; j++) {
          if (j !== i && String(vals[j][1]) === barcode) throw new Error("این بارکد برای کتاب دیگری ثبت شده است");
        }
        sh.getRange(i + 2, 2, 1, 5).setValues([[barcode, title, String(p.author || "").trim(), String(p.publisher || "").trim(), String(p.dkpc || "").trim()]]);
        return { ok: true };
      }
    }
    throw new Error("کتاب یافت نشد");
  } finally { lock.releaseLock(); }
}

/** حذف کتاب — فقط اگر گردشی برایش ثبت نشده باشد */
function apiDeleteBook(token, id) {
  assertAuth_(token);
  const used = rowsOf_(SHEETS_.TX).some(function (t) { return Number(t.bookId) === Number(id); });
  if (used) throw new Error("برای این کتاب گردش انبار ثبت شده و حذف آن ممکن نیست؛ اطلاعاتش را ویرایش کنید");
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS_.BOOKS);
  const lr = sh.getLastRow();
  if (lr < 2) throw new Error("کتاب یافت نشد");
  const vals = sh.getRange(2, 1, lr - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (Number(vals[i][0]) === Number(id)) {
      sh.deleteRow(i + 2);
      return { ok: true };
    }
  }
  throw new Error("کتاب یافت نشد");
}

/** افزودن کانال ورودی/خروجی دلخواه */
function apiAddChannel(token, p) {
  assertAuth_(token);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    ensureSetup_();
    const name = String(p.name || "").trim();
    const direction = p.direction === "in" ? "in" : "out";
    if (!name) throw new Error("نام کانال الزامی است");
    const dup = rowsOf_(SHEETS_.CHANNELS).some(function (c) {
      return String(c.name) === name && String(c.direction) === direction;
    });
    if (dup) throw new Error("کانالی با این نام از قبل وجود دارد");
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS_.CHANNELS);
    sh.appendRow([nextId_(sh), name, direction, false, Date.now()]);
    return { ok: true };
  } finally { lock.releaseLock(); }
}

/** حذف کانال — کانال‌های پیش‌فرض و کانال‌های دارای گردش حذف نمی‌شوند */
function apiDeleteChannel(token, id) {
  assertAuth_(token);
  const rows = rowsOf_(SHEETS_.CHANNELS);
  const ch = rows.filter(function (c) { return Number(c.id) === Number(id); })[0];
  if (!ch) throw new Error("کانال یافت نشد");
  if (ch.isSystem === true || String(ch.isSystem) === "TRUE") throw new Error("کانال‌های پیش‌فرض سیستم قابل حذف نیستند");
  const used = rowsOf_(SHEETS_.TX).some(function (t) { return Number(t.channelId) === Number(id); });
  if (used) throw new Error("برای این کانال گردش ثبت شده و قابل حذف نیست");
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS_.CHANNELS);
  const lr = sh.getLastRow();
  const vals = sh.getRange(2, 1, lr - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (Number(vals[i][0]) === Number(id)) {
      sh.deleteRow(i + 2);
      return { ok: true };
    }
  }
  throw new Error("کانال یافت نشد");
}

/** ثبت گردش انبار (ورودی یا خروجی) با کنترل موجودی */
function apiAddTransaction(token, p) {
  assertAuth_(token);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    ensureSetup_();
    const bookId = Number(p.bookId);
    const channelId = Number(p.channelId);
    const qty = parseInt(p.quantity, 10);
    const direction = p.direction === "out" ? "out" : "in";
    const note = String(p.note || "").trim();

    if (!qty || qty < 1) throw new Error("تعداد باید عددی صحیح و حداقل ۱ باشد");

    const book = rowsOf_(SHEETS_.BOOKS).filter(function (b) { return Number(b.id) === bookId; })[0];
    if (!book) throw new Error("کتاب یافت نشد");
    const ch = rowsOf_(SHEETS_.CHANNELS).filter(function (c) { return Number(c.id) === channelId; })[0];
    if (!ch || String(ch.direction) !== direction) throw new Error("کانال انتخاب‌شده با نوع گردش هم‌خوانی ندارد");

    const s = stockMap_()[String(bookId)] || { totalIn: 0, totalOut: 0, stock: 0 };
    if (direction === "out" && qty > s.stock) {
      throw new Error("موجودی کافی نیست؛ موجودی فعلی «" + book.title + "» فقط " + s.stock + " عدد است");
    }

    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS_.TX);
    sh.appendRow([nextId_(sh), Date.now(), bookId, String(book.barcode), String(book.title), direction, channelId, String(ch.name), qty, note]);

    const newStock = direction === "in" ? s.stock + qty : s.stock - qty;
    return { ok: true, newStock: newStock, bookTitle: String(book.title) };
  } finally { lock.releaseLock(); }
}

/** سوابق گردش — آخرین‌ها اول */
function apiListTransactions(token, limit) {
  assertAuth_(token);
  const lim = Math.min(Number(limit) || 300, 1000);
  return rowsOf_(SHEETS_.TX).slice(-lim).reverse().map(txView_);
}

/* --------------------- ورود گروهی از اکسل --------------------- */

/** تبدیل ارقام فارسی/عربی به لاتین */
function normDigits_(s) {
  const FA = "۰۱۲۳۴۵۶۷۸۹", AR = "٠١٢٣٤٥٦٧٨٩";
  return String(s)
    .replace(/[۰-۹]/g, function (d) { return String(FA.indexOf(d)); })
    .replace(/[٠-٩]/g, function (d) { return String(AR.indexOf(d)); });
}

/** سلول (متن/عدد/تاریخ) را به رشته تمیز تبدیل می‌کند */
function cell_(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "number" && isFinite(v)) {
    return normDigits_(String(Math.round(v)));
  }
  return normDigits_(String(v)).trim();
}

/**
 * پردازش جدول خام (آرایه‌ای از آرایه‌ها شامل ردیف سرستون)
 * خروجی: {inserted, duplicates, stockTx, failed:[{row,reason}]}
 */
function processSheetValues_(values) {
  if (!values || values.length < 2) throw new Error("فایل خالی است؛ حداقل یک ردیف کتاب لازم است");
  if (values.length > IMPORT_MAX_ROWS_ + 1) throw new Error("حداکثر " + IMPORT_MAX_ROWS_ + " ردیف در هر بار قابل ورود است");

  const clean = function (s) { return String(s || "").replace(/[\s‌_]+/g, ""); };
  const header = values[0].map(clean);
  const colOf = function (names) {
    for (let i = 0; i < header.length; i++) {
      for (let j = 0; j < names.length; j++) {
        if (header[i] === clean(names[j])) return i;
      }
    }
    return -1;
  };

  const iBar = colOf(["بارکد", "ISBN", "شناسه"]);
  const iTitle = colOf(["عنوان", "عنوان کتاب", "نام کتاب"]);
  const iAuth = colOf(["نویسنده", "مولف"]);
  const iPub = colOf(["ناشر"]);
  const iDkpc = colOf(["کد تنوع دیجی‌کالا", "کد تنوع", "DKPC"]);
  const iQty = colOf(["موجودی اولیه", "موجودی", "تعداد"]);

  if (iBar < 0 || iTitle < 0) {
    throw new Error("ستون‌های «بارکد» و «عنوان» پیدا نشدند؛ از فایل اکسل نمونه استفاده کنید");
  }

  // ردیف‌های خالی کامل را حذف کن
  const body = [];
  values.slice(1).forEach(function (r, i) {
    const empty = r.every(function (c) { return String(c).trim() === ""; });
    if (!empty) body.push({ row: i + 2, cells: r });
  });
  if (body.length === 0) throw new Error("فایل خالی است؛ حداقل یک ردیف کتاب لازم است");

  const failed = [];
  const parsed = [];
  const seen = {};

  body.forEach(function (item) {
    const r = item.cells;
    const barcode = cell_(r[iBar]);
    const title = iTitle >= 0 ? String(r[iTitle] || "").trim() : "";
    const author = iAuth >= 0 ? String(r[iAuth] || "").trim() : "";
    const publisher = iPub >= 0 ? String(r[iPub] || "").trim() : "";
    const dkpc = iDkpc >= 0 ? cell_(r[iDkpc]) : "";
    let qty = 0;
    if (iQty >= 0 && String(r[iQty]).trim() !== "") {
      qty = Math.floor(Number(normDigits_(String(r[iQty]))));
    }
    if (!barcode || !title) {
      failed.push({ row: item.row, reason: !barcode && !title ? "بارکد و عنوان خالی است" : (!barcode ? "بارکد خالی است" : "عنوان خالی است") });
      return;
    }
    if (qty < 0 || isNaN(qty)) {
      failed.push({ row: item.row, reason: "موجودی اولیه نامعتبر است" });
      return;
    }
    if (seen[barcode]) {
      failed.push({ row: item.row, reason: "بارکد تکراری داخل خود فایل" });
      return;
    }
    seen[barcode] = true;
    parsed.push({ barcode: barcode, title: title, author: author, publisher: publisher, dkpc: dkpc, qty: qty });
  });

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    ensureSetup_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const shBooks = ss.getSheetByName(SHEETS_.BOOKS);

    // بارکدهای موجود
    const existing = {};
    rowsOf_(SHEETS_.BOOKS).forEach(function (b) { existing[String(b.barcode)] = true; });
    const toInsert = parsed.filter(function (p) { return !existing[p.barcode]; });
    const duplicates = parsed.length - toInsert.length;

    let stockTx = 0;

    if (toInsert.length) {
      // درج دسته‌ای کتاب‌ها
      let id = nextId_(shBooks);
      const now = Date.now();
      const newRows = toInsert.map(function (p) {
        const row = [id, p.barcode, p.title, p.author, p.publisher, p.dkpc, now];
        p.id = id;
        id++;
        return row;
      });
      shBooks.getRange(shBooks.getLastRow() + 1, 1, newRows.length, 7).setValues(newRows);

      // موجودی اولیه به‌عنوان گردش ورودی «خرید»
      const withQty = toInsert.filter(function (p) { return p.qty > 0; });
      if (withQty.length) {
        const shCh = ss.getSheetByName(SHEETS_.CHANNELS);
        let buy = rowsOf_(SHEETS_.CHANNELS).filter(function (c) { return c.direction === "in" && String(c.name) === "خرید"; })[0];
        if (!buy) {
          shCh.appendRow([nextId_(shCh), "خرید", "in", true, now]);
          buy = { id: nextId_(shCh) - 1, name: "خرید" };
        }
        const shTx = ss.getSheetByName(SHEETS_.TX);
        let txId = nextId_(shTx);
        const txRows = withQty.map(function (p) {
          return [txId++, now, p.id, p.barcode, p.title, "in", Number(buy.id), "خرید", p.qty, "موجودی اولیه (ورود از اکسل)"];
        });
        shTx.getRange(shTx.getLastRow() + 1, 1, txRows.length, 10).setValues(txRows);
        stockTx = txRows.length;
      }
    }

    return { inserted: toInsert.length, duplicates: duplicates, stockTx: stockTx, failed: failed };
  } finally { lock.releaseLock(); }
}

/** ورود گروهی با ردیف‌هایی که از CSV در مرورگر خوانده شده‌اند */
function apiImportBooks(token, rows) {
  assertAuth_(token);
  if (!Array.isArray(rows) || !rows.length) throw new Error("ردیفی از فایل خوانده نشد");
  return processSheetValues_(rows);
}

/** ورود گروهی از فایل اکسل (تبدیل موقت به گوگل‌شیت با سرویس Drive) */
function apiImportExcel(token, file) {
  assertAuth_(token);
  if (!file || !file.base64) throw new Error("فایل دریافت نشد");
  const bytes = Utilities.base64Decode(file.base64);
  if (bytes.length > 5 * 1024 * 1024) throw new Error("حجم فایل بیش از ۵ مگابایت است");

  const blob = Utilities.newBlob(
    bytes,
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    String(file.name || "import.xlsx")
  );

  let tmpId = null;
  try {
    const created = Drive.Files.create({ name: "excel-import-" + Date.now(), mimeType: MimeType.GOOGLE_SHEETS }, blob);
    tmpId = created.id;
    const sh = SpreadsheetApp.openById(tmpId).getSheets()[0];
    const values = sh.getDataRange().getValues();
    return processSheetValues_(values);
  } catch (e) {
    const msg = String((e && e.message) || e);
    if (msg.indexOf("بارکد") >= 0 || msg.indexOf("فایل خالی") >= 0 || msg.indexOf("ردیف") >= 0) throw e;
    throw new Error("خواندن فایل اکسل ممکن نشد. مطمئن شوید سرویس پیشرفته Drive فعال است (راهنما را ببینید) یا فایل را به‌صورت CSV ذخیره و بارگذاری کنید");
  } finally {
    if (tmpId) { try { Drive.Files.remove(tmpId); } catch (_) {} }
  }
}

/** ساخت و دانلود فایل اکسل نمونه — ردیف‌های مثال برای پرکردن */
function apiGetSampleFile(token) {
  assertAuth_(token);
  const headers = ["بارکد", "عنوان", "نویسنده", "ناشر", "کد تنوع دیجی‌کالا", "موجودی اولیه"];
  const rows = [
    ["9789643694406", "کیمیاگر", "پائولو کوئیلو", "نشر کاروان", "1045321", 12],
    ["9789640013325", "صد سال تنهایی", "گابریل گارسیا مارکز", "امیرکبیر", "", 8],
    ["9786227443042", "۱۹۸۴", "جورج اورول", "نشر جامی", "", 20]
  ];

  const tmp = SpreadsheetApp.create("books-sample-tmp");
  let bytes;
  try {
    const sh = tmp.getSheets()[0];
    sh.setName("کتاب‌ها");
    sh.getRange(1, 1, 1, 6).setValues([headers]).setFontWeight("bold").setBackground("#dcefe4");
    // بارکد به‌صورت متن ذخیره شود تا اکسل آن را علمی نکند
    for (let i = 0; i < rows.length; i++) {
      sh.getRange(i + 2, 1).setNumberFormat("@").setValue(rows[i][0]);
      sh.getRange(i + 2, 2, 1, 5).setValues([rows[i].slice(1)]);
    }
    [18, 30, 24, 20, 18, 12].forEach(function (w, i) { sh.setColumnWidth(i + 1, w * 7); });
    SpreadsheetApp.flush();

    const url = "https://docs.google.com/spreadsheets/d/" + tmp.getId() + "/export?format=xlsx";
    const resp = UrlFetchApp.fetch(url, {
      headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) throw new Error("export failed");
    bytes = resp.getBlob().getBytes();
  } finally {
    try { DriveApp.getFileById(tmp.getId()).setTrashed(true); } catch (_) {}
  }

  return {
    name: "books-sample.xlsx",
    mime: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    base64: Utilities.base64Encode(bytes)
  };
}
