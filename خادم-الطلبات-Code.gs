/**
 * خادم طلبات الكتالوج — إدارة أعمالي
 * يحفظ طلبات الزبائن فور إرسالها في Google Sheets، ويقدّم الكتالوج المنشور من التطبيق.
 * تصميم وبرمجة: م. ماجد الرميمي
 * لا تغيّر السطر التالي: مفتاح المالك الذي يربط هذا الخادم بتطبيقك.
 */
const OWNER_KEY = "__KEY__";
const SHEET_NAME = "الطلبات";

function _json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function _props() { return PropertiesService.getScriptProperties(); }
function _r3(n) { return Math.round((+n || 0) * 1000) / 1000; }
function _clean(s, n) { return String(s == null ? "" : s).replace(/[\u0000-\u001f\u007f]/g, " ").trim().slice(0, n || 200); }

function _catFile() {
  const id = _props().getProperty("CAT_FILE");
  if (id) { try { return DriveApp.getFileById(id); } catch (e) {} }
  const f = DriveApp.createFile("idarat-catalog.json", "{}", MimeType.PLAIN_TEXT);
  _props().setProperty("CAT_FILE", f.getId());
  return f;
}
function _catalog() { try { return JSON.parse(_catFile().getBlob().getDataAsString() || "{}"); } catch (e) { return {}; } }

function _sheet() {
  const pr = _props(); let ss = null; const ssId = pr.getProperty("SS");
  if (ssId) { try { ss = SpreadsheetApp.openById(ssId); } catch (e) {} }
  if (!ss) { ss = SpreadsheetApp.create("طلبات الكتالوج — إدارة أعمالي"); pr.setProperty("SS", ss.getId()); }
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(["رقم الطلب", "التاريخ", "المحل", "المستلم", "الهاتف", "المنطقة", "العنوان", "الدفع", "المجموع", "التوصيل", "الإجمالي", "الأصناف", "ملاحظات", "id", "json"]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    if (p.action === "catalog") return _json({ ok: true, catalog: _catalog() });
    if (p.action === "ping") return _json({ ok: true, v: 1, auth: p.key === OWNER_KEY, hasCatalog: !!(_catalog().products) });
    if (p.action === "orders") {
      if (p.key !== OWNER_KEY) return _json({ ok: false, error: "auth" });
      const since = +p.since || 0; const sh = _sheet(); const n = sh.getLastRow();
      if (n < 2) return _json({ ok: true, orders: [] });
      const from = Math.max(2, n - 1500 + 1);
      const rows = sh.getRange(from, 15, n - from + 1, 1).getValues(); const out = [];
      rows.forEach(function (r) { try { const o = JSON.parse(r[0]); if (o.ts > since) out.push(o); } catch (x) {} });
      return _json({ ok: true, orders: out });
    }
    return _json({ ok: false, error: "action" });
  } catch (err) { return _json({ ok: false, error: String(err) }); }
}

function doPost(e) {
  let b = {};
  try { b = JSON.parse((e && e.postData && e.postData.contents) || "{}"); } catch (x) { return _json({ ok: false, error: "json" }); }
  try {
    if (b.action === "publish") {
      if (b.key !== OWNER_KEY) return _json({ ok: false, error: "auth" });
      const s = JSON.stringify(b.catalog || {});
      if (s.length > 8000000) return _json({ ok: false, error: "size" });
      _catFile().setContent(s);
      return _json({ ok: true, at: Date.now() });
    }
    if (b.action === "order") return _order(b.order || {});
    return _json({ ok: false, error: "action" });
  } catch (err) { return _json({ ok: false, error: String(err) }); }
}

function _order(o) {
  const lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    const cid = _clean(o.cid, 60); const cache = CacheService.getScriptCache();
    if (cid) { const prev = cache.get("cid_" + cid); if (prev) return ContentService.createTextOutput(prev).setMimeType(ContentService.MimeType.JSON); }
    const cat = _catalog();
    if (!cat.products) return _json({ ok: false, error: "nocatalog" });
    const name = _clean(o.name, 80), phone = _clean(o.phone, 30).replace(/[^\d+]/g, "").slice(0, 20);
    if (!name || phone.replace(/\D/g, "").length < 7) return _json({ ok: false, error: "contact" });
    const byId = {}; cat.products.forEach(function (p) { byId[p.id] = p; });
    const items = []; const skipped = [];
    (o.items || []).slice(0, 200).forEach(function (it) {
      const p = byId[it.id]; const q = Math.floor(+it.qty || 0);
      if (q < 1 || q > 100000) return;
      if (!p || p.out) { skipped.push(_clean(p ? p.name : it.id, 80)); return; }
      const piece = it.mode === "piece" && +p.piecePrice > 0;
      items.push({ id: p.id, name: p.name, unit: piece ? (p.pieceUnit || "حبة") : (p.unit || ""), mode: piece ? "piece" : "pack", qty: q, price: piece ? +p.piecePrice : +p.price, f: piece ? (1 / (+p.packQty || 1)) : 1 });
    });
    if (!items.length) return _json({ ok: false, error: "items", skipped: skipped });
    const sub = _r3(items.reduce(function (a, i) { return a + i.qty * i.price; }, 0));
    if (+cat.minOrder && sub < +cat.minOrder) return _json({ ok: false, error: "min", min: +cat.minOrder });
    let fee = 0; const area = _clean(o.area, 60);
    if (cat.areas && cat.areas.length) {
      const a = cat.areas.filter(function (x) { return x.name === area; })[0];
      if (!a && cat.areasOnly) return _json({ ok: false, error: "area" });
      if (a) { fee = +a.fee || 0; if (+a.freeOver && sub >= +a.freeOver) fee = 0; }
    }
    const pays = cat.pay && cat.pay.length ? cat.pay : ["cash"];
    const pay = pays.indexOf(o.pay) >= 0 ? o.pay : pays[0];
    const now = new Date(); const day = Utilities.formatDate(now, cat.tz || "Asia/Muscat", "yyMMdd");
    const pr = _props(); const cnt = (+pr.getProperty("C_" + day) || 0) + 1; pr.setProperty("C_" + day, String(cnt));
    const ord = { id: Utilities.getUuid(), number: "ORD-" + day + "-" + ("00" + cnt).slice(-3), ts: now.getTime(), biz: cat.biz || "",
      shop: _clean(o.shop, 80), name: name, phone: phone, area: area, address: _clean(o.address, 200), notes: _clean(o.notes, 500), pay: pay,
      items: items, skipped: skipped, subtotal: sub, fee: _r3(fee), total: _r3(sub + fee), lang: o.lang === "en" ? "en" : "ar" };
    _sheet().appendRow([ord.number, now, ord.shop, ord.name, "'" + ord.phone, ord.area, ord.address, ord.pay, ord.subtotal, ord.fee, ord.total,
      items.map(function (i) { return i.name + " × " + i.qty + " " + i.unit; }).join("\n"), ord.notes, ord.id, JSON.stringify(ord)]);
    const res = JSON.stringify({ ok: true, number: ord.number, subtotal: sub, fee: ord.fee, total: ord.total, items: items, skipped: skipped, ts: ord.ts });
    if (cid) cache.put("cid_" + cid, res, 21600);
    if (cat.notify) _notify(cat, ord);
    return ContentService.createTextOutput(res).setMimeType(ContentService.MimeType.JSON);
  } finally { lock.releaseLock(); }
}

function _notify(cat, o) {
  try {
    const to = cat.notifyTo || Session.getEffectiveUser().getEmail(); if (!to) return;
    const L = ["طلب جديد من الكتالوج: " + o.number, "", (o.shop ? "المحل: " + o.shop + "\n" : "") + "المستلم: " + o.name, "الهاتف: " + o.phone];
    if (o.area) L.push("المنطقة: " + o.area); if (o.address) L.push("العنوان: " + o.address);
    L.push("", o.items.map(function (i, n) { return (n + 1) + ") " + i.name + " × " + i.qty + " " + i.unit + " = " + _r3(i.qty * i.price).toFixed(3); }).join("\n"), "");
    if (o.fee) L.push("التوصيل: " + o.fee.toFixed(3)); L.push("الإجمالي: " + o.total.toFixed(3) + " ر.ع", "الدفع: " + o.pay);
    if (o.notes) L.push("ملاحظات: " + o.notes); if (o.skipped.length) L.push("أصناف لم تُضف لنفاد المخزون: " + o.skipped.join("، "));
    L.push("", "افتح «إدارة أعمالي» ← المبيعات ← طلبات الكتالوج لتأكيده.");
    MailApp.sendEmail(to, "🛒 طلب جديد " + o.number + " — " + (o.shop || o.name) + " — " + o.total.toFixed(3) + " ر.ع", L.join("\n"));
  } catch (e) {}
}
